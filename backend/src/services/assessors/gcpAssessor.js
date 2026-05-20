'use strict';

const fs = require('fs');
const path = require('path');
const { detectGcpColumns } = require('../utils/typeDetector');

/**
 * GCP-to-Azure Migration Assessor
 * Self-contained module for mapping GCP resources to Azure equivalents.
 * Errors in this module will NOT crash other assessment modes.
 */
class GcpAssessor {
  constructor() {
    this._mappings = null;
    this._displayNameLookup = null;
    this._metadata = null;
  }

  /**
   * Load GCP-to-Azure mapping data from JSON.
   */
  _loadMapping() {
    if (this._mappings) return;
    const mappingPath = path.join(__dirname, '..', '..', 'data', 'gcp-azure-mapping.json');
    const data = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'));
    this._mappings = data.mappings;
    this._displayNameLookup = data.displayNameLookup;
    this._metadata = data._metadata;
  }

  /**
   * Force reload mapping data (useful after updates to the JSON file).
   */
  reloadMapping() {
    this._mappings = null;
    this._displayNameLookup = null;
    this._metadata = null;
    this._loadMapping();
  }

  /**
   * Normalize a GCP resource type to canonical format (gcp::service::resource).
   */
  normalizeType(rawType) {
    if (!rawType) return '';
    let t = String(rawType).trim().toLowerCase();

    // Already in canonical format (gcp::...)
    if (t.startsWith('gcp::')) return t;

    // GCP resource name format: projects/*/service/resource or //service.googleapis.com/projects/...
    if (t.includes('googleapis.com/')) {
      const match = t.match(/\/\/([a-z]+)\.googleapis\.com\//);
      if (match) {
        const service = match[1];
        const serviceLookup = {
          'compute': 'gcp::compute::instance',
          'storage': 'gcp::storage::bucket',
          'cloudfunctions': 'gcp::functions::function',
          'run': 'gcp::run::service',
          'container': 'gcp::container::cluster',
          'sqladmin': 'gcp::sql::instance',
          'spanner': 'gcp::spanner::instance',
          'bigtable': 'gcp::bigtable::instance',
          'firestore': 'gcp::firestore::database',
          'redis': 'gcp::memorystore::instance',
          'bigquery': 'gcp::bigquery::dataset',
          'dataflow': 'gcp::dataflow::job',
          'dataproc': 'gcp::dataproc::cluster',
          'pubsub': 'gcp::pubsub::topic',
          'dns': 'gcp::dns::managedzone',
          'cloudkms': 'gcp::kms::keyring',
          'secretmanager': 'gcp::secretmanager::secret',
          'logging': 'gcp::logging::logbucket',
          'monitoring': 'gcp::monitoring::alertpolicy',
          'cloudbuild': 'gcp::cloudbuild::trigger',
          'iam': 'gcp::iam::serviceaccount',
          'artifactregistry': 'gcp::artifactregistry::repository'
        };
        if (serviceLookup[service]) return serviceLookup[service];
      }
    }

    // Try display name lookup (e.g. "Cloud Run", "BigQuery", "Compute Engine")
    if (this._displayNameLookup && this._displayNameLookup[t]) {
      return this._displayNameLookup[t];
    }

    // Try partial match in display name lookup
    for (const [name, canonicalType] of Object.entries(this._displayNameLookup || {})) {
      if (t.includes(name) || name.includes(t)) {
        return canonicalType;
      }
    }

    return t;
  }

  /**
   * Assess Azure equivalent for a given GCP resource type.
   */
  assessType(rawType, skuOrSize) {
    this._loadMapping();
    const normalized = this.normalizeType(rawType);

    // Direct match
    if (this._mappings[normalized]) {
      const m = this._mappings[normalized];
      const result = {
        azureService: m.azureService,
        azureResourceType: m.azureResourceType,
        category: m.category,
        similarity: m.similarity,
        description: m.description,
        migrationNotes: m.migrationNotes,
        skuRecommendation: m.defaultSku || ''
      };
      // SKU recommendation
      if (skuOrSize && m.skuMapping) {
        const skuLower = String(skuOrSize).toLowerCase().trim();
        // Exact match
        const skuMatch = Object.entries(m.skuMapping).find(([k]) => k.toLowerCase() === skuLower);
        if (skuMatch) {
          result.skuRecommendation = skuMatch[1];
        } else {
          // Family-based match for GCE: extract family prefix (e.g., "n2-standard" from "n2-standard-4")
          const familyMatch = skuLower.match(/^([a-z]\d+[a-z]?-[a-z]+)/);
          if (familyMatch) {
            const family = familyMatch[1];
            let familyEntries = Object.entries(m.skuMapping)
              .filter(([k]) => k.toLowerCase().startsWith(family + '-'));
            if (familyEntries.length > 0) {
              // Find closest size by vCPU count
              const targetVcpus = parseInt(skuLower.split('-').pop()) || 0;
              let best = familyEntries[familyEntries.length - 1];
              if (targetVcpus > 0) {
                for (const entry of familyEntries) {
                  const entryVcpus = parseInt(entry[0].split('-').pop()) || 0;
                  if (entryVcpus >= targetVcpus) { best = entry; break; }
                }
              }
              result.skuRecommendation = best[1] + ' (approx)';
            }
          }
        }
      }
      return result;
    }

    // Try parent type (e.g. gcp::compute::instance/something → gcp::compute::instance)
    const parts = normalized.split('::');
    if (parts.length > 3) {
      const parentType = parts.slice(0, 3).join('::');
      if (this._mappings[parentType]) {
        const m = this._mappings[parentType];
        return {
          azureService: m.azureService,
          azureResourceType: m.azureResourceType,
          category: m.category,
          similarity: m.similarity,
          description: m.description + ' (Matched via parent type)',
          migrationNotes: m.migrationNotes,
          skuRecommendation: ''
        };
      }
    }

    // Try service-level match (e.g. gcp::compute::*)
    if (parts.length >= 2) {
      const servicePrefix = parts.slice(0, 2).join('::') + '::';
      const serviceMatch = Object.entries(this._mappings).find(([key]) => key.startsWith(servicePrefix));
      if (serviceMatch) {
        return {
          azureService: 'Review Required',
          azureResourceType: serviceMatch[1].azureResourceType,
          category: serviceMatch[1].category,
          similarity: 'No Direct Mapping',
          description: `No specific Azure mapping found for ${rawType}. Related services in the same GCP namespace exist.`,
          migrationNotes: `Review ${serviceMatch[1].azureService} and related Azure services for potential alternatives.`,
          skuRecommendation: ''
        };
      }
    }

    return {
      azureService: 'No Azure Equivalent Found',
      azureResourceType: 'N/A',
      category: 'Unknown',
      similarity: 'No Direct Mapping',
      description: `No Azure equivalent identified for ${rawType}. This may be a GCP-specific service or a resource type not yet mapped.`,
      migrationNotes: 'Research Azure services manually or consult https://learn.microsoft.com/en-us/azure/architecture/gcp-professional/services',
      skuRecommendation: ''
    };
  }

  /**
   * Extract machine type from a usage string.
   * E.g., "Compute Engine: n2-standard-4 in us-central1" → "n2-standard-4"
   */
  _extractMachineType(usage) {
    if (!usage) return '';
    const str = String(usage);
    // Pattern: machine type names like n2-standard-4, e2-micro, c2-standard-30
    const match = str.match(/\b([a-z]\d+[a-z]?-(?:standard|highmem|highcpu|megamem|ultramem|highgpu)-\d+[a-z]?)\b/i);
    if (match) return match[1];
    // Also handle custom machine types: custom-4-16384
    const customMatch = str.match(/\b(custom-\d+-\d+)\b/i);
    if (customMatch) return customMatch[1];
    // Also handle predefined types: e2-micro, e2-small, e2-medium
    const predefMatch = str.match(/\b([a-z]\d+[a-z]?-(?:micro|small|medium))\b/i);
    if (predefMatch) return predefMatch[1];
    return '';
  }

  /**
   * Assess an array of resource rows from an uploaded file.
   */
  assessResources(resources) {
    this._loadMapping();
    if (!resources || resources.length === 0) return [];

    const firstRow = resources[0];
    const { serviceCol, nameCol, skuCol } = detectGcpColumns(firstRow);

    return resources.map(row => {
      const rawService = serviceCol ? String(row[serviceCol] || '') : '';
      const rawName = nameCol ? String(row[nameCol] || '') : '';
      let rawSku = skuCol ? String(row[skuCol] || '') : '';

      // Try to extract machine type from SKU/usage column
      if (rawSku && !rawSku.match(/^[a-z]\d/i)) {
        const extracted = this._extractMachineType(rawSku);
        if (extracted) rawSku = extracted;
      }

      const assessment = this.assessType(rawService, rawSku);

      return {
        'GCP SERVICE': rawService || 'Unknown',
        'NAME': rawName || '',
        'AZURE EQUIVALENT': assessment.azureService,
        'AZURE RESOURCE TYPE': assessment.azureResourceType,
        'CATEGORY': assessment.category,
        'SIMILARITY': assessment.similarity,
        'SKU RECOMMENDATION': assessment.skuRecommendation,
        'DESCRIPTION': assessment.description,
        'MIGRATION NOTES': assessment.migrationNotes,
        ...row
      };
    });
  }

  /**
   * Get summary statistics from assessed resources.
   */
  getSummary(assessed) {
    const total = assessed.length;
    let directEquivalent = 0;
    let similar = 0;
    let partial = 0;
    let noMapping = 0;
    const categories = {};

    for (const item of assessed) {
      const sim = (item['SIMILARITY'] || '').toLowerCase();
      if (sim === 'direct equivalent') directEquivalent++;
      else if (sim === 'similar') similar++;
      else if (sim === 'partial') partial++;
      else noMapping++;

      const cat = item['CATEGORY'] || 'Unknown';
      categories[cat] = (categories[cat] || 0) + 1;
    }

    return { total, directEquivalent, similar, partial, noMapping, categories };
  }
}

module.exports = new GcpAssessor();
