'use strict';

const fs = require('fs');
const path = require('path');
const { detectAwsColumns } = require('../utils/typeDetector');

/**
 * AWS-to-Azure Migration Assessor
 * Self-contained module for mapping AWS resources to Azure equivalents.
 * Errors in this module will NOT crash other assessment modes.
 */
class AwsAssessor {
  constructor() {
    this._mappings = null;
    this._displayNameLookup = null;
    this._metadata = null;
  }

  /**
   * Load AWS-to-Azure mapping data from JSON.
   */
  _loadMapping() {
    if (this._mappings) return;
    const mappingPath = path.join(__dirname, '..', '..', 'data', 'aws-azure-mapping.json');
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
   * Normalize an AWS resource type to CloudFormation format (aws::service::resource).
   */
  normalizeType(rawType) {
    if (!rawType) return '';
    let t = String(rawType).trim().toLowerCase();

    // Already in CloudFormation format (aws::...)
    if (t.startsWith('aws::')) return t;

    // ARN format: arn:aws:service:region:account:resource
    if (t.startsWith('arn:aws:')) {
      const parts = t.split(':');
      if (parts.length >= 3) {
        const service = parts[2];
        const arnLookup = {
          'ec2': 'aws::ec2::instance',
          's3': 'aws::s3::bucket',
          'lambda': 'aws::lambda::function',
          'rds': 'aws::rds::dbinstance',
          'dynamodb': 'aws::dynamodb::table',
          'ecs': 'aws::ecs::service',
          'eks': 'aws::eks::cluster',
          'elasticache': 'aws::elasticache::cachecluster',
          'sqs': 'aws::sqs::queue',
          'sns': 'aws::sns::topic',
          'kinesis': 'aws::kinesis::stream',
          'logs': 'aws::logs::loggroup',
          'events': 'aws::events::rule',
          'states': 'aws::stepfunctions::statemachine',
          'kms': 'aws::kms::key',
          'secretsmanager': 'aws::secretsmanager::secret'
        };
        if (arnLookup[service]) return arnLookup[service];
      }
    }

    // Try display name lookup (e.g. "Amazon EC2", "Lambda", "S3", "AmazonEC2")
    if (this._displayNameLookup && this._displayNameLookup[t]) {
      return this._displayNameLookup[t];
    }

    // Try partial match in display name lookup
    for (const [name, cfnType] of Object.entries(this._displayNameLookup || {})) {
      if (t.includes(name) || name.includes(t)) {
        return cfnType;
      }
    }

    return t;
  }

  /**
   * Assess Azure equivalent for a given AWS resource type.
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
      // SKU recommendation: try specific instance mapping first, then family match, then defaultSku
      if (skuOrSize && m.skuMapping) {
        const skuLower = String(skuOrSize).toLowerCase().trim();
        // Exact match
        const skuMatch = Object.entries(m.skuMapping).find(([k]) => k.toLowerCase() === skuLower);
        if (skuMatch) {
          result.skuRecommendation = skuMatch[1];
        } else {
          // Family-based match: extract family prefix (e.g., "m5" from "m5.24xlarge", "c5n" from "c5n.4xlarge")
          const familyMatch = skuLower.match(/^((?:db\.|cache\.)?[a-z]\d+[a-z]*)\./);
          if (familyMatch) {
            const family = familyMatch[1];
            // Try exact family first, then base family (e.g., c5n → c5, m5ad → m5a → m5)
            let familyEntries = Object.entries(m.skuMapping)
              .filter(([k]) => k.toLowerCase().startsWith(family + '.'));
            if (familyEntries.length === 0) {
              // Strip trailing letters to find base family: c5n→c5, m5ad→m5a→m5, r5b→r5
              const baseMatch = family.match(/^((?:db\.|cache\.)?[a-z]\d+)/);
              if (baseMatch) {
                familyEntries = Object.entries(m.skuMapping)
                  .filter(([k]) => k.toLowerCase().startsWith(baseMatch[1] + '.') || k.toLowerCase().startsWith(baseMatch[1] + 'a.'));
              }
            }
            familyEntries.sort((a, b) => {
              const sizeOrder = ['nano','micro','small','medium','large','xlarge','2xlarge','4xlarge','8xlarge','12xlarge','16xlarge','24xlarge','32xlarge','48xlarge','64xlarge','96xlarge','metal'];
              const aSize = a[0].split('.').pop().toLowerCase();
              const bSize = b[0].split('.').pop().toLowerCase();
              return sizeOrder.indexOf(aSize) - sizeOrder.indexOf(bSize);
            });
            if (familyEntries.length > 0) {
              // Find closest size or use the largest mapped size
              const sizeOrder = ['nano','micro','small','medium','large','xlarge','2xlarge','4xlarge','8xlarge','12xlarge','16xlarge','24xlarge','32xlarge','48xlarge','64xlarge','96xlarge','metal'];
              const targetSize = skuLower.split('.').pop();
              const targetIdx = sizeOrder.indexOf(targetSize);
              let best = familyEntries[familyEntries.length - 1]; // default to largest
              if (targetIdx >= 0) {
                for (const entry of familyEntries) {
                  const entrySize = entry[0].toLowerCase().split('.').pop();
                  const entryIdx = sizeOrder.indexOf(entrySize);
                  if (entryIdx >= targetIdx) { best = entry; break; }
                }
              }
              result.skuRecommendation = best[1] + ' (approx)';
            }
          }
        }
      }
      return result;
    }

    // Try parent type (e.g. aws::ec2::instance/something → aws::ec2::instance)
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

    // Try service-level match (e.g. aws::ec2::*)
    if (parts.length >= 2) {
      const servicePrefix = parts.slice(0, 2).join('::') + '::';
      const serviceMatch = Object.entries(this._mappings).find(([key]) => key.startsWith(servicePrefix));
      if (serviceMatch) {
        return {
          azureService: 'Review Required',
          azureResourceType: serviceMatch[1].azureResourceType,
          category: serviceMatch[1].category,
          similarity: 'No Direct Mapping',
          description: `No specific Azure mapping found for ${rawType}. Related services in the same AWS namespace exist.`,
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
      description: `No Azure equivalent identified for ${rawType}. This may be an AWS-specific service or a resource type not yet mapped.`,
      migrationNotes: 'Research Azure services manually or consult https://learn.microsoft.com/en-us/azure/architecture/aws-professional/',
      skuRecommendation: ''
    };
  }

  /**
   * Extract instance type from a UsageType string.
   * E.g., "APS3-BoxUsage:c6a.2xlarge" → "c6a.2xlarge"
   * E.g., "USE1-InstanceUsage:db.t4g.medium" → "db.t4g.medium"
   */
  _extractInstanceFromUsageType(usageType) {
    if (!usageType) return '';
    const str = String(usageType);
    // Patterns like "APS3-BoxUsage:c6a.2xlarge" or "USE1-InstanceUsage:db.t4g.medium"
    const match = str.match(/(?:BoxUsage|InstanceUsage|NodeUsage|CacheUsage|Multi-AZUsage):(.+)$/i);
    if (match) return match[1];
    // RDS-style: "db.t4g.medium", "db.r5.large"
    const rdsMatch = str.match(/\b(db\.[a-z]\d[a-z]?\.\w+)\b/i);
    if (rdsMatch) return rdsMatch[1];
    // Cache-style: "cache.t3.micro"
    const cacheMatch = str.match(/\b(cache\.[a-z]\d[a-z]?\.\w+)\b/i);
    if (cacheMatch) return cacheMatch[1];
    // Generic instance types: "c6a.2xlarge", "m5.large", "t3.micro"
    const instanceMatch = str.match(/\b([a-z]\d[a-z]?\.\w+)\b/i);
    if (instanceMatch) return instanceMatch[1];
    return '';
  }

  /**
   * Check if a raw value looks like an actual instance/SKU identifier.
   */
  _looksLikeSku(value) {
    if (!value) return false;
    const str = String(value).trim();
    // Instance types: t3.micro, m5.large, db.t4g.medium, Standard_D2s_v3, etc.
    if (/^(db\.)?[a-z]\d[a-z]?\.\w+$/i.test(str)) return true;
    // Azure-style SKU: Standard_F8s_v2, Basic_A1, etc.
    if (/^(Standard|Basic|Premium)_/i.test(str)) return true;
    // Cache node types: cache.t3.micro
    if (/^cache\.[a-z]\d[a-z]?\.\w+$/i.test(str)) return true;
    return false;
  }

  /**
   * Assess Azure equivalents for an array of AWS resources.
   */
  assessResources(resources) {
    if (!resources || resources.length === 0) return [];

    const firstRow = resources[0];
    const { serviceCol, nameCol } = detectAwsColumns(firstRow);

    return resources.map(resource => {
      const serviceValue = serviceCol ? (resource[serviceCol] || '') : '';
      const nameValue = nameCol ? (resource[nameCol] || '') : '';
      const assessment = this.assessType(serviceValue);

      return {
        ...resource,
        'NAME': nameValue,
        'AWS SERVICE': serviceValue,
        'AZURE EQUIVALENT': assessment.azureService,
        'AZURE RESOURCE TYPE': assessment.azureResourceType,
        'CATEGORY': assessment.category,
        'SIMILARITY': assessment.similarity,
        'DESCRIPTION': assessment.description,
        'MIGRATION NOTES': assessment.migrationNotes
      };
    });
  }

  /**
   * Get summary for AWS-to-Azure assessment.
   */
  getSummary(assessedResources) {
    const total = assessedResources.length;
    const directEquivalent = assessedResources.filter(r => r['SIMILARITY'] === 'Direct Equivalent').length;
    const similar = assessedResources.filter(r => r['SIMILARITY'] === 'Similar').length;
    const partial = assessedResources.filter(r => r['SIMILARITY'] === 'Partial').length;
    const noMapping = assessedResources.filter(r => r['SIMILARITY'] === 'No Direct Mapping').length;

    return {
      total,
      directEquivalent,
      similar,
      partial,
      noMapping,
      yes: directEquivalent,
      no: noMapping,
      review: similar + partial
    };
  }
}

module.exports = new AwsAssessor();
