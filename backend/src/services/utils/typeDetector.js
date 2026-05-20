'use strict';

/**
 * Shared utility for detecting column names in uploaded resource data.
 * Used by all assessor modules to find the relevant columns (type, name, etc.)
 */

/**
 * Detect the column containing resource type information.
 */
function detectTypeColumn(row) {
  const candidates = [
    'RESOURCE TYPE', 'Resource Type', 'Resource type', 'resource type',
    'Resource_Type', 'resource_type', 'RESOURCE_TYPE',
    'ResourceType', 'resourceType', 'resourcetype',
    'Azure Resource Type', 'azure resource type',
    'Resource Provider/Type',
    'TYPE', 'Type', 'type',
  ];

  for (const col of candidates) {
    if (row[col] !== undefined && row[col] !== null && row[col] !== '') {
      return col;
    }
  }

  const keys = Object.keys(row);
  for (const key of keys) {
    if (/\btype\b/i.test(key)) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
        return key;
      }
    }
  }

  return null;
}

/**
 * Detect the column containing resource name.
 */
function detectNameColumn(row) {
  const candidates = [
    'NAME', 'Name', 'name',
    'Resource Name', 'RESOURCE NAME', 'resource name',
    'Resource_Name', 'resource_name',
    'ResourceName', 'resourceName',
    'Display Name', 'display name', 'DISPLAY NAME',
  ];

  for (const col of candidates) {
    if (row[col] !== undefined && row[col] !== null) {
      return col;
    }
  }

  const keys = Object.keys(row);
  for (const key of keys) {
    if (/\bname\b/i.test(key) && !/namespace/i.test(key)) {
      if (row[key] !== undefined && row[key] !== null) {
        return key;
      }
    }
  }

  return null;
}

/**
 * Detect the column containing location/region information.
 */
function detectLocationColumn(row) {
  const candidates = [
    'LOCATION', 'Location', 'location',
    'REGION', 'Region', 'region',
    'Azure Region', 'azure region',
  ];

  for (const col of candidates) {
    if (row[col] !== undefined && row[col] !== null && row[col] !== '') {
      return col;
    }
  }

  const keys = Object.keys(row);
  for (const key of keys) {
    if (/\b(location|region)\b/i.test(key)) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') {
        return key;
      }
    }
  }

  return null;
}

/**
 * Detect AWS-specific columns (for billing/CUR exports).
 * Returns { serviceCol, nameCol, skuCol }
 */
function detectAwsColumns(row, fallbacks = {}) {
  const keys = Object.keys(row);
  const keyLower = keys.map(k => ({ orig: k, lower: k.toLowerCase().replace(/[^a-z]/g, '') }));

  // AWS service column: ProductCode > ProductName > lineItem/ProductCode > Service
  let serviceCol = null;
  const serviceCandidates = ['productcode', 'productname', 'lineitemproductcode', 'service', 'servicename', 'awsservice'];
  for (const pattern of serviceCandidates) {
    const match = keyLower.find(k => k.lower === pattern || k.lower.includes(pattern));
    if (match && row[match.orig] !== undefined && row[match.orig] !== null && row[match.orig] !== '') {
      serviceCol = match.orig;
      break;
    }
  }
  if (!serviceCol) {
    serviceCol = fallbacks.typeCol || detectTypeColumn(row);
  }

  // Name column: ProductName > ItemDescription > Name > ResourceId (NOT PayerAccountName)
  let nameCol = null;
  const nameCandidates = ['productname', 'itemdescription', 'name', 'resourcename', 'resourceid', 'resource'];
  for (const pattern of nameCandidates) {
    const match = keyLower.find(k => k.lower === pattern || k.lower.includes(pattern));
    if (match && row[match.orig] !== undefined) {
      nameCol = match.orig;
      break;
    }
  }
  if (!nameCol) nameCol = fallbacks.nameCol || detectNameColumn(row);

  // SKU column: instance type > sku > size > UsageType
  let skuCol = null;
  const skuPatterns = ['instance type', 'instancetype', 'sku', 'size', 'instance_type', 'node type', 'nodetype', 'class'];
  for (const pattern of skuPatterns) {
    const match = keys.find(k => k.toLowerCase().includes(pattern));
    if (match) { skuCol = match; break; }
  }
  if (!skuCol) {
    const usageTypeCol = keyLower.find(k => k.lower === 'usagetype');
    if (usageTypeCol) skuCol = usageTypeCol.orig;
  }

  return { serviceCol, nameCol, skuCol };
}

/**
 * Detect GCP-specific columns (for billing exports / asset inventories).
 * Returns { serviceCol, nameCol, skuCol }
 */
function detectGcpColumns(row, fallbacks = {}) {
  const keys = Object.keys(row);
  const keyLower = keys.map(k => ({ orig: k, lower: k.toLowerCase().replace(/[^a-z]/g, '') }));

  // GCP service column: Service > ServiceDescription > ProductName > SKU Description
  let serviceCol = null;
  const serviceCandidates = ['service', 'servicedescription', 'serviceid', 'productname', 'gcpservice', 'skudescription'];
  for (const pattern of serviceCandidates) {
    const match = keyLower.find(k => k.lower === pattern || k.lower.includes(pattern));
    if (match && row[match.orig] !== undefined && row[match.orig] !== null && row[match.orig] !== '') {
      serviceCol = match.orig;
      break;
    }
  }
  if (!serviceCol) {
    serviceCol = fallbacks.typeCol || detectTypeColumn(row);
  }

  // Name column: Resource Name > Name > ResourceId > Project
  let nameCol = null;
  const nameCandidates = ['resourcename', 'name', 'resourceid', 'resource', 'project', 'projectname'];
  for (const pattern of nameCandidates) {
    const match = keyLower.find(k => k.lower === pattern || k.lower.includes(pattern));
    if (match && row[match.orig] !== undefined) {
      nameCol = match.orig;
      break;
    }
  }
  if (!nameCol) nameCol = fallbacks.nameCol || detectNameColumn(row);

  // SKU column: Machine Type > SKU > Instance Type > Size
  let skuCol = null;
  const skuPatterns = ['machine type', 'machinetype', 'machine_type', 'sku', 'instance type', 'size', 'tier'];
  for (const pattern of skuPatterns) {
    const match = keys.find(k => k.toLowerCase().includes(pattern));
    if (match) { skuCol = match; break; }
  }

  return { serviceCol, nameCol, skuCol };
}

/**
 * Strip assessment-generated columns from a resource before re-assessment.
 */
function stripAssessmentColumns(resource) {
  const assessmentKeys = [
    'SUBSCRIPTION MOVE SUPPORTED', 'REGION MOVE SUPPORTED',
    'NORMALIZED TYPE', 'REMARKS', 'JIO REGION AVAILABLE',
    'AWS SERVICE', 'GCP SERVICE', 'AZURE EQUIVALENT', 'AZURE RESOURCE TYPE',
    'CATEGORY', 'SIMILARITY', 'SKU RECOMMENDATION',
    'DESCRIPTION', 'MIGRATION NOTES', 'NAME'
  ];
  const cleaned = {};
  for (const [key, value] of Object.entries(resource)) {
    if (!assessmentKeys.includes(key)) {
      cleaned[key] = value;
    }
  }
  return cleaned;
}

module.exports = {
  detectTypeColumn,
  detectNameColumn,
  detectLocationColumn,
  detectAwsColumns,
  detectGcpColumns,
  stripAssessmentColumns
};
