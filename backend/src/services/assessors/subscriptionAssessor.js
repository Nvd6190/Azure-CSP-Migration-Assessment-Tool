'use strict';

const { detectTypeColumn, detectNameColumn, stripAssessmentColumns } = require('../utils/typeDetector');

/**
 * Subscription Move Assessor
 * Self-contained module for assessing subscription move support.
 * Errors in this module will NOT crash other assessment modes.
 */
class SubscriptionAssessor {
  constructor() {
    // Will be populated by init() call from orchestrator
    this.rules = {};
    this._ruleKeys = [];
  }

  /**
   * Initialize with rules from the orchestrator.
   * Called after rules are loaded/refreshed.
   */
  init(rules) {
    this.rules = rules || {};
    this._ruleKeys = Object.keys(this.rules);
  }

  /**
   * Assess subscription move support for a given resource type.
   */
  assessType(rawType, normalizeTypeFn) {
    const normalized = normalizeTypeFn(rawType);
    const rule = this._fuzzyMatchRule(normalized);

    if (rule) {
      const actionableRemark = this._subscriptionRemarks[normalized]
        || rule.remarks
        || this._generateRemark(normalized, rule.subscriptionMove);
      return {
        originalType: rawType,
        normalizedType: normalized,
        subscriptionMove: rule.subscriptionMove,
        remarks: actionableRemark
      };
    }

    return {
      originalType: rawType,
      normalizedType: normalized,
      subscriptionMove: 'Review',
      remarks: 'Resource type not found in the move matrix. Verify manually at https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources'
    };
  }

  _fuzzyMatchRule(normalizedType) {
    if (this.rules[normalizedType]) {
      return this.rules[normalizedType];
    }

    const parts = normalizedType.split('/');
    if (parts.length > 2) {
      const parentType = parts.slice(0, 2).join('/');
      if (this.rules[parentType]) {
        return {
          ...this.rules[parentType],
          remarks: this.rules[parentType].remarks
            ? `${this.rules[parentType].remarks} (Matched via parent type: ${parentType})`
            : `Matched via parent type: ${parentType}. Child resources typically follow parent move support.`
        };
      }
    }

    if (!normalizedType.startsWith('microsoft.')) {
      for (const key of this._ruleKeys) {
        if (key.endsWith('/' + normalizedType)) {
          return {
            ...this.rules[key],
            remarks: this.rules[key].remarks
              ? `${this.rules[key].remarks} (Matched as ${key})`
              : `Matched as ${key}`
          };
        }
      }
    }

    return null;
  }

  /**
   * Assess subscription move support for an array of resources.
   */
  assessResources(resources, normalizeTypeFn) {
    if (!resources || resources.length === 0) return [];

    const firstRow = resources[0];
    const typeCol = detectTypeColumn(firstRow);
    const nameCol = detectNameColumn(firstRow);

    if (!typeCol) {
      const availableCols = Object.keys(firstRow).join(', ');
      throw new Error(`Could not find a "Resource Type" column in your file. Available columns: [${availableCols}]. Expected one of: TYPE, Resource Type, ResourceType, or a column containing the word "type".`);
    }
    console.log(`[Subscription Assessment] Detected resource type column: "${typeCol}" | Name column: "${nameCol}" | All columns: [${Object.keys(firstRow).join(', ')}]`);

    return resources.map(resource => {
      const clean = stripAssessmentColumns(resource);
      const typeValue = typeCol ? (resource[typeCol] || '') : '';
      const nameValue = nameCol ? (resource[nameCol] || '') : '';
      const assessment = this.assessType(typeValue, normalizeTypeFn);

      return {
        ...clean,
        'NAME': nameValue,
        'SUBSCRIPTION MOVE SUPPORTED': assessment.subscriptionMove,
        'NORMALIZED TYPE': assessment.normalizedType,
        'REMARKS': assessment.remarks
      };
    });
  }

  getSummary(assessedResources) {
    const total = assessedResources.length;
    const yes = assessedResources.filter(r => r['SUBSCRIPTION MOVE SUPPORTED'] === 'Yes').length;
    const no = assessedResources.filter(r => r['SUBSCRIPTION MOVE SUPPORTED'] === 'No').length;
    const conditional = assessedResources.filter(r => r['SUBSCRIPTION MOVE SUPPORTED'] === 'Conditional').length;
    const review = assessedResources.filter(r => r['SUBSCRIPTION MOVE SUPPORTED'] === 'Review').length;

    return { total, yes, no, review, conditional };
  }

  _generateRemark(normalizedType, subscriptionMove) {
    const parts = normalizedType.split('/');
    const provider = (parts[0] || '').replace('microsoft.', '');
    const resource = parts.slice(1).join('/');

    const providerGuidance = {
      'compute': { yes: 'Can be moved between subscriptions. Ensure all dependent resources (NICs, disks, IPs) are moved together.', no: 'CANNOT be moved. Recreate the compute resource in the target subscription.', conditional: 'Move support depends on the SKU/configuration. Check specific prerequisites on Microsoft Learn.' },
      'network': { yes: 'Can be moved between subscriptions. Move with associated VNet and dependent resources.', no: 'CANNOT be moved. Recreate the networking resource in the target subscription.', conditional: 'Move support depends on SKU (Basic vs Standard). Check prerequisites on Microsoft Learn.' },
      'storage': { yes: 'Can be moved between subscriptions. Data remains intact. Update RBAC and network rules post-move.', no: 'CANNOT be moved. Create a new storage account and copy data using AzCopy.', conditional: 'Move may require disabling certain features first. Check Microsoft Learn.' },
      'sql': { yes: 'Can be moved between subscriptions. Databases move with the server.', no: 'CANNOT be moved. Use backup/restore or geo-replication to migrate.', conditional: 'Move depends on specific configuration. Check failover groups and linked servers.' },
      'web': { yes: 'Can be moved between subscriptions within the same region. Move App Service Plan and all apps together.', no: 'CANNOT be moved. Redeploy using ARM templates or CI/CD pipelines.', conditional: 'Some web resources have region or plan restrictions. Check App Service move limitations.' },
      'keyvault': { yes: 'Can be moved. Disable disk encryption links first if any. Update access policies for the target subscription tenant.', no: 'CANNOT be moved. Create a new vault and migrate secrets/keys/certificates.', conditional: 'Move depends on soft-delete status and encryption usage. Check prerequisites.' },
      'containerservice': { yes: 'Can be moved between subscriptions.', no: 'CANNOT be moved. Recreate cluster and redeploy workloads in the target subscription.', conditional: 'Check specific AKS version and feature constraints.' },
      'insights': { yes: 'Can be moved between subscriptions. Update instrumentation keys/connection strings in applications.', no: 'CANNOT be moved. Recreate monitoring resource and reconfigure data collection.', conditional: 'Some monitoring resources have move constraints. Check Microsoft Learn.' },
      'logic': { yes: 'Can be moved. Recreate or move API connections separately.', no: 'CANNOT be moved. Export and re-import workflow definitions.', conditional: 'Check managed connector dependencies before move.' },
      'automation': { yes: 'Can be moved. Runbooks and schedules are preserved.', no: 'CANNOT be moved. Create new account and migrate runbooks.', conditional: 'Check Run As account and hybrid worker dependencies.' },
      'servicebus': { yes: 'Can be moved. Queues, topics, subscriptions are preserved.', no: 'CANNOT be moved. Recreate namespace and migrate data.', conditional: 'Check geo-DR pairing before move.' },
      'eventhub': { yes: 'Can be moved. Event hubs and consumer groups are preserved.', no: 'CANNOT be moved. Recreate namespace.', conditional: 'Check dedicated cluster and geo-DR status.' },
      'datafactory': { yes: 'Can be moved. Pipelines and linked services preserved. Reconfigure self-hosted IRs if needed.', no: 'CANNOT be moved. Export ARM template and recreate.', conditional: 'Check managed VNet and self-hosted IR status.' },
      'documentdb': { yes: 'Can be moved. All databases and containers are preserved.', no: 'CANNOT be moved. Create new account and migrate data.', conditional: 'Check multi-region write and private endpoint configuration.' },
      'devices': { yes: 'Can be moved. Device identities are preserved.', no: 'CANNOT be moved. Create new IoT Hub and re-register devices.', conditional: 'Check DPS enrollment and routing configuration.' },
      'recoveryservices': { yes: 'Can be moved with constraints. Stop active backup protection first if protecting Azure VMs.', no: 'CANNOT be moved. Create new vault and re-protect workloads.', conditional: 'Move depends on backup items. Stop protection or delete backup data first.' },
      'cache': { yes: 'Can be moved. Data persistence and settings are preserved.', no: 'CANNOT be moved. Create new cache and migrate data via export/import.', conditional: 'Check VNet and private endpoint configuration.' },
      'cognitiveservices': { yes: 'Can be moved. API keys remain the same. Update endpoint URLs.', no: 'CANNOT be moved. Create new resource and update application endpoints.', conditional: 'Check custom domain and VNet configuration.' },
      'dbformysql': { yes: 'Can be moved. Reconfigure VNet rules post-move.', no: 'CANNOT be moved. Use backup/restore or read replicas.', conditional: 'Check VNet integration and private endpoint status.' },
      'dbforpostgresql': { yes: 'Can be moved. Reconfigure VNet rules post-move.', no: 'CANNOT be moved. Use backup/restore or read replicas.', conditional: 'Check VNet integration and private endpoint status.' },
      'containerregistry': { yes: 'Can be moved. Images preserved. Update webhooks and VNet rules post-move.', no: 'CANNOT be moved. Create new registry and push images.', conditional: 'Check geo-replication and VNet rules.' },
      'machinelearningservices': { yes: 'Can be moved. Compute targets may need reconfiguration.', no: 'CANNOT be moved. Create new workspace and retrain models.', conditional: 'Check compute instances and managed endpoints.' },
      'search': { yes: 'Can be moved. Indexes and data preserved.', no: 'CANNOT be moved. Create new service and rebuild indexes.', conditional: 'Check private endpoint and managed identity configuration.' },
    };

    const guidance = providerGuidance[provider];
    if (guidance) {
      const moveKey = subscriptionMove.toLowerCase().includes('yes') ? 'yes'
        : subscriptionMove.toLowerCase().includes('no') ? 'no' : 'conditional';
      return guidance[moveKey] || guidance.yes;
    }

    if (subscriptionMove === 'Yes') {
      return `Can be moved between subscriptions. Use az resource move CLI or Azure portal. Verify dependent resources are included.`;
    }
    if (subscriptionMove === 'No') {
      return `CANNOT be moved between subscriptions. Workaround: Export ARM template, recreate ${resource} in the target subscription. See https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources`;
    }
    return `Conditional move — check specific prerequisites for ${resource} on https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources`;
  }

  get _subscriptionRemarks() {
    if (this.__subscriptionRemarks) return this.__subscriptionRemarks;
    this.__subscriptionRemarks = {
      'microsoft.compute/virtualmachines': 'Can be moved directly. Pre-checks: (1) Detach any Azure Backup vault before move. (2) Move all dependent NICs, disks, and public IPs together. (3) VMs in availability sets must be moved as a set.',
      'microsoft.compute/disks': 'Managed disks can be moved directly between subscriptions. Ensure the VM is deallocated first if the disk is attached.',
      'microsoft.compute/snapshots': 'Full snapshots: movable directly. Incremental snapshots: CANNOT be moved across subscriptions.',
      'microsoft.compute/virtualmachinescalesets': 'Can be moved between subscriptions. Move all dependent resources (VNets, load balancers, public IPs) together.',
      'microsoft.network/publicipaddresses': 'Basic SKU: Can be moved directly. Standard SKU: MUST detach from associated resource first, then move separately.',
      'microsoft.network/loadbalancers': 'Basic SKU: Can be moved directly. Standard SKU: CANNOT be moved. Recreate using ARM template export.',
      'microsoft.network/virtualnetworks': 'Can be moved but ALL dependent resources (VMs, NICs, NSGs, route tables) must be moved together.',
      'microsoft.network/azurefirewalls': 'CANNOT be moved. Export ARM template, delete firewall, recreate in target subscription.',
      'microsoft.sql/servers': 'Can be moved. All databases under the server move with it. Update firewall rules and AAD admin config post-move.',
      'microsoft.storage/storageaccounts': 'Can be moved. Private endpoints, VNet rules, and RBAC must be updated post-move. Data remains intact.',
      'microsoft.keyvault/vaults': 'Can be moved with CRITICAL prerequisites: (1) Disable disk encryption links. (2) Soft-delete must be enabled. (3) Update access policies for new tenant.',
      'microsoft.web/sites': 'Can be moved within the SAME region. Move App Service Plan together. Custom domains and SSL certs must be reconfigured.',
      'microsoft.containerservice/managedclusters': 'AKS clusters CANNOT be moved. Export workload manifests, create new cluster in target, redeploy.',
      'microsoft.recoveryservices/vaults': 'Can be moved with constraints: Cannot move if protecting Azure VMs. Stop backup first.',
      'microsoft.containerregistry/registries': 'Can be moved. Image data is preserved. Update webhook URLs and VNet rules post-move.',
      'microsoft.storagesync/storagesyncservices': 'Can be moved between subscriptions.',
      'microsoft.powerbidedicated/capacities': 'Can be moved between subscriptions.',
    };
    return this.__subscriptionRemarks;
  }
}

module.exports = new SubscriptionAssessor();
