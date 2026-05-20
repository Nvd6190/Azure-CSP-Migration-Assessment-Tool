'use strict';

const { detectTypeColumn, detectNameColumn, stripAssessmentColumns } = require('../utils/typeDetector');

/**
 * Region Move Assessor
 * Self-contained module for assessing region move support.
 * Errors in this module will NOT crash other assessment modes.
 */
class RegionAssessor {
  constructor() {
    this.regionRules = {};
  }

  /**
   * Initialize with region rules from the orchestrator.
   */
  init(regionRules) {
    this.regionRules = regionRules || {};
  }

  /**
   * Assess region move support for a given resource type.
   */
  assessType(rawType, normalizeTypeFn) {
    const normalized = normalizeTypeFn(rawType);

    if (this.regionRules[normalized]) {
      const regionMove = this.regionRules[normalized].regionMove;
      const remark = this._regionRemarks[normalized] || this._generateRemark(normalized, regionMove);
      return {
        originalType: rawType,
        normalizedType: normalized,
        regionMove,
        remarks: remark
      };
    }

    const parts = normalized.split('/');
    if (parts.length > 2) {
      const parentType = parts.slice(0, 2).join('/');
      if (this.regionRules[parentType]) {
        const regionMove = this.regionRules[parentType].regionMove;
        return {
          originalType: rawType,
          normalizedType: normalized,
          regionMove,
          remarks: `Matched via parent type: ${parentType}. Child resources typically follow parent region move support.`
        };
      }
    }

    return {
      originalType: rawType,
      normalizedType: normalized,
      regionMove: 'Review',
      remarks: 'Resource type not found in the region move matrix. Verify manually at https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-region'
    };
  }

  /**
   * Assess region move support for an array of resources.
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
    console.log(`[Region Assessment] Detected resource type column: "${typeCol}" | Name column: "${nameCol}"`);

    return resources.map(resource => {
      const clean = stripAssessmentColumns(resource);
      const typeValue = typeCol ? (resource[typeCol] || '') : '';
      const nameValue = nameCol ? (resource[nameCol] || '') : '';
      const assessment = this.assessType(typeValue, normalizeTypeFn);

      return {
        ...clean,
        'NAME': nameValue,
        'REGION MOVE SUPPORTED': assessment.regionMove,
        'NORMALIZED TYPE': assessment.normalizedType,
        'REMARKS': assessment.remarks
      };
    });
  }

  getSummary(assessedResources) {
    const total = assessedResources.length;
    const yes = assessedResources.filter(r => r['REGION MOVE SUPPORTED'] === 'Yes').length;
    const no = assessedResources.filter(r => r['REGION MOVE SUPPORTED'] === 'No').length;
    const review = assessedResources.filter(r => r['REGION MOVE SUPPORTED'] === 'Review').length;

    return { total, yes, no, review };
  }

  _generateRemark(normalizedType, regionMove) {
    const parts = normalizedType.split('/');
    const provider = (parts[0] || '').replace('microsoft.', '');
    const resource = parts.slice(1).join('/');

    const providerGuidance = {
      'compute': { yes: 'Use Azure Resource Mover for cross-region migration.', no: 'Region move not supported. Redeploy in the target region using ARM/Bicep templates.' },
      'network': { yes: 'Can be moved across regions using Azure Resource Mover.', no: 'Region move not supported. Recreate the networking resource in the target region.' },
      'storage': { yes: 'Create a new storage account in the target region and use AzCopy to migrate data.', no: 'Region move not supported. Create new storage in target region and copy data using AzCopy.' },
      'sql': { yes: 'Use active geo-replication or failover groups for cross-region move.', no: 'Region move not supported. Use backup/restore or geo-replication to migrate.' },
      'dbformysql': { yes: 'Use read replicas or backup/restore for cross-region migration.', no: 'Region move not supported. Use backup/restore or create a new server in the target region.' },
      'dbforpostgresql': { yes: 'Use read replicas or backup/restore for cross-region migration.', no: 'Region move not supported. Use backup/restore or create a new server in the target region.' },
      'web': { yes: 'Redeploy to the target region using ARM templates or CI/CD.', no: 'Region move not supported. Redeploy using ARM templates or CI/CD.' },
      'containerservice': { yes: 'Use Azure Resource Mover or redeploy cluster in the target region.', no: 'Region move not supported. Redeploy cluster and workloads in the target region.' },
      'keyvault': { yes: 'Move vault to target region.', no: 'Region move not supported. Create a new vault in the target region and migrate secrets/keys/certificates.' },
      'documentdb': { yes: 'Use multi-region writes or create account in the new region.', no: 'Region move not supported. Add target region as a replica, then remove old region.' },
      'cache': { yes: 'Create new cache in target region and migrate data.', no: 'Region move not supported. Create a new cache instance and migrate data.' },
      'insights': { yes: 'Recreate monitoring resource in the target region.', no: 'Region move not supported. Create a new resource in the target region.' },
      'servicebus': { yes: 'Use geo-disaster recovery pairing for cross-region move.', no: 'Region move not supported. Recreate namespace in the target region.' },
      'eventhub': { yes: 'Use geo-disaster recovery pairing for cross-region move.', no: 'Region move not supported. Recreate namespace.' },
      'datafactory': { yes: 'Export and re-import pipelines in a new factory in the target region.', no: 'Region move not supported. Create a new factory.' },
      'logic': { yes: 'Export and re-import Logic App definition in the target region.', no: 'Region move not supported. Export and re-import definition.' },
    };

    const guidance = providerGuidance[provider];
    if (guidance) {
      return regionMove === 'Yes' ? guidance.yes : guidance.no;
    }

    if (regionMove === 'Yes') {
      return `Region move supported. Use Azure Resource Mover or redeploy ${resource} in the target region.`;
    }
    return `Region move not supported for ${normalizedType}. Redeploy in the target region using ARM/Bicep templates.`;
  }

  get _regionRemarks() {
    if (this.__regionRemarks) return this.__regionRemarks;
    this.__regionRemarks = {
      'microsoft.compute/virtualmachines': 'Use Azure Resource Mover to move VMs across regions.',
      'microsoft.compute/disks': 'Managed disks can be moved across regions using Azure Resource Mover.',
      'microsoft.compute/virtualmachinescalesets': 'VMSS does not support direct region move. Redeploy using ARM templates.',
      'microsoft.network/virtualnetworks': 'VNets can be moved across regions using Azure Resource Mover.',
      'microsoft.network/networksecuritygroups': 'NSGs can be moved across regions using Azure Resource Mover.',
      'microsoft.network/publicipaddresses': 'Public IPs can be moved across regions. Note: IP address value will change.',
      'microsoft.network/loadbalancers': 'Load balancers can be moved across regions using Azure Resource Mover.',
      'microsoft.network/applicationgateways': 'Application Gateway does not support region move. Redeploy in the target region.',
      'microsoft.network/azurefirewalls': 'Azure Firewall does not support region move. Deploy a new firewall in the target region.',
      'microsoft.network/bastionhosts': 'Bastion does not support region move. Deploy a new Bastion host.',
      'microsoft.network/vpngateways': 'VPN Gateways do not support region move. Create a new gateway.',
      'microsoft.network/expressroutecircuits': 'ExpressRoute circuits are tied to peering locations.',
      'microsoft.network/dnszones': 'DNS zones are global. No region move needed.',
      'microsoft.network/privatednszones': 'Private DNS zones are global. Re-link VNets in the target region.',
      'microsoft.network/trafficmanagerprofiles': 'Traffic Manager is global. No region move needed.',
      'microsoft.network/frontdoors': 'Front Door is global. No region move needed.',
      'microsoft.sql/servers': 'Use active geo-replication or failover groups for cross-region move.',
      'microsoft.sql/servers/databases': 'Databases can be moved using geo-replication failover or backup/restore.',
      'microsoft.sql/managedinstances': 'SQL MI supports cross-region move via re-creation in target region.',
      'microsoft.storage/storageaccounts': 'Create a new account in the target region and copy data with AzCopy.',
      'microsoft.documentdb/databaseaccounts': 'Add target region as replica, failover, then remove the old region.',
      'microsoft.cache/redis': 'Azure Cache for Redis does not support region move. Create new cache in target region.',
      'microsoft.keyvault/vaults': 'Key Vault does not support region move. Create a new vault and migrate secrets/keys.',
      'microsoft.web/sites': 'App Service does not support direct region move. Redeploy or clone app.',
      'microsoft.web/serverfarms': 'App Service Plans do not support region move. Create a new plan.',
      'microsoft.containerservice/managedclusters': 'AKS clusters do not support region move. Redeploy cluster and workloads.',
      'microsoft.containerregistry/registries': 'ACR does not support region move. Enable geo-replication to the target region.',
      'microsoft.insights/components': 'Application Insights does not support region move. Create new resource and reconfigure.',
      'microsoft.operationalinsights/workspaces': 'Log Analytics workspaces do not support region move. Create new workspace.',
      'microsoft.servicebus/namespaces': 'Service Bus can be moved using geo-disaster recovery pairing.',
      'microsoft.eventhub/namespaces': 'Event Hubs can be moved using geo-disaster recovery pairing.',
      'microsoft.datafactory/factories': 'Data Factory does not support region move. Create new factory and export/import pipelines.',
      'microsoft.logic/workflows': 'Logic Apps do not support region move. Export and redeploy.',
      'microsoft.apimanagement/service': 'API Management supports region move via backup/restore.',
      'microsoft.recoveryservices/vaults': 'Recovery Services vaults do not support region move. Create new vault.',
      'microsoft.devices/iothubs': 'IoT Hub supports cross-region move via manual failover.',
      'microsoft.cognitiveservices/accounts': 'Cognitive Services does not support region move. Create new resource.',
      'microsoft.signalrservice/signalr': 'SignalR does not support region move. Create a new instance.',
      'microsoft.cdn/profiles': 'CDN is global. No region move needed.',
      'microsoft.portal/dashboards': 'Dashboards are global. No region move needed.',
    };
    return this.__regionRemarks;
  }
}

module.exports = new RegionAssessor();
