'use strict';

const fs = require('fs');
const path = require('path');
const { detectTypeColumn, detectNameColumn, detectLocationColumn, stripAssessmentColumns } = require('../utils/typeDetector');

/**
 * Jio Region Availability Assessor
 * Self-contained module for checking Azure service availability in Jio India West region.
 * Errors in this module will NOT crash other assessment modes.
 */
class JioAssessor {
  constructor() {
    this._jioJsonPath = path.join(__dirname, '..', '..', 'data', 'jio-availability.json');
    this.jioServices = {};
    this.jioVMs = {};
    this.jioMetadata = {};
    this.__armToJioName = null;
    this._loadData();
  }

  _loadData() {
    try {
      const data = JSON.parse(fs.readFileSync(this._jioJsonPath, 'utf-8'));
      this.jioServices = data.services || {};
      this.jioVMs = data.vms || {};
      this.jioMetadata = data._metadata || {};
    } catch (err) {
      console.error('[JioAssessor] Failed to load jio-availability.json:', err.message);
    }
  }

  /**
   * Parse an uploaded Jio availability Excel and update the JSON + in-memory data.
   */
  refreshFromExcel(excelPath) {
    const XLSX = require('xlsx');
    const workbook = XLSX.readFile(excelPath);

    const servicesSheetName = workbook.SheetNames.find(s => /services/i.test(s) && !/availability/i.test(s));
    if (!servicesSheetName) throw new Error('Could not find a "Services" sheet in the uploaded file.');
    const servicesRows = XLSX.utils.sheet_to_json(workbook.Sheets[servicesSheetName]);

    const services = {};
    for (const row of servicesRows) {
      const nameKey = Object.keys(row).find(k => /service/i.test(k) && /name/i.test(k));
      const statusKey = Object.keys(row).find(k => /availability|status/i.test(k));
      if (!nameKey || !statusKey) continue;
      const name = String(row[nameKey] || '').trim();
      const status = String(row[statusKey] || '').trim();
      if (!name) continue;
      services[name.toLowerCase()] = { name, available: status };
    }

    const vmsSheetName = workbook.SheetNames.find(s => /vm/i.test(s));
    const vms = {};
    if (vmsSheetName) {
      const vmRows = XLSX.utils.sheet_to_json(workbook.Sheets[vmsSheetName]);
      for (const row of vmRows) {
        const nameKey = Object.keys(row).find(k => /vm|series/i.test(k));
        const statusKey = Object.keys(row).find(k => /availability|status/i.test(k));
        const remarksKey = Object.keys(row).find(k => /remark/i.test(k));
        if (!nameKey || !statusKey) continue;
        const name = String(row[nameKey] || '').trim();
        const status = String(row[statusKey] || '').trim();
        const remarks = remarksKey ? String(row[remarksKey] || '').trim() : '';
        if (!name) continue;
        vms[name.toLowerCase()] = { name, available: status, remarks };
      }
    }

    if (Object.keys(services).length === 0) {
      throw new Error('No services found in the uploaded Excel. Check sheet format (expected columns: "Services Names", "Availability Status").');
    }

    const jioData = {
      services,
      vms,
      _metadata: {
        source: 'Jio Region Availability Matrix',
        totalServices: Object.keys(services).length,
        totalVMs: Object.keys(vms).length,
        lastUpdated: new Date().toISOString()
      }
    };

    fs.writeFileSync(this._jioJsonPath, JSON.stringify(jioData, null, 2), 'utf-8');

    const backupPath = path.join(path.dirname(this._jioJsonPath), 'jio-availability.xlsx');
    fs.copyFileSync(excelPath, backupPath);

    this.jioServices = services;
    this.jioVMs = vms;
    this.jioMetadata = jioData._metadata;
    this.__armToJioName = null;

    console.log(`Jio data refreshed: ${jioData._metadata.totalServices} services, ${jioData._metadata.totalVMs} VMs`);
    return jioData._metadata;
  }

  get _armToJioName() {
    if (this.__armToJioName) return this.__armToJioName;
    this.__armToJioName = {
      'microsoft.network/applicationgateways': 'Application Gateway : Basic and Standard and WAF',
      'microsoft.documentdb/databaseaccounts': 'Azure Cosmos DB',
      'microsoft.network/dnszones': 'Azure DNS',
      'microsoft.network/privatednszones': 'Azure DNS',
      'microsoft.eventhub/namespaces': 'Azure Event Hubs',
      'microsoft.network/expressroutecircuits': 'Azure ExpressRoute',
      'microsoft.keyvault/vaults': 'Azure Key Vault',
      'microsoft.keyvault/managedhsms': 'Azure Key Vault',
      'microsoft.containerservice/managedclusters': 'Azure Kubernetes Service (AKS)',
      'microsoft.network/loadbalancers': 'Azure Load Balancer',
      'microsoft.network/natgateways': 'Azure NAT Gateway',
      'microsoft.network/publicipaddresses': 'Azure Public IP',
      'microsoft.servicebus/namespaces': 'Azure Service Bus',
      'microsoft.servicefabric/clusters': 'Azure Service Fabric',
      'microsoft.recoveryservices/vaults': 'Azure Site Recovery',
      'microsoft.sql/servers': 'Azure SQL',
      'microsoft.sql/servers/databases': 'Azure SQL',
      'microsoft.sql/servers/elasticpools': 'Azure SQL',
      'microsoft.sql/managedinstances': 'Azure SQL Managed Instance',
      'microsoft.storage/storageaccounts': 'Azure Storage : Blob Storage',
      'microsoft.compute/virtualmachines': 'Virtual Machines',
      'microsoft.network/virtualnetworks': 'Azure Virtual Network',
      'microsoft.network/virtualnetworkgateways': 'Azure VPN Gateway',
      'microsoft.network/vpngateways': 'Azure VPN Gateway',
      'microsoft.backup/vaults': 'Azure Backup',
      'microsoft.recoveryservices/vaults/backuppolicies': 'Azure Backup',
      'microsoft.search/searchservices': 'Azure AI Search',
      'microsoft.cognitiveservices/accounts': 'Azure AI services',
      'microsoft.apimanagement/service': 'Azure API Management',
      'microsoft.appconfiguration/configurationstores': 'Azure App Configuration',
      'microsoft.web/sites': 'Azure App Service',
      'microsoft.web/serverfarms': 'Azure App Service',
      'microsoft.network/bastionhosts': 'Azure Bastion',
      'microsoft.batch/batchaccounts': 'Azure Batch',
      'microsoft.cache/redis': 'Azure Cache for Redis',
      'microsoft.containerinstance/containergroups': 'Azure Container Instances',
      'microsoft.containerregistry/registries': 'Azure Container Registry',
      'microsoft.kusto/clusters': 'Azure Data Explorer',
      'microsoft.datafactory/factories': 'Azure Data Factory',
      'microsoft.dbformysql/servers': 'Azure Database for MySQL',
      'microsoft.dbformysql/flexibleservers': 'Azure Database for MySQL - Flexible Servers',
      'microsoft.dbforpostgresql/servers': 'Azure Database for PostgreSQL',
      'microsoft.dbforpostgresql/flexibleservers': 'Azure Database for PostgreSQL',
      'microsoft.network/ddosprotectionplans': 'Azure DDoS Protection',
      'microsoft.databricks/workspaces': 'Azure Databricks',
      'microsoft.network/azurefirewalls': 'Azure Firewall',
      'microsoft.network/firewallpolicies': 'Azure Firewall',
      'microsoft.cdn/profiles': 'Azure Front Door : Standard and Premium Profiles',
      'microsoft.network/frontdoors': 'Azure Front Door : Standard and Premium Profiles',
      'microsoft.compute/disks': 'Azure Managed Disks',
      'microsoft.insights/components': 'Azure Monitor',
      'microsoft.insights/actiongroups': 'Azure Monitor',
      'microsoft.insights/metricalerts': 'Azure Monitor',
      'microsoft.operationalinsights/workspaces': 'Azure Monitor : Log Analytics',
      'microsoft.network/networksecuritygroups': 'Azure Network Security Group',
      'microsoft.network/networkinterfaces': 'Azure Virtual Network',
      'microsoft.network/privateendpoints': 'Azure Private Link',
      'microsoft.network/privatelinkservices': 'Azure Private Link',
      'microsoft.portal/dashboards': 'Azure portal',
      'microsoft.network/routetables': 'Azure Route Server',
      'microsoft.signalrservice/signalr': 'Azure SignalR Service',
      'microsoft.network/trafficmanagerprofiles': 'Azure Traffic Manager',
      'microsoft.compute/virtualmachinescalesets': 'Virtual Machine Scale Sets',
      'microsoft.logic/workflows': 'Azure Logic Apps',
      'microsoft.automation/automationaccounts': 'Azure Automation',
      'microsoft.devices/iothubs': 'Azure IoT Hub',
      'microsoft.devices/provisioningservices': 'Azure IoT Hub Device Provisioning Service',
      'microsoft.hdinsight/clusters': 'Azure HDInsight',
      'microsoft.machinelearningservices/workspaces': 'Azure Machine Learning',
      'microsoft.media/mediaservices': 'Azure Media Services',
      'microsoft.purview/accounts': 'Microsoft Purview',
      'microsoft.eventgrid/topics': 'Azure Event Grid',
      'microsoft.eventgrid/domains': 'Azure Event Grid',
      'microsoft.eventgrid/systemtopics': 'Azure Event Grid',
      'microsoft.notificationhubs/namespaces': 'Azure Notification Hubs',
      'microsoft.relay/namespaces': 'Azure Relay',
      'microsoft.app/containerapps': 'Azure Container Apps',
      'microsoft.app/managedenvironments': 'Azure Container Apps',
      'microsoft.desktopvirtualization/hostpools': 'Azure Virtual Desktop',
      'microsoft.desktopvirtualization/applicationgroups': 'Azure Virtual Desktop',
      'microsoft.desktopvirtualization/workspaces': 'Azure Virtual Desktop',
      'microsoft.maps/accounts': 'Azure Maps',
      'microsoft.communication/communicationservices': 'Azure Communication Services',
      'microsoft.managedidentity/userassignedidentities': 'Azure Managed Identity',
      'microsoft.dbformariadb/servers': 'Azure Database for MariaDB',
      'microsoft.datamigration/services': 'Azure Database Migration Service',
      'microsoft.synapse/workspaces': 'Azure Synapse Analytics',
      'microsoft.streamanalytics/streamingjobs': 'Azure Stream Analytics',
      'microsoft.datalakeanalytics/accounts': 'Azure Data Lake',
      'microsoft.datalakestore/accounts': 'Azure Data Lake',
      'microsoft.datashare/accounts': 'Azure Data Share',
      'microsoft.analysisservices/servers': 'Azure Analysis Services',
      'microsoft.healthcareapis/services': 'Azure API for FHIR',
      'microsoft.digitaltwins/digitaltwinsinstances': 'Azure Digital Twins',
      'microsoft.iotcentral/iotapps': 'Azure IoT Central',
      'microsoft.botservice/botservices': 'Azure Bot Service',
      'microsoft.web/staticsites': 'Azure Static Web Apps',
      'microsoft.storagesync/storagesyncservices': 'Azure File Sync',
      'microsoft.netapp/netappaccounts': 'Azure NetApp Files',
      'microsoft.compute/availabilitysets': 'Availability Sets',
      'microsoft.compute/snapshots': 'Azure Managed Disks',
      'microsoft.compute/images': 'Virtual Machines',
      'microsoft.web/hostingenvironments': 'Azure App Service Environments',
      'microsoft.powerbidedicated/capacities': 'Power BI Embedded',
      'microsoft.sqlvirtualmachine/sqlvirtualmachines': 'SQL Server on Azure VMs',
      'microsoft.redhatopenshift/openshiftclusters': 'Azure Red Hat OpenShift',
      'microsoft.appplatform/spring': 'Azure Spring Apps',
      'microsoft.loadtestservice/loadtests': 'Azure Load Testing',
      'microsoft.avs/privateclouds': 'Azure VMware Solution',
      'microsoft.devtestlab/labs': 'Azure DevTest Labs',
      'microsoft.hybridcompute/machines': 'Azure Arc',
    };
    return this.__armToJioName;
  }

  /**
   * Assess Jio region availability for a given resource type.
   * Requires normalizeType from parent service.
   */
  assessType(rawType, normalizeTypeFn) {
    const normalized = normalizeTypeFn(rawType);

    const jioName = this._armToJioName[normalized];
    if (jioName) {
      const jioKey = jioName.toLowerCase();
      if (this.jioServices[jioKey]) {
        return {
          originalType: rawType,
          normalizedType: normalized,
          jioAvailable: this.jioServices[jioKey].available,
          jioServiceName: this.jioServices[jioKey].name,
          remarks: this.jioServices[jioKey].available === 'Yes'
            ? `${this.jioServices[jioKey].name} is available in Jio India West region.`
            : `${this.jioServices[jioKey].name} is NOT available in Jio India West region. Consider alternative services or regions.`
        };
      }
    }

    const displayName = rawType.trim().toLowerCase();
    const fuzzyMatch = this._fuzzyMatch(displayName) || this._fuzzyMatch(normalized);
    if (fuzzyMatch) {
      return {
        originalType: rawType,
        normalizedType: normalized,
        jioAvailable: fuzzyMatch.available,
        jioServiceName: fuzzyMatch.name,
        remarks: fuzzyMatch.available === 'Yes'
          ? `${fuzzyMatch.name} is available in Jio India West region.`
          : `${fuzzyMatch.name} is NOT available in Jio India West region. Consider alternative services or regions.`
      };
    }

    return {
      originalType: rawType,
      normalizedType: normalized,
      jioAvailable: 'Review',
      jioServiceName: '',
      remarks: 'Service not found in Jio India West availability matrix. Verify manually at https://learn.microsoft.com/en-us/azure/reliability/availability-service-by-category'
    };
  }

  _fuzzyMatch(input) {
    if (!input) return null;
    const lower = input.toLowerCase().replace(/^microsoft\.\w+\//, '').replace(/\//g, ' ');

    if (this.jioServices[lower]) return this.jioServices[lower];

    const jioKeys = Object.keys(this.jioServices);
    for (const key of jioKeys) {
      if (key.includes(lower) || lower.includes(key)) {
        return this.jioServices[key];
      }
    }

    const words = lower.replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(w => w.length > 2);
    if (words.length > 0) {
      for (const key of jioKeys) {
        const matchCount = words.filter(w => key.includes(w)).length;
        if (matchCount >= Math.max(1, words.length - 1)) {
          return this.jioServices[key];
        }
      }
    }

    return null;
  }

  /**
   * Assess Jio availability for an array of resources.
   */
  assessResources(resources, normalizeTypeFn) {
    if (!resources || resources.length === 0) return [];

    const firstRow = resources[0];
    const typeCol = detectTypeColumn(firstRow);
    const nameCol = detectNameColumn(firstRow);
    const locationCol = detectLocationColumn(firstRow);

    if (!typeCol) {
      const availableCols = Object.keys(firstRow).join(', ');
      throw new Error(`Could not find a "Resource Type" column in your file. Available columns: [${availableCols}]. Expected one of: TYPE, Resource Type, ResourceType, or a column containing the word "type".`);
    }
    console.log(`[Jio Assessment] Detected resource type column: "${typeCol}" | Name column: "${nameCol}"`);

    const indiaRegions = new Set([
      'centralindia', 'southindia', 'westindia', 'jioindiawest', 'jioindiacentral',
      'central india', 'south india', 'west india', 'jio india west', 'jio india central'
    ]);

    return resources.map(resource => {
      const clean = stripAssessmentColumns(resource);
      const typeValue = typeCol ? (resource[typeCol] || '') : '';
      const nameValue = nameCol ? (resource[nameCol] || '') : '';
      const assessment = this.assessType(typeValue, normalizeTypeFn);

      let regionWarning = '';
      if (locationCol) {
        const location = (resource[locationCol] || '').toString().toLowerCase().trim();
        if (location && !indiaRegions.has(location)) {
          regionWarning = `Current region: ${resource[locationCol]}. Only India regions are supported for Jio migration.`;
        }
      }

      const combinedRemarks = [assessment.remarks, regionWarning].filter(Boolean).join(' | ');

      return {
        ...clean,
        'NAME': nameValue,
        'JIO REGION AVAILABLE': assessment.jioAvailable,
        'JIO SERVICE NAME': assessment.jioServiceName,
        'CURRENT REGION': locationCol ? (resource[locationCol] || '') : '',
        'INDIA REGION': locationCol
          ? (indiaRegions.has((resource[locationCol] || '').toString().toLowerCase().trim()) ? 'Yes' : 'No')
          : '',
        'NORMALIZED TYPE': assessment.normalizedType,
        'REMARKS': combinedRemarks
      };
    });
  }

  getSummary(assessedResources) {
    const total = assessedResources.length;
    const yes = assessedResources.filter(r => r['JIO REGION AVAILABLE'] === 'Yes').length;
    const no = assessedResources.filter(r => r['JIO REGION AVAILABLE'] === 'No').length;
    const review = assessedResources.filter(r => r['JIO REGION AVAILABLE'] === 'Review').length;

    return { total, yes, no, review };
  }
}

module.exports = new JioAssessor();
