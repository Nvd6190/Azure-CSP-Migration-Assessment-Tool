const fs = require('fs');
const path = require('path');
const { fetchAllRules } = require('./rulesFetcher');

// Isolated assessor modules — errors in one won't crash others
let awsAssessor = null;
try {
  awsAssessor = require('./assessors/awsAssessor');
} catch (err) {
  console.error('[WARN] AWS Assessor module failed to load:', err.message);
  console.error('[WARN] AWS-to-Azure assessment will be unavailable. Other modes still work.');
}

let gcpAssessor = null;
try {
  gcpAssessor = require('./assessors/gcpAssessor');
} catch (err) {
  console.error('[WARN] GCP Assessor module failed to load:', err.message);
  console.error('[WARN] GCP-to-Azure assessment will be unavailable. Other modes still work.');
}

let subscriptionAssessor = null;
try {
  subscriptionAssessor = require('./assessors/subscriptionAssessor');
} catch (err) {
  console.error('[WARN] Subscription Assessor module failed to load:', err.message);
  console.error('[WARN] Subscription move assessment will be unavailable. Other modes still work.');
}

let regionAssessor = null;
try {
  regionAssessor = require('./assessors/regionAssessor');
} catch (err) {
  console.error('[WARN] Region Assessor module failed to load:', err.message);
  console.error('[WARN] Region move assessment will be unavailable. Other modes still work.');
}

let jioAssessor = null;
try {
  jioAssessor = require('./assessors/jioAssessor');
} catch (err) {
  console.error('[WARN] Jio Assessor module failed to load:', err.message);
  console.error('[WARN] Jio region assessment will be unavailable. Other modes still work.');
}

class MigrationService {
  constructor() {
    const matrixPath = path.join(__dirname, '..', 'data', 'azureMoveMatrix.json');
    const matrixData = JSON.parse(fs.readFileSync(matrixPath, 'utf-8'));
    this.staticRules = matrixData.rules;
    this.rules = { ...this.staticRules };
    this.regionRules = {};
    this.metadata = matrixData._metadata;
    this.rulesSource = 'static';
    this.lastRefreshed = null;
    this.csvRuleCount = 0;
    this.regionCsvRuleCount = 0;

    // Load Jio availability data
    this._jioJsonPath = path.join(__dirname, '..', 'data', 'jio-availability.json');
    this._loadJioData();

    // Build a reverse lookup index for faster fuzzy matching
    this._ruleKeys = Object.keys(this.rules);

    // Auto-fetch dynamic rules on startup
    this.refreshRules().catch(err => {
      console.warn('Dynamic CSV fetch failed on startup, using static rules:', err.message);
    });
  }

  /**
   * Fetch the latest rules from Microsoft's GitHub CSV and merge with static overrides.
   * Static JSON always wins (manually verified corrections + remarks).
   */
  async refreshRules() {
    const { subscriptionRules, regionRules, source } = await fetchAllRules();

    this.csvRuleCount = Object.keys(subscriptionRules).length;
    this.regionCsvRuleCount = Object.keys(regionRules).length;

    // Merge subscription rules: fetched as base, static overrides win
    this.rules = { ...subscriptionRules, ...this.staticRules };
    this._ruleKeys = Object.keys(this.rules);

    // Region rules from fetched data
    this.regionRules = regionRules;

    this.rulesSource = source + '+static';
    this.lastRefreshed = new Date().toISOString();
    this.metadata.lastUpdated = this.lastRefreshed;

    // Initialize isolated assessor modules with latest rules
    if (subscriptionAssessor) subscriptionAssessor.init(this.rules);
    if (regionAssessor) regionAssessor.init(this.regionRules);

    console.log(`Rules refreshed: ${this.csvRuleCount} subscription + ${this.regionCsvRuleCount} region from ${source}, ${Object.keys(this.staticRules).length} static overrides = ${this._ruleKeys.length} total subscription rules`);
    return {
      csvRules: this.csvRuleCount,
      regionCsvRules: this.regionCsvRuleCount,
      staticOverrides: Object.keys(this.staticRules).length,
      totalRules: this._ruleKeys.length,
      totalRegionRules: Object.keys(this.regionRules).length,
      source: this.rulesSource,
      lastRefreshed: this.lastRefreshed
    };
  }

  /**
   * Load Jio availability data from the JSON file on disk.
   */
  _loadJioData() {
    const jioData = JSON.parse(fs.readFileSync(this._jioJsonPath, 'utf-8'));
    this.jioServices = jioData.services;
    this.jioVMs = jioData.vms;
    this.jioMetadata = jioData._metadata;
  }

  /**
   * Parse an uploaded Jio availability Excel and update the JSON + in-memory data.
   * Expected sheets: "Services" (columns: Services Names, Availability Status)
   *                  "VMs" (columns: VM Series, Availability Status, Remarks)
   */
  refreshJioFromExcel(excelPath) {
    if (jioAssessor) {
      const result = jioAssessor.refreshFromExcel(excelPath);
      // Keep orchestrator in-memory copy in sync
      this.jioServices = jioAssessor.jioServices;
      this.jioVMs = jioAssessor.jioVMs;
      this.jioMetadata = jioAssessor.jioMetadata;
      this.__armToJioName = null;
      return result;
    }
    // Fallback: inline logic if module failed to load
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
      throw new Error('No services found in the uploaded Excel.');
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
    this._loadJioData();
    this.__armToJioName = null;
    console.log(`Jio data refreshed: ${jioData._metadata.totalServices} services, ${jioData._metadata.totalVMs} VMs`);
    return jioData._metadata;
  }

  /**
   * Normalize an Azure resource type string to the ARM format used in the rules JSON.
   * Handles: ARM types (microsoft.xxx/yyy), Azure portal display names, and common aliases.
   */
  normalizeType(rawType) {
    if (!rawType) return '';

    let type = rawType.trim();

    // Already an ARM resource type (case-insensitive check)
    if (/^microsoft\./i.test(type)) {
      return type.toLowerCase();
    }

    const lower = type.toLowerCase();

    // Check display name map
    if (this._displayNameMap[lower]) return this._displayNameMap[lower];

    // Try stripping common prefixes/suffixes
    const variants = [
      lower,
      lower.replace(/^azure\s+/i, ''),
      lower.replace(/s$/, ''),
      lower.replace(/^azure\s+/i, '').replace(/s$/, ''),
    ];

    for (const v of variants) {
      if (this._displayNameMap[v]) return this._displayNameMap[v];
    }

    return type.toLowerCase();
  }

  // Display name → ARM type mapping
  get _displayNameMap() {
    if (this.__displayNameMap) return this.__displayNameMap;
    this.__displayNameMap = {
      // Compute
      'virtual machine': 'microsoft.compute/virtualmachines',
      'vm': 'microsoft.compute/virtualmachines',
      'disk': 'microsoft.compute/disks',
      'managed disk': 'microsoft.compute/disks',
      'availability set': 'microsoft.compute/availabilitysets',
      'virtual machine scale set': 'microsoft.compute/virtualmachinescalesets',
      'vmss': 'microsoft.compute/virtualmachinescalesets',
      'image': 'microsoft.compute/images',
      'snapshot': 'microsoft.compute/snapshots',
      'shared image gallery': 'microsoft.compute/galleries',
      'compute gallery': 'microsoft.compute/galleries',
      'ssh public key': 'microsoft.compute/sshpublickeys',
      'proximity placement group': 'microsoft.compute/proximityplacementgroups',
      'capacity reservation group': 'microsoft.compute/capacityreservationgroups',
      'disk encryption set': 'microsoft.compute/diskencryptionsets',
      'dedicated host group': 'microsoft.compute/hostgroups',
      'dedicated host': 'microsoft.compute/hostgroups/hosts',
      'restore point collection': 'microsoft.compute/restorepointcollections',
      'cloud service': 'microsoft.compute/cloudservices',
      'cloud service (extended support)': 'microsoft.compute/cloudservices',
      'disk access': 'microsoft.compute/diskaccesses',

      // Network
      'virtual network': 'microsoft.network/virtualnetworks',
      'vnet': 'microsoft.network/virtualnetworks',
      'network security group': 'microsoft.network/networksecuritygroups',
      'nsg': 'microsoft.network/networksecuritygroups',
      'network interface': 'microsoft.network/networkinterfaces',
      'nic': 'microsoft.network/networkinterfaces',
      'public ip address': 'microsoft.network/publicipaddresses',
      'public ip': 'microsoft.network/publicipaddresses',
      'public ip prefix': 'microsoft.network/publicipprefixes',
      'load balancer': 'microsoft.network/loadbalancers',
      'application gateway': 'microsoft.network/applicationgateways',
      'azure firewall': 'microsoft.network/azurefirewalls',
      'firewall': 'microsoft.network/azurefirewalls',
      'bastion': 'microsoft.network/bastionhosts',
      'bastion host': 'microsoft.network/bastionhosts',
      'vpn gateway': 'microsoft.network/vpngateways',
      'virtual network gateway': 'microsoft.network/virtualnetworkgateways',
      'expressroute circuit': 'microsoft.network/expressroutecircuits',
      'expressroute gateway': 'microsoft.network/expressroutegateways',
      'route table': 'microsoft.network/routetables',
      'dns zone': 'microsoft.network/dnszones',
      'private dns zone': 'microsoft.network/privatednszones',
      'private dns resolver': 'microsoft.network/privatednsresolver',
      'traffic manager profile': 'microsoft.network/trafficmanagerprofiles',
      'traffic manager': 'microsoft.network/trafficmanagerprofiles',
      'front door': 'microsoft.network/frontdoors',
      'azure front door': 'microsoft.network/frontdoors',
      'front door (classic)': 'microsoft.network/frontdoors',
      'ddos protection plan': 'microsoft.network/ddosprotectionplans',
      'private endpoint': 'microsoft.network/privateendpoints',
      'private link service': 'microsoft.network/privatelinkservices',
      'nat gateway': 'microsoft.network/natgateways',
      'ip group': 'microsoft.network/ipgroups',
      'network virtual appliance': 'microsoft.network/networkvirtualappliances',
      'network watcher': 'microsoft.network/networkwatchers',
      'connection': 'microsoft.network/connections',
      'waf policy': 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies',
      'web application firewall policy': 'microsoft.network/applicationgatewaywebapplicationfirewallpolicies',
      'firewall policy': 'microsoft.network/firewallpolicies',
      'application security group': 'microsoft.network/applicationsecuritygroups',
      'virtual wan': 'microsoft.network/virtualwans',
      'virtual hub': 'microsoft.network/virtualhubs',
      'local network gateway': 'microsoft.network/localnetworkgateways',
      'route filter': 'microsoft.network/routefilters',
      'service endpoint policy': 'microsoft.network/serviceendpointpolicies',
      'virtual router': 'microsoft.network/virtualrouters',
      'ip allocation': 'microsoft.network/ipallocations',

      // Storage
      'storage account': 'microsoft.storage/storageaccounts',
      'classic storage account': 'microsoft.classicstorage/storageaccounts',

      // SQL
      'sql server': 'microsoft.sql/servers',
      'sql database': 'microsoft.sql/servers/databases',
      'elastic pool': 'microsoft.sql/servers/elasticpools',
      'sql elastic pool': 'microsoft.sql/servers/elasticpools',
      'sql managed instance': 'microsoft.sql/managedinstances',
      'failover group': 'microsoft.sql/servers/failovergroups',
      'sql virtual machine': 'microsoft.sqlvirtualmachine/sqlvirtualmachines',
      'sql vm': 'microsoft.sqlvirtualmachine/sqlvirtualmachines',

      // MySQL / PostgreSQL / MariaDB
      'mysql server': 'microsoft.dbformysql/servers',
      'mysql flexible server': 'microsoft.dbformysql/flexibleservers',
      'azure database for mysql flexible server': 'microsoft.dbformysql/flexibleservers',
      'azure database for mysql server': 'microsoft.dbformysql/servers',
      'postgresql server': 'microsoft.dbforpostgresql/servers',
      'postgresql flexible server': 'microsoft.dbforpostgresql/flexibleservers',
      'azure database for postgresql flexible server': 'microsoft.dbforpostgresql/flexibleservers',
      'azure database for postgresql server': 'microsoft.dbforpostgresql/servers',
      'mariadb server': 'microsoft.dbformariadb/servers',
      'azure database for mariadb server': 'microsoft.dbformariadb/servers',

      // Cosmos DB
      'cosmos db account': 'microsoft.documentdb/databaseaccounts',
      'azure cosmos db account': 'microsoft.documentdb/databaseaccounts',
      'cosmosdb': 'microsoft.documentdb/databaseaccounts',
      'cosmos db': 'microsoft.documentdb/databaseaccounts',
      'azure cosmos db': 'microsoft.documentdb/databaseaccounts',

      // Cache
      'redis cache': 'microsoft.cache/redis',
      'azure cache for redis': 'microsoft.cache/redis',
      'redis enterprise': 'microsoft.cache/redisenterprise',

      // Web / App Service
      'app service': 'microsoft.web/sites',
      'web app': 'microsoft.web/sites',
      'function app': 'microsoft.web/sites',
      'app service plan': 'microsoft.web/serverfarms',
      'app service certificate': 'microsoft.web/certificates',
      'static web app': 'microsoft.web/staticsites',
      'app service environment': 'microsoft.web/hostingenvironments',

      // Containers
      'kubernetes service': 'microsoft.containerservice/managedclusters',
      'aks': 'microsoft.containerservice/managedclusters',
      'aks cluster': 'microsoft.containerservice/managedclusters',
      'azure kubernetes service': 'microsoft.containerservice/managedclusters',
      'container registry': 'microsoft.containerregistry/registries',
      'acr': 'microsoft.containerregistry/registries',
      'container instance': 'microsoft.containerinstance/containergroups',
      'container group': 'microsoft.containerinstance/containergroups',
      'container app': 'microsoft.app/containerapps',
      'container apps environment': 'microsoft.app/managedenvironments',
      'managed environment': 'microsoft.app/managedenvironments',

      // Key Vault / Identity
      'key vault': 'microsoft.keyvault/vaults',
      'managed hsm': 'microsoft.keyvault/managedhsms',
      'user assigned managed identity': 'microsoft.managedidentity/userassignedidentities',
      'managed identity': 'microsoft.managedidentity/userassignedidentities',

      // Monitoring
      'metric alert': 'microsoft.insights/metricalerts',
      'metric alert rule': 'microsoft.insights/metricalerts',
      'activity log alert': 'microsoft.insights/activitylogalerts',
      'activity log alert rule': 'microsoft.insights/activitylogalerts',
      'log alert rule': 'microsoft.insights/scheduledqueryrules',
      'scheduled query rule': 'microsoft.insights/scheduledqueryrules',
      'action group': 'microsoft.insights/actiongroups',
      'application insights': 'microsoft.insights/components',
      'app insights': 'microsoft.insights/components',
      'data collection rule': 'microsoft.insights/datacollectionrules',
      'autoscale setting': 'microsoft.insights/autoscalesettings',
      'workbook': 'microsoft.insights/workbooks',
      'web test': 'microsoft.insights/webtests',
      'availability test': 'microsoft.insights/webtests',
      'log analytics workspace': 'microsoft.operationalinsights/workspaces',
      'smart detection rule': 'microsoft.alertsmanagement/smartdetectoralertrules',
      'smart detector alert rule': 'microsoft.alertsmanagement/smartdetectoralertrules',

      // Automation / Logic
      'automation account': 'microsoft.automation/automationaccounts',
      'logic app': 'microsoft.logic/workflows',
      'integration account': 'microsoft.logic/integrationaccounts',

      // Messaging
      'service bus namespace': 'microsoft.servicebus/namespaces',
      'service bus': 'microsoft.servicebus/namespaces',
      'event hub namespace': 'microsoft.eventhub/namespaces',
      'event hubs namespace': 'microsoft.eventhub/namespaces',
      'event hub cluster': 'microsoft.eventhub/clusters',
      'event hubs cluster': 'microsoft.eventhub/clusters',
      'relay namespace': 'microsoft.relay/namespaces',
      'notification hub namespace': 'microsoft.notificationhubs/namespaces',
      'notification hub': 'microsoft.notificationhubs/namespaces/notificationhubs',

      // AI / Cognitive
      'cognitive services account': 'microsoft.cognitiveservices/accounts',
      'cognitive services': 'microsoft.cognitiveservices/accounts',
      'azure ai services': 'microsoft.cognitiveservices/accounts',
      'azure openai': 'microsoft.cognitiveservices/accounts',
      'openai service': 'microsoft.cognitiveservices/accounts',
      'azure openai service': 'microsoft.cognitiveservices/accounts',
      'machine learning workspace': 'microsoft.machinelearningservices/workspaces',
      'azure ml workspace': 'microsoft.machinelearningservices/workspaces',
      'azure machine learning': 'microsoft.machinelearningservices/workspaces',

      // Search / SignalR / API Management
      'search service': 'microsoft.search/searchservices',
      'azure cognitive search': 'microsoft.search/searchservices',
      'azure ai search': 'microsoft.search/searchservices',
      'signalr': 'microsoft.signalrservice/signalr',
      'signalr service': 'microsoft.signalrservice/signalr',
      'api management': 'microsoft.apimanagement/service',
      'api management service': 'microsoft.apimanagement/service',
      'apim': 'microsoft.apimanagement/service',

      // CDN / Media
      'cdn profile': 'microsoft.cdn/profiles',
      'cdn endpoint': 'microsoft.cdn/profiles/endpoints',
      'front door profile': 'microsoft.cdn/profiles',
      'front door standard/premium': 'microsoft.cdn/profiles',
      'media services': 'microsoft.media/mediaservices',
      'media service': 'microsoft.media/mediaservices',

      // Backup / Recovery
      'batch account': 'microsoft.batch/batchaccounts',
      'recovery services vault': 'microsoft.recoveryservices/vaults',
      'backup vault': 'microsoft.dataprotection/backupvaults',

      // Migration
      'azure migrate project': 'microsoft.migrate/projects',
      'migrate project': 'microsoft.migrate/projects',
      'database migration service': 'microsoft.datamigration/services',

      // Data / Analytics
      'data factory': 'microsoft.datafactory/factories',
      'azure data factory': 'microsoft.datafactory/factories',
      'databricks workspace': 'microsoft.databricks/workspaces',
      'azure databricks': 'microsoft.databricks/workspaces',
      'synapse workspace': 'microsoft.synapse/workspaces',
      'azure synapse analytics': 'microsoft.synapse/workspaces',
      'hdinsight cluster': 'microsoft.hdinsight/clusters',
      'analysis services server': 'microsoft.analysisservices/servers',
      'stream analytics job': 'microsoft.streamanalytics/streamingjobs',
      'data lake analytics': 'microsoft.datalakeanalytics/accounts',
      'data lake store': 'microsoft.datalakestore/accounts',
      'data share account': 'microsoft.datashare/accounts',
      'data catalog': 'microsoft.datacatalog/catalogs',

      // IoT
      'iot hub': 'microsoft.devices/iothubs',
      'device provisioning service': 'microsoft.devices/provisioningservices',
      'iot hub device provisioning service': 'microsoft.devices/provisioningservices',
      'iot central application': 'microsoft.iotcentral/iotapps',
      'iot central': 'microsoft.iotcentral/iotapps',
      'time series insights environment': 'microsoft.timeseriesinsights/environments',

      // Other
      'azure maps account': 'microsoft.maps/accounts',
      'dashboard': 'microsoft.portal/dashboards',
      'power bi embedded': 'microsoft.powerbidedicated/capacities',
      'purview account': 'microsoft.purview/accounts',
      'microsoft purview': 'microsoft.purview/accounts',
      'data explorer cluster': 'microsoft.kusto/clusters',
      'azure data explorer': 'microsoft.kusto/clusters',
      'service fabric cluster': 'microsoft.servicefabric/clusters',
      'app configuration': 'microsoft.appconfiguration/configurationstores',
      'event grid topic': 'microsoft.eventgrid/topics',
      'event grid domain': 'microsoft.eventgrid/domains',
      'event grid system topic': 'microsoft.eventgrid/systemtopics',
      'communication service': 'microsoft.communication/communicationservices',
      'azure communication services': 'microsoft.communication/communicationservices',
      'azure api for fhir': 'microsoft.healthcareapis/services',
      'healthcare apis': 'microsoft.healthcareapis/services',
      'digital twins': 'microsoft.digitaltwins/digitaltwinsinstances',
      'azure digital twins': 'microsoft.digitaltwins/digitaltwinsinstances',
      'bot service': 'microsoft.botservice/botservices',
      'azure bot': 'microsoft.botservice/botservices',
      'maintenance configuration': 'microsoft.maintenance/maintenanceconfigurations',
      'host pool': 'microsoft.desktopvirtualization/hostpools',
      'application group': 'microsoft.desktopvirtualization/applicationgroups',
      'avd workspace': 'microsoft.desktopvirtualization/workspaces',
      'virtual desktop workspace': 'microsoft.desktopvirtualization/workspaces',
      'azure netapp files': 'microsoft.netapp/netappaccounts',
      'netapp account': 'microsoft.netapp/netappaccounts',
      'spring app': 'microsoft.appplatform/spring',
      'azure spring apps': 'microsoft.appplatform/spring',
      'classic virtual machine': 'microsoft.classiccompute/virtualmachines',
      'classic virtual network': 'microsoft.classicnetwork/virtualnetworks',
      'resource group': 'microsoft.resources/resourcegroups',
      'load test': 'microsoft.loadtestservice/loadtests',
      'azure load testing': 'microsoft.loadtestservice/loadtests',
      'azure vmware solution': 'microsoft.avs/privateclouds',
      'dev test lab': 'microsoft.devtestlab/labs',
      'devtest lab': 'microsoft.devtestlab/labs',
      'domain': 'microsoft.domainregistration/domains',
      'certificate order': 'microsoft.certificateregistration/certificateorders',
      'arc machine': 'microsoft.hybridcompute/machines',
      'arc-enabled server': 'microsoft.hybridcompute/machines',
      'storage sync service': 'microsoft.storagesync/storagesyncservices',
      'azure red hat openshift': 'microsoft.redhatopenshift/openshiftclusters',
      'aro cluster': 'microsoft.redhatopenshift/openshiftclusters',
    };
    return this.__displayNameMap;
  }

  /**
   * Try to find a matching rule using partial/fuzzy ARM type matching.
   * Handles cases where the exact type isn't in the matrix but a parent or variant is.
   */
  _fuzzyMatchRule(normalizedType) {
    // Exact match
    if (this.rules[normalizedType]) {
      return this.rules[normalizedType];
    }

    // Try parent type (e.g., microsoft.sql/servers/databases → microsoft.sql/servers)
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

    // Try suffix match (e.g., if someone provides just "virtualmachines" without provider)
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

  assessType(rawType) {
    const normalized = this.normalizeType(rawType);
    const rule = this._fuzzyMatchRule(normalized);

    if (rule) {
      // Priority: static JSON remarks > subscription remarks map > Learn page remarks > generated fallback
      const actionableRemark = this._subscriptionRemarks[normalized]
        || rule.remarks
        || this._generateSubscriptionRemark(normalized, rule.subscriptionMove);
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

  /**
   * Assess region move support for a given resource type.
   */
  assessRegionType(rawType) {
    const normalized = this.normalizeType(rawType);

    // Check region rules (from CSV)
    if (this.regionRules[normalized]) {
      const regionMove = this.regionRules[normalized].regionMove;
      const remark = this._regionRemarks[normalized] || this._generateRegionRemark(normalized, regionMove);
      return {
        originalType: rawType,
        normalizedType: normalized,
        regionMove,
        remarks: remark
      };
    }

    // Try parent type match for region rules
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
   * Actionable subscription-move remarks — resource-specific migration guidance.
   * These provide concrete steps (detach, re-attach, prerequisites, tools) not just "Yes/No".
   */
  get _subscriptionRemarks() {
    if (this.__subscriptionRemarks) return this.__subscriptionRemarks;
    this.__subscriptionRemarks = {
      // ── Compute ──
      'microsoft.compute/virtualmachines': 'Can be moved directly. Pre-checks: (1) Detach any Azure Backup vault before move. (2) Move all dependent NICs, disks, and public IPs together. (3) VMs in availability sets must be moved as a set. Use Azure Resource Mover or az resource move CLI.',
      'microsoft.compute/virtualmachines/extensions': 'VM extensions move with the parent VM automatically. No separate action needed.',
      'microsoft.compute/availabilitysets': 'Move all VMs in the availability set together in a single move operation. Cannot move individual VMs separately from the set.',
      'microsoft.compute/disks': 'Managed disks can be moved directly between subscriptions. Ensure the VM is deallocated first if the disk is attached.',
      'microsoft.compute/snapshots': 'Full snapshots: movable directly. Incremental snapshots: CANNOT be moved across subscriptions. Workaround: Create a full snapshot from the incremental one, then move the full snapshot.',
      'microsoft.compute/images': 'Custom images can be moved directly between subscriptions. Alternatively, use Shared Image Gallery for cross-subscription sharing.',
      'microsoft.compute/galleries': 'Shared Image Galleries CANNOT be moved. Workaround: Replicate images to a new gallery in the target subscription, or share via RBAC across subscriptions.',
      'microsoft.compute/galleries/images': 'Gallery images CANNOT be moved. Workaround: Re-create images in a new gallery in the target subscription.',
      'microsoft.compute/galleries/images/versions': 'Image versions CANNOT be moved. Workaround: Copy VHDs and recreate versions in a new gallery.',
      'microsoft.compute/diskencryptionsets': 'CANNOT be moved. Workaround: Create a new disk encryption set in the target subscription linked to the Key Vault key, then re-encrypt disks.',
      'microsoft.compute/diskaccesses': 'CANNOT be moved across subscriptions. Recreate the disk access resource in the target subscription.',
      'microsoft.compute/hostgroups': 'Dedicated host groups CANNOT be moved. Workaround: Create a new host group in the target subscription and redeploy VMs.',
      'microsoft.compute/hostgroups/hosts': 'Dedicated hosts CANNOT be moved. Recreate hosts in a new host group in the target subscription.',
      'microsoft.compute/proximityplacementgroups': 'Can be moved directly. Move all associated VMs and availability sets together.',
      'microsoft.compute/sshpublickeys': 'CANNOT be moved. Recreate SSH public key resources in the target subscription.',
      'microsoft.compute/restorepointcollections': 'CANNOT be moved. Create new restore point collections in the target subscription after moving VMs.',
      'microsoft.compute/virtualmachinescalesets': 'Can be moved between subscriptions. Move all dependent resources (VNets, load balancers, public IPs) together.',
      'microsoft.compute/capacityreservationgroups': 'Check the current move support. May require recreating in the target subscription.',
      'microsoft.compute/cloudservices': 'Cloud Services (extended support) can be moved. Ensure all dependencies are in the same move batch.',

      // ── Network ──
      'microsoft.network/publicipaddresses': 'Basic SKU: Can be moved directly between subscriptions. Standard SKU: MUST detach from the associated resource (VM/NIC/LB) first, then move the IP separately, then re-attach in the target subscription. See https://learn.microsoft.com/en-us/azure/virtual-network/move-across-regions-publicip-portal',
      'microsoft.network/loadbalancers': 'Basic SKU: Can be moved directly. Standard SKU: CANNOT be moved. Workaround for Standard: (1) Detach backend pool members, (2) Delete the LB, (3) Recreate in the target subscription using ARM template export, (4) Re-attach backends.',
      'microsoft.network/virtualnetworks': 'Can be moved but ALL dependent resources (VMs, NICs, NSGs, route tables) must be moved together. Peered VNets: Disable peering first, move, then re-establish peering.',
      'microsoft.network/networksecuritygroups': 'Can be moved directly between subscriptions. Move with associated VNets/NICs for consistency.',
      'microsoft.network/networkinterfaces': 'Can be moved. Must move together with the parent VM. Ensure NSG and public IP associations are handled.',
      'microsoft.network/applicationgateways': 'Can be moved between subscriptions. Move the VNet and all dependent resources together. WAF policies move separately.',
      'microsoft.network/azurefirewalls': 'CANNOT be moved. Workaround: Export ARM template, delete firewall, recreate in target subscription with the exported config. Firewall policy can be moved separately.',
      'microsoft.network/bastionhosts': 'CANNOT be moved. Workaround: Delete and recreate Bastion in the target subscription VNet.',
      'microsoft.network/vpngateways': 'CANNOT be moved. Workaround: Delete the gateway, move the VNet, recreate the gateway in the target subscription, and reconfigure VPN tunnels.',
      'microsoft.network/virtualnetworkgateways': 'CANNOT be moved across subscriptions. Workaround: Delete gateway, move VNet, recreate gateway and connections.',
      'microsoft.network/expressroutecircuits': 'Can be moved between subscriptions. Circuit peerings and authorizations move with the circuit.',
      'microsoft.network/routetables': 'Can be moved directly. Move with the associated VNet.',
      'microsoft.network/dnszones': 'Can be moved between subscriptions. DNS records remain intact after move.',
      'microsoft.network/privatednszones': 'Can be moved. VNet links may need to be updated post-move.',
      'microsoft.network/trafficmanagerprofiles': 'Can be moved directly between subscriptions. Endpoints remain configured.',
      'microsoft.network/frontdoors': 'Can be moved between subscriptions directly.',
      'microsoft.network/ddosprotectionplans': 'Can be moved. Dissociate from VNets first, move, then re-associate in target subscription.',
      'microsoft.network/privateendpoints': 'Supported for specific private-link resource types only. Check Learn docs for your specific resource. May need to recreate if the linked resource changes subscription.',
      'microsoft.network/privatelinkservices': 'CANNOT be moved. Recreate Private Link service in the target subscription.',
      'microsoft.network/natgateways': 'Can be moved between subscriptions. Move with associated subnet/VNet.',
      'microsoft.network/firewallpolicies': 'Can be moved between subscriptions. Move before or with the firewall.',
      'microsoft.network/applicationsecuritygroups': 'Can be moved directly between subscriptions.',
      'microsoft.network/networkwatchers': 'CANNOT be moved. Network Watcher is auto-created per subscription/region. No action needed.',
      'microsoft.network/connections': 'Can be moved directly. Move together with associated gateways.',
      'microsoft.network/localnetworkgateways': 'Can be moved directly between subscriptions.',
      'microsoft.network/virtualwans': 'CANNOT be moved. Recreate Virtual WAN and hubs in the target subscription.',
      'microsoft.network/ipgroups': 'Can be moved between subscriptions.',
      'microsoft.network/serviceendpointpolicies': 'Can be moved directly between subscriptions.',
      'microsoft.network/publicipprefixes': 'Can be moved directly between subscriptions.',
      'microsoft.network/virtualrouters': 'Can be moved between subscriptions.',

      // ── SQL ──
      'microsoft.sql/servers': 'Can be moved between subscriptions. All databases under the server move with it. Ensure firewall rules and AAD admin config are updated post-move.',
      'microsoft.sql/servers/databases': 'Databases move with the parent SQL server. Cannot move individual databases independently.',
      'microsoft.sql/servers/elasticpools': 'Elastic pools move with the parent SQL server.',
      'microsoft.sql/managedinstances': 'Can be moved between subscriptions within the same region. Requires specific permissions: Microsoft.Sql/managedInstances/write.',
      'microsoft.sql/servers/failovergroups': 'CANNOT be moved directly. Remove failover group, move servers, then recreate the failover group.',
      'microsoft.sql/virtualclusters': 'Move with the managed instance.',

      // ── MySQL / PostgreSQL / MariaDB ──
      'microsoft.dbformysql/servers': 'Can be moved between subscriptions. VNet rules and private endpoints may need reconfiguration post-move.',
      'microsoft.dbformysql/flexibleservers': 'Can be moved between subscriptions. Update VNet integration and private DNS zone links post-move.',
      'microsoft.dbforpostgresql/servers': 'Can be moved between subscriptions. Reconfigure VNet rules post-move.',
      'microsoft.dbforpostgresql/flexibleservers': 'Can be moved between subscriptions. Update VNet integration post-move.',
      'microsoft.dbforpostgresql/serversv2': 'Can be moved between subscriptions.',
      'microsoft.dbformariadb/servers': 'Can be moved between subscriptions.',

      // ── Cosmos DB ──
      'microsoft.documentdb/databaseaccounts': 'Can be moved between subscriptions. All databases, containers, and throughput settings are preserved.',

      // ── Storage ──
      'microsoft.storage/storageaccounts': 'Can be moved between subscriptions. Private endpoints, VNet rules, and RBAC assignments must be updated post-move. Data remains intact.',

      // ── Key Vault ──
      'microsoft.keyvault/vaults': 'Can be moved but with CRITICAL prerequisites: (1) If used with disk encryption, disable encryption first. (2) Soft-delete must be enabled. (3) Update access policies for the new subscription tenant if different. See https://learn.microsoft.com/en-us/azure/key-vault/general/move-subscription',
      'microsoft.keyvault/managedhsms': 'Check current support. May require recreating in the target subscription.',

      // ── Web / App Service ──
      'microsoft.web/sites': 'Can be moved between subscriptions within the SAME region. (1) Move the App Service Plan first or together. (2) Custom domains and SSL certs must be reconfigured. (3) Deployment slots move with the app. See https://learn.microsoft.com/en-us/azure/app-service/app-service-move-limitations',
      'microsoft.web/serverfarms': 'App Service Plans can be moved between subscriptions within the same region. All apps in the plan must be moved together.',
      'microsoft.web/certificates': 'Certificates CANNOT be moved directly. Workaround: Delete and recreate the certificate binding in the target subscription.',
      'microsoft.web/staticsites': 'Can be moved between subscriptions. Static Web Apps are global.',
      'microsoft.web/hostingenvironments': 'App Service Environment (ASE) can be moved between subscriptions. All apps and plans in the ASE must move together.',
      'microsoft.web/connectiongateways': 'Can be moved between subscriptions.',
      'microsoft.web/sites/premieraddons': 'Move with the parent App Service app.',

      // ── Functions ──
      'microsoft.web/sites/slots': 'Deployment slots move with the parent app. No separate action needed.',

      // ── Containers ──
      'microsoft.containerservice/managedclusters': 'AKS clusters CANNOT be moved across subscriptions. Workaround: (1) Export workload manifests, (2) Create a new AKS cluster in the target subscription, (3) Redeploy workloads, (4) Update DNS and traffic routing.',
      'microsoft.containerregistry/registries': 'Can be moved between subscriptions. Image data is preserved. Update webhook URLs and VNet rules post-move.',
      'microsoft.containerinstance/containergroups': 'CANNOT be moved. Recreate container groups in the target subscription.',
      'microsoft.app/containerapps': 'Can be moved between subscriptions. Check managed environment compatibility.',
      'microsoft.app/managedenvironments': 'Can be moved between subscriptions.',

      // ── Monitoring & Insights ──
      'microsoft.insights/components': 'Application Insights can be moved between subscriptions. Instrumentation key remains the same. Update connection strings if workspace changes.',
      'microsoft.insights/actiongroups': 'Can be moved between subscriptions.',
      'microsoft.insights/metricalerts': 'CANNOT be moved directly. Workaround: Delete alert rules, move target resources, recreate alerts in target subscription.',
      'microsoft.insights/activitylogalerts': 'Can be moved between subscriptions.',
      'microsoft.insights/scheduledqueryrules': 'Can be moved between subscriptions.',
      'microsoft.insights/autoscalesettings': 'Can be moved between subscriptions. Verify target resource references post-move.',
      'microsoft.insights/workbooks': 'Can be moved between subscriptions.',
      'microsoft.insights/webtests': 'Can be moved between subscriptions.',
      'microsoft.insights/datacollectionrules': 'Can be moved between subscriptions.',
      'microsoft.operationalinsights/workspaces': 'Can be moved between subscriptions. Connected data sources and solutions are preserved. Update agent configurations if needed.',

      // ── AI & Cognitive Services ──
      'microsoft.cognitiveservices/accounts': 'Can be moved between subscriptions. API keys remain the same. Update endpoint URLs in applications.',
      'microsoft.machinelearningservices/workspaces': 'Can be moved between subscriptions. Compute targets and endpoints may need reconfiguration.',
      'microsoft.search/searchservices': 'Can be moved between subscriptions. Indexes and data are preserved.',

      // ── Messaging ──
      'microsoft.servicebus/namespaces': 'Can be moved between subscriptions. Queues, topics, and subscriptions are preserved.',
      'microsoft.eventhub/namespaces': 'Can be moved between subscriptions. Event hubs and consumer groups are preserved.',
      'microsoft.eventhub/clusters': 'Check current move support. Dedicated clusters may require recreation.',
      'microsoft.relay/namespaces': 'Can be moved between subscriptions.',
      'microsoft.notificationhubs/namespaces': 'Can be moved between subscriptions.',

      // ── Integration ──
      'microsoft.logic/workflows': 'Can be moved between subscriptions. API connections must be recreated or moved separately. Update managed connector references post-move.',
      'microsoft.logic/integrationaccounts': 'Can be moved between subscriptions.',
      'microsoft.automation/automationaccounts': 'Can be moved between subscriptions. Runbooks, schedules, and variables are preserved. Update Run As account if used.',

      // ── API Management ──
      'microsoft.apimanagement/service': 'Can be moved between subscriptions. Custom domains, certificates, and named values are preserved.',

      // ── Data & Analytics ──
      'microsoft.datafactory/factories': 'Can be moved between subscriptions. Pipelines, datasets, and linked services are preserved. Self-hosted IRs may need reconfiguration.',
      'microsoft.databricks/workspaces': 'Can be moved between subscriptions. Notebooks, jobs, and clusters are preserved. Update VNet injection config if used.',
      'microsoft.synapse/workspaces': 'Can be moved between subscriptions. Linked services and pipelines are preserved.',
      'microsoft.hdinsight/clusters': 'CANNOT be moved across subscriptions. Workaround: Delete cluster, move storage, recreate cluster in target subscription.',
      'microsoft.kusto/clusters': 'Can be moved between subscriptions.',
      'microsoft.streamanalytics/streamingjobs': 'Can be moved between subscriptions. Update input/output connections post-move.',
      'microsoft.datalakestore/accounts': 'Can be moved between subscriptions.',
      'microsoft.datalakeanalytics/accounts': 'Can be moved between subscriptions.',
      'microsoft.purview/accounts': 'Can be moved between subscriptions.',
      'microsoft.datashare/accounts': 'Can be moved between subscriptions.',
      'microsoft.analysisservices/servers': 'Can be moved between subscriptions.',

      // ── IoT ──
      'microsoft.devices/iothubs': 'Can be moved between subscriptions. Device identities and routes are preserved. Update DPS enrollment if linked.',
      'microsoft.devices/provisioningservices': 'Can be moved between subscriptions.',
      'microsoft.iotcentral/iotapps': 'Can be moved between subscriptions.',
      'microsoft.timeseriesinsights/environments': 'Can be moved between subscriptions.',

      // ── Identity & Security ──
      'microsoft.managedidentity/userassignedidentities': 'Can be moved between subscriptions. CRITICAL: Update all RBAC role assignments post-move as the resource ID changes.',
      'microsoft.aad/domainservices': 'Azure AD DS CANNOT be moved. Workaround: Delete and recreate in the target subscription/VNet. This is destructive—back up domain data first.',

      // ── Recovery & Backup ──
      'microsoft.recoveryservices/vaults': 'Recovery Services vaults can be moved but with constraints: (1) Cannot move if vault is protecting Azure VMs with Backup. (2) Stop backup and delete data, or use Azure Backup to restore in target. See https://learn.microsoft.com/en-us/azure/backup/backup-azure-move-recovery-services-vault',
      'microsoft.dataprotection/backupvaults': 'Check current move support. May require stopping protection first.',

      // ── CDN ──
      'microsoft.cdn/profiles': 'Can be moved between subscriptions. CDN endpoints and custom domains are preserved.',
      'microsoft.cdn/profiles/endpoints': 'CDN endpoints move with the parent profile.',

      // ── DevOps ──
      'microsoft.devtestlab/labs': 'Can be moved between subscriptions. VMs and artifacts within the lab are preserved.',

      // ── Misc ──
      'microsoft.signalrservice/signalr': 'Can be moved between subscriptions.',
      'microsoft.portal/dashboards': 'Can be moved between subscriptions. Dashboards are global.',
      'microsoft.appconfiguration/configurationstores': 'Can be moved between subscriptions. Configuration data is preserved.',
      'microsoft.communication/communicationservices': 'Can be moved but resources with attached phone numbers CANNOT be moved to subscriptions in different data locations.',
      'microsoft.eventgrid/topics': 'Can be moved between subscriptions.',
      'microsoft.eventgrid/domains': 'Can be moved between subscriptions.',
      'microsoft.eventgrid/systemtopics': 'Can be moved between subscriptions.',
      'microsoft.batch/batchaccounts': 'Can be moved between subscriptions.',
      'microsoft.maps/accounts': 'Can be moved between subscriptions. Maps is a global service.',
      'microsoft.media/mediaservices': 'Can be moved between subscriptions.',
      'microsoft.desktopvirtualization/hostpools': 'Can be moved between subscriptions. Move all application groups and workspaces together.',
      'microsoft.desktopvirtualization/applicationgroups': 'Move with the parent host pool.',
      'microsoft.desktopvirtualization/workspaces': 'Can be moved between subscriptions.',
      'microsoft.netapp/netappaccounts': 'CANNOT be moved across subscriptions. Workaround: Replicate data using cross-region replication, recreate in target subscription.',
      'microsoft.cache/redis': 'Can be moved between subscriptions. Data persistence and firewall rules are preserved.',
      'microsoft.cache/redisenterprise': 'Check current move support. May require recreation.',
      'microsoft.botservice/botservices': 'Can be moved between subscriptions.',
      'microsoft.digitaltwins/digitaltwinsinstances': 'Can be moved between subscriptions.',
      'microsoft.migrate/projects': 'Can be moved between subscriptions.',
      'microsoft.hybridcompute/machines': 'Can be moved between subscriptions.',
      'microsoft.redhatopenshift/openshiftclusters': 'CANNOT be moved. Recreate the ARO cluster in the target subscription.',
      'microsoft.servicefabric/clusters': 'CANNOT be moved. Recreate the cluster and redeploy applications in the target subscription.',
      'microsoft.appplatform/spring': 'Can be moved between subscriptions.',
      'microsoft.loadtestservice/loadtests': 'Can be moved between subscriptions.',
      'microsoft.avs/privateclouds': 'CANNOT be moved. Recreate the private cloud in the target subscription.',
      'microsoft.healthcareapis/services': 'Can be moved between subscriptions.',
      'microsoft.sqlvirtualmachine/sqlvirtualmachines': 'Can be moved. Move the underlying VM first, then the SQL VM resource follows. Re-register with SQL IaaS Agent extension if needed.',
      'microsoft.storagesync/storagesyncservices': 'Can be moved between subscriptions.',
      'microsoft.powerbidedicated/capacities': 'Can be moved between subscriptions.',
    };
    return this.__subscriptionRemarks;
  }

  /**
   * Generate a fallback subscription-move remark based on the provider and move status.
   */
  _generateSubscriptionRemark(normalizedType, subscriptionMove) {
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

    // Generic fallback
    if (subscriptionMove === 'Yes') {
      return `Can be moved between subscriptions. Use az resource move CLI or Azure portal. Verify dependent resources are included.`;
    }
    if (subscriptionMove === 'No') {
      return `CANNOT be moved between subscriptions. Workaround: Export ARM template, recreate ${resource} in the target subscription. See https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources`;
    }
    return `Conditional move — check specific prerequisites for ${resource} on https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources`;
  }

  /**
   * Generate a default remark based on the resource provider and move status when no specific remark exists.
   */
  _generateRegionRemark(normalizedType, regionMove) {
    // Extract a friendly service name from the ARM type
    const parts = normalizedType.split('/');
    const provider = (parts[0] || '').replace('microsoft.', '');
    const resource = parts.slice(1).join('/');

    // Provider-level guidance map
    const providerGuidance = {
      'compute': { yes: 'Use Azure Resource Mover for cross-region migration.', no: 'Region move not supported. Redeploy in the target region using ARM/Bicep templates.' },
      'network': { yes: 'Can be moved across regions using Azure Resource Mover. Network config is recreated in the target region.', no: 'Region move not supported. Recreate the networking resource in the target region.' },
      'storage': { yes: 'Create a new storage account in the target region and use AzCopy to migrate data.', no: 'Region move not supported. Create new storage in target region and copy data using AzCopy.' },
      'sql': { yes: 'Use active geo-replication or failover groups for cross-region move.', no: 'Region move not supported. Use backup/restore or geo-replication to migrate.' },
      'dbformysql': { yes: 'Use read replicas or backup/restore for cross-region migration.', no: 'Region move not supported. Use backup/restore or create a new server in the target region.' },
      'dbforpostgresql': { yes: 'Use read replicas or backup/restore for cross-region migration.', no: 'Region move not supported. Use backup/restore or create a new server in the target region.' },
      'dbformariadb': { yes: 'Use backup/restore for cross-region migration.', no: 'Region move not supported. Use backup/restore to recreate in the target region.' },
      'web': { yes: 'Redeploy to the target region using ARM templates, CI/CD pipelines, or backup/restore.', no: 'Region move not supported. Redeploy using ARM templates, CI/CD, or backup/restore.' },
      'containerservice': { yes: 'Use Azure Resource Mover or redeploy cluster in the target region.', no: 'Region move not supported. Redeploy cluster and workloads in the target region.' },
      'containerregistry': { yes: 'Use geo-replication to replicate images to the target region.', no: 'Region move not supported. Use geo-replication or create a new registry in the target region.' },
      'containerinstance': { yes: 'Redeploy container group in the target region.', no: 'Region move not supported. Redeploy the container group in the target region.' },
      'app': { yes: 'Redeploy Container App in the target region.', no: 'Region move not supported. Redeploy Container App and environment in the target region.' },
      'keyvault': { yes: 'Move vault to target region.', no: 'Region move not supported. Create a new vault in the target region and migrate secrets/keys/certificates.' },
      'documentdb': { yes: 'Use multi-region writes or create account in the new region and migrate data.', no: 'Region move not supported. Add target region as a replica, then remove old region.' },
      'cache': { yes: 'Create new cache in target region and migrate data.', no: 'Region move not supported. Create a new cache instance and migrate data.' },
      'insights': { yes: 'Recreate monitoring resource in the target region and reconfigure.', no: 'Region move not supported. Create a new resource in the target region and reconfigure instrumentation.' },
      'operationalinsights': { yes: 'Create a new workspace in the target region.', no: 'Region move not supported. Create a new workspace in the target region.' },
      'cognitiveservices': { yes: 'Redeploy Cognitive Services resource in the target region.', no: 'Region move not supported. Create a new resource in the target region and update application endpoints.' },
      'machinelearningservices': { yes: 'Redeploy workspace in the target region.', no: 'Region move not supported. Create a new workspace in the target region and retrain/redeploy models.' },
      'search': { yes: 'Create a new search service in the target region and rebuild indexes.', no: 'Region move not supported. Create a new search service and rebuild indexes.' },
      'servicebus': { yes: 'Use geo-disaster recovery pairing for cross-region move.', no: 'Region move not supported. Use geo-disaster recovery or recreate namespace in the target region.' },
      'eventhub': { yes: 'Use geo-disaster recovery pairing for cross-region move.', no: 'Region move not supported. Use geo-disaster recovery or recreate namespace.' },
      'devices': { yes: 'Use manual failover for cross-region migration.', no: 'Region move not supported. Create a new IoT Hub in the target region and re-register devices.' },
      'recoveryservices': { yes: 'Move vault to the target region.', no: 'Region move not supported. Create a new vault in the target region and re-protect workloads.' },
      'datafactory': { yes: 'Export and re-import pipelines in a new factory in the target region.', no: 'Region move not supported. Create a new factory in the target region and re-import pipelines.' },
      'databricks': { yes: 'Create a new workspace in the target region.', no: 'Region move not supported. Create a new workspace in the target region and migrate notebooks/jobs.' },
      'synapse': { yes: 'Create a new workspace in the target region.', no: 'Region move not supported. Create new workspace in the target region and migrate artifacts.' },
      'logic': { yes: 'Export and re-import Logic App definition in the target region.', no: 'Region move not supported. Export and re-import Logic App definition in the target region.' },
      'automation': { yes: 'Create a new automation account in the target region and migrate runbooks.', no: 'Region move not supported. Create a new account and migrate runbooks/configurations.' },
      'apimanagement': { yes: 'Use backup/restore for cross-region migration.', no: 'Region move not supported. Use backup/restore to migrate API Management instance.' },
      'signalrservice': { yes: 'Create a new SignalR instance in the target region.', no: 'Region move not supported. Create a new instance in the target region.' },
      'cdn': { yes: 'CDN is a global service; update endpoint origins to the new region.', no: 'CDN is a global service. Update origins to point to resources in the new region.' },
      'batch': { yes: 'Create a new Batch account in the target region.', no: 'Region move not supported. Create a new Batch account in the target region.' },
      'media': { yes: 'Create a new Media Services account in the target region.', no: 'Region move not supported. Create a new Media Services account in the target region.' },
      'hdinsight': { yes: 'Create a new HDInsight cluster in the target region.', no: 'Region move not supported. Create a new cluster in the target region.' },
      'kusto': { yes: 'Create a new Data Explorer cluster in the target region.', no: 'Region move not supported. Create a new cluster in the target region and migrate databases.' },
      'purview': { yes: 'Create a new Purview account in the target region.', no: 'Region move not supported. Create a new account in the target region.' },
      'eventgrid': { yes: 'Recreate Event Grid resource in the target region.', no: 'Region move not supported. Recreate topics/domains in the target region.' },
      'relay': { yes: 'Create a new Relay namespace in the target region.', no: 'Region move not supported. Create a new namespace in the target region.' },
      'notificationhubs': { yes: 'Create a new Notification Hub in the target region.', no: 'Region move not supported. Create new namespace and hub in the target region.' },
      'maps': { yes: 'Azure Maps is a global service; no region move required.', no: 'Azure Maps is a global service; no region move typically needed.' },
      'managedidentity': { yes: 'Recreate the managed identity in the target region.', no: 'Region move not supported. Create a new managed identity and update role assignments.' },
      'portal': { yes: 'Dashboards are global; no region move needed.', no: 'Dashboards are global; no region move typically needed.' },
      'appconfiguration': { yes: 'Create new App Configuration store in the target region and export/import settings.', no: 'Region move not supported. Create new store and export/import configuration.' },
      'migrate': { yes: 'Create a new project in the target region.', no: 'Region move not supported. Create a new project in the target region.' },
      'datamigration': { yes: 'Create a new DMS instance in the target region.', no: 'Region move not supported. Create a new DMS instance in the target region.' },
      'dataprotection': { yes: 'Create a new backup vault in the target region.', no: 'Region move not supported. Create a new backup vault and re-configure backup policies.' },
      'botservice': { yes: 'Redeploy bot in the target region.', no: 'Region move not supported. Redeploy bot service in the target region.' },
      'communication': { yes: 'Create a new Communication Services resource in the target region.', no: 'Region move not supported. Create a new resource in the target region.' },
      'digitaltwins': { yes: 'Create a new Digital Twins instance in the target region.', no: 'Region move not supported. Create a new instance and migrate models/twins.' },
      'iotcentral': { yes: 'Create a new IoT Central application in the target region.', no: 'Region move not supported. Create a new application in the target region.' },
      'timeseriesinsights': { yes: 'Create a new TSI environment in the target region.', no: 'Region move not supported. Create a new environment in the target region.' },
      'netapp': { yes: 'Use cross-region replication to migrate volumes.', no: 'Region move not supported. Use cross-region replication or create new account in target region.' },
      'desktopvirtualization': { yes: 'Recreate host pool and session hosts in the target region.', no: 'Region move not supported. Recreate host pool and session hosts in the target region.' },
      'appplatform': { yes: 'Create a new Spring Apps instance in the target region.', no: 'Region move not supported. Create new instance in the target region and redeploy apps.' },
      'maintenance': { yes: 'Recreate maintenance configuration in the target region.', no: 'Region move not supported. Recreate configuration in the target region.' },
      'loadtestservice': { yes: 'Create a new load test resource in the target region.', no: 'Region move not supported. Create a new load test resource in the target region.' },
      'hybridcompute': { yes: 'Re-register Arc machines in the target region.', no: 'Region move not supported. Re-register Arc-enabled servers in the target region.' },
      'redhatopenshift': { yes: 'Create a new ARO cluster in the target region.', no: 'Region move not supported. Create a new ARO cluster in the target region.' },
      'servicefabric': { yes: 'Create a new Service Fabric cluster in the target region.', no: 'Region move not supported. Create a new cluster and redeploy applications.' },
      'powerbidedicated': { yes: 'Create a new Power BI Embedded capacity in the target region.', no: 'Region move not supported. Create a new capacity in the target region.' },
      'storagesync': { yes: 'Create a new Storage Sync Service in the target region.', no: 'Region move not supported. Create a new sync service in the target region.' },
      'healthcareapis': { yes: 'Create a new FHIR service in the target region.', no: 'Region move not supported. Create a new service in the target region.' },
      'sqlvirtualmachine': { yes: 'Move the underlying VM and re-register with SQL VM resource provider.', no: 'Region move not supported. Move the VM and re-register in the target region.' },
      'datashare': { yes: 'Create a new Data Share account in the target region.', no: 'Region move not supported. Create a new account in the target region.' },
      'streamanalytics': { yes: 'Create a new Stream Analytics job in the target region.', no: 'Region move not supported. Create a new job in the target region.' },
      'datalakeanalytics': { yes: 'Create a new Data Lake Analytics account in the target region.', no: 'Region move not supported. Create a new account in the target region.' },
      'datalakestore': { yes: 'Create a new Data Lake Store account in the target region.', no: 'Region move not supported. Create a new account and migrate data.' },
      'analysisservices': { yes: 'Create a new Analysis Services server in the target region.', no: 'Region move not supported. Create a new server and restore models.' },
      'security': { yes: 'Security configurations are global; recreate in target if needed.', no: 'Security settings are typically subscription-scoped. Reconfigure for new region if needed.' },
      'alertsmanagement': { yes: 'Recreate alert rules in the target region.', no: 'Region move not supported. Recreate alert rules in the target region.' },
      'devtestlab': { yes: 'Create a new lab in the target region.', no: 'Region move not supported. Create a new lab in the target region.' },
      'avs': { yes: 'Create a new AVS private cloud in the target region.', no: 'Region move not supported. Create a new private cloud in the target region.' },
      'classiccompute': { yes: 'Migrate to ARM and use Azure Resource Mover.', no: 'Classic resources do not support region move. Migrate to ARM first, then use Azure Resource Mover.' },
      'classicnetwork': { yes: 'Migrate to ARM and recreate in target region.', no: 'Classic resources do not support region move. Migrate to ARM-based networking.' },
      'classicstorage': { yes: 'Migrate to ARM-based storage and copy data.', no: 'Classic resources do not support region move. Migrate to ARM storage first.' },
      'mobilenetwork': { yes: 'Use site-level move for mobile network resources.', no: 'Region move not supported for mobile network resources.' },
      'resources': { yes: 'Resource groups are metadata; create new group in target region.', no: 'Resource groups are region-scoped metadata containers.' },
    };

    const guidance = providerGuidance[provider];
    if (guidance) {
      return regionMove === 'Yes' ? guidance.yes : guidance.no;
    }

    // Generic fallback
    if (regionMove === 'Yes') {
      return `Region move supported. Use Azure Resource Mover or redeploy ${resource} in the target region. See https://learn.microsoft.com/en-us/azure/resource-mover/overview`;
    }
    return `Region move not supported for ${normalizedType}. Redeploy in the target region using ARM/Bicep templates or automation.`;
  }

  /**
   * Static remarks for region move guidance — resource-specific overrides.
   */
  get _regionRemarks() {
    if (this.__regionRemarks) return this.__regionRemarks;
    this.__regionRemarks = {
      // Compute
      'microsoft.compute/virtualmachines': 'Use Azure Resource Mover to move VMs across regions. VM will be recreated in the target region. See https://learn.microsoft.com/en-us/azure/resource-mover/tutorial-move-region-virtual-machines',
      'microsoft.compute/availabilitysets': 'Can be moved with Azure Resource Mover as part of VM region move.',
      'microsoft.compute/disks': 'Managed disks can be moved across regions using Azure Resource Mover. Snapshots are created and used to recreate disks in the target region.',
      'microsoft.compute/virtualmachinescalesets': 'VMSS does not support direct region move. Redeploy in the target region using ARM templates or Bicep.',
      'microsoft.compute/snapshots': 'Copy snapshot to the target region using Azure CLI or PowerShell, then recreate disk from it.',
      'microsoft.compute/images': 'Copy image to the target region using Azure CLI or shared image gallery replication.',
      'microsoft.compute/galleries': 'Shared Image Gallery does not support region move. Use image replication to the target region.',
      'microsoft.compute/proximityplacementgroups': 'Create a new proximity placement group in the target region and reassign VMs.',
      'microsoft.compute/diskencryptionsets': 'Create a new disk encryption set in the target region with a new Key Vault key.',
      'microsoft.compute/hostgroups': 'Dedicated host groups are region-bound. Create a new host group in the target region.',
      'microsoft.compute/cloudservices': 'Cloud Services (extended support) do not support region move. Redeploy in the target region.',
      'microsoft.compute/sshpublickeys': 'SSH public keys do not support region move. Create new SSH key resources in the target region.',
      'microsoft.compute/capacityreservationgroups': 'Create a new capacity reservation group in the target region.',

      // Network
      'microsoft.network/virtualnetworks': 'VNets can be moved across regions using Azure Resource Mover. Address spaces and subnets are recreated in the target region.',
      'microsoft.network/networksecuritygroups': 'NSGs can be moved across regions using Azure Resource Mover. Security rules are recreated in the target region.',
      'microsoft.network/networkinterfaces': 'NICs are moved as part of VM region move with Azure Resource Mover.',
      'microsoft.network/publicipaddresses': 'Public IPs can be moved across regions. Note: IP address value will change. Use DNS names for references.',
      'microsoft.network/loadbalancers': 'Load balancers can be moved across regions using Azure Resource Mover or export/import ARM template.',
      'microsoft.network/applicationgateways': 'Application Gateway does not support region move. Redeploy in the target region with new config.',
      'microsoft.network/azurefirewalls': 'Azure Firewall does not support region move. Deploy a new firewall in the target region.',
      'microsoft.network/bastionhosts': 'Bastion does not support region move. Deploy a new Bastion host in the target VNet.',
      'microsoft.network/vpngateways': 'VPN Gateways do not support region move. Create a new gateway in the target region and reconfigure tunnels.',
      'microsoft.network/virtualnetworkgateways': 'VNet Gateways do not support region move. Create a new gateway in the target region.',
      'microsoft.network/expressroutecircuits': 'ExpressRoute circuits are tied to peering locations. Create a new circuit for the target region if needed.',
      'microsoft.network/routetables': 'Route tables can be moved across regions. Routes are recreated in the target region.',
      'microsoft.network/dnszones': 'DNS zones are global. No region move needed.',
      'microsoft.network/privatednszones': 'Private DNS zones are global. No region move needed, but re-link VNets in the target region.',
      'microsoft.network/trafficmanagerprofiles': 'Traffic Manager is a global DNS service. No region move needed; update endpoints to target region.',
      'microsoft.network/frontdoors': 'Front Door is a global service. No region move needed; update backend pools to target region.',
      'microsoft.network/ddosprotectionplans': 'DDoS Protection Plans do not support region move. Create a new plan in the target region.',
      'microsoft.network/privateendpoints': 'Private endpoints are region-specific. Create new endpoints in the target region linked to moved resources.',
      'microsoft.network/privatelinkservices': 'Private Link services do not support region move. Recreate in the target region.',
      'microsoft.network/natgateways': 'NAT Gateways do not support region move. Create a new NAT gateway in the target region.',
      'microsoft.network/firewallpolicies': 'Firewall policies do not support region move. Create a new policy in the target region.',
      'microsoft.network/applicationsecuritygroups': 'ASGs can be moved across regions. Recreate and reassign in the target region.',
      'microsoft.network/networkwatchers': 'Network Watcher is auto-created per region. No manual move needed.',
      'microsoft.network/connections': 'Gateway connections must be recreated in the target region after moving gateways.',
      'microsoft.network/localnetworkgateways': 'Local network gateways must be recreated in the target region.',
      'microsoft.network/virtualwans': 'Virtual WAN does not support region move. Create new WAN and hubs in the target region.',
      'microsoft.network/virtualhubs': 'Virtual Hubs do not support region move. Create new hubs in the target region.',
      'microsoft.network/ipgroups': 'IP Groups do not support region move. Recreate in the target region.',
      'microsoft.network/serviceendpointpolicies': 'Service endpoint policies can be recreated in the target region.',

      // SQL
      'microsoft.sql/servers': 'SQL servers can be moved across regions using active geo-replication or failover groups. See https://learn.microsoft.com/en-us/azure/azure-sql/database/move-resources-across-regions',
      'microsoft.sql/servers/databases': 'Databases can be moved using geo-replication failover or backup/restore. See https://learn.microsoft.com/en-us/azure/azure-sql/database/move-resources-across-regions',
      'microsoft.sql/servers/elasticpools': 'Elastic pools can be moved across regions via geo-replication.',
      'microsoft.sql/managedinstances': 'SQL Managed Instance supports cross-region move. Requires re-creation in target region. See https://learn.microsoft.com/en-us/azure/azure-sql/managed-instance/move-across-regions',
      'microsoft.sql/managedinstances/databases': 'MI databases move with the managed instance during cross-region migration.',

      // MySQL / PostgreSQL / MariaDB
      'microsoft.dbformysql/servers': 'Use read replicas for cross-region migration, then promote replica. Or use backup/restore.',
      'microsoft.dbformysql/flexibleservers': 'Use read replicas or backup/restore for cross-region migration of MySQL Flexible Server.',
      'microsoft.dbforpostgresql/servers': 'Use read replicas for cross-region migration, then promote replica.',
      'microsoft.dbforpostgresql/flexibleservers': 'Use read replicas or backup/restore for cross-region migration of PostgreSQL Flexible Server.',
      'microsoft.dbformariadb/servers': 'Use backup/restore for cross-region migration of MariaDB servers.',

      // Storage
      'microsoft.storage/storageaccounts': 'Move storage by creating a new account in the target region and copying data with AzCopy. See https://learn.microsoft.com/en-us/azure/storage/common/storage-account-move',

      // Cosmos DB
      'microsoft.documentdb/databaseaccounts': 'Cosmos DB does not support direct region move. Add the target region as a replica, failover, then remove the old region.',

      // Cache
      'microsoft.cache/redis': 'Azure Cache for Redis does not support region move. Create a new cache in the target region and migrate data using import/export.',
      'microsoft.cache/redisenterprise': 'Redis Enterprise does not support region move. Create a new cluster in the target region.',

      // Key Vault
      'microsoft.keyvault/vaults': 'Key Vault does not support region move. Create a new vault in the target region and migrate secrets/keys/certificates. See https://learn.microsoft.com/en-us/azure/key-vault/general/move-region',
      'microsoft.keyvault/managedhsms': 'Managed HSM does not support region move. Create a new HSM in the target region.',

      // Web / App Service
      'microsoft.web/sites': 'App Service does not support direct region move. Redeploy using ARM templates, CI/CD, or clone app to new plan. See https://learn.microsoft.com/en-us/azure/app-service/manage-move-across-regions',
      'microsoft.web/serverfarms': 'App Service Plans do not support region move. Create a new plan in the target region and redeploy apps.',
      'microsoft.web/staticsites': 'Static Web Apps are global. No region move needed.',
      'microsoft.web/certificates': 'App Service certificates must be recreated in the target region.',
      'microsoft.web/hostingenvironments': 'App Service Environment does not support region move. Create new ASE in the target region.',

      // Containers
      'microsoft.containerservice/managedclusters': 'AKS clusters do not support region move. Redeploy cluster, node pools, and workloads in the target region.',
      'microsoft.containerregistry/registries': 'ACR does not support region move. Enable geo-replication to the target region or create a new registry.',
      'microsoft.containerinstance/containergroups': 'Container Instances do not support region move. Redeploy container group in the target region.',
      'microsoft.app/containerapps': 'Container Apps do not support region move. Redeploy in a new managed environment in the target region.',
      'microsoft.app/managedenvironments': 'Container Apps environments do not support region move. Create a new environment in the target region.',

      // Monitoring
      'microsoft.insights/components': 'Application Insights does not support region move. Create a new resource and reconfigure instrumentation keys/connection strings.',
      'microsoft.insights/actiongroups': 'Action Groups do not support region move. Recreate in the target region.',
      'microsoft.insights/metricalerts': 'Metric alerts do not support region move. Recreate alert rules in the target region.',
      'microsoft.insights/activitylogalerts': 'Activity log alerts are subscription-scoped. No region move needed.',
      'microsoft.insights/scheduledqueryrules': 'Log alert rules can be recreated in the target region.',
      'microsoft.insights/autoscalesettings': 'Autoscale settings must be recreated for resources in the target region.',
      'microsoft.insights/workbooks': 'Workbooks can be recreated in the target region.',
      'microsoft.insights/webtests': 'Web tests must be recreated in the target region.',
      'microsoft.insights/datacollectionrules': 'Data collection rules do not support region move. Recreate in the target region.',
      'microsoft.operationalinsights/workspaces': 'Log Analytics workspaces do not support region move. Create a new workspace and reconfigure agents/data sources.',

      // AI / Cognitive
      'microsoft.cognitiveservices/accounts': 'Cognitive Services does not support region move. Create a new resource in the target region and update API endpoints.',
      'microsoft.machinelearningservices/workspaces': 'Azure ML workspaces do not support region move. Create new workspace and retrain/redeploy models.',
      'microsoft.search/searchservices': 'Azure AI Search does not support region move. Create a new service and rebuild indexes.',

      // Messaging
      'microsoft.servicebus/namespaces': 'Service Bus can be moved using geo-disaster recovery pairing. See https://learn.microsoft.com/en-us/azure/service-bus-messaging/service-bus-geo-dr',
      'microsoft.eventhub/namespaces': 'Event Hubs can be moved using geo-disaster recovery pairing. See https://learn.microsoft.com/en-us/azure/event-hubs/event-hubs-geo-dr',
      'microsoft.eventhub/clusters': 'Event Hub clusters do not support region move. Create a new cluster in the target region.',
      'microsoft.relay/namespaces': 'Relay namespaces do not support region move. Create a new namespace in the target region.',
      'microsoft.notificationhubs/namespaces': 'Notification Hub does not support region move. Create new namespace and hub in the target region.',

      // Data & Analytics
      'microsoft.datafactory/factories': 'Data Factory does not support region move. Create a new factory and export/import pipelines using ARM templates.',
      'microsoft.databricks/workspaces': 'Databricks workspaces do not support region move. Create a new workspace and migrate notebooks/jobs.',
      'microsoft.synapse/workspaces': 'Synapse workspaces do not support region move. Create a new workspace and migrate artifacts.',
      'microsoft.hdinsight/clusters': 'HDInsight clusters do not support region move. Create a new cluster in the target region.',
      'microsoft.kusto/clusters': 'Data Explorer clusters do not support region move. Create a new cluster and migrate databases.',
      'microsoft.streamanalytics/streamingjobs': 'Stream Analytics jobs must be recreated in the target region.',
      'microsoft.datalakeanalytics/accounts': 'Data Lake Analytics does not support region move. Create a new account.',
      'microsoft.datalakestore/accounts': 'Data Lake Store does not support region move. Create a new account and migrate data.',
      'microsoft.datashare/accounts': 'Data Share accounts do not support region move. Create a new account.',
      'microsoft.analysisservices/servers': 'Analysis Services does not support region move. Create a new server and restore models.',
      'microsoft.purview/accounts': 'Purview does not support region move. Create a new account in the target region.',

      // IoT
      'microsoft.devices/iothubs': 'IoT Hub supports cross-region move via manual failover. See https://learn.microsoft.com/en-us/azure/iot-hub/iot-hub-ha-dr',
      'microsoft.devices/provisioningservices': 'DPS does not support region move. Create a new DPS instance in the target region.',
      'microsoft.iotcentral/iotapps': 'IoT Central does not support region move. Create a new application in the target region.',
      'microsoft.timeseriesinsights/environments': 'Time Series Insights does not support region move. Create a new environment.',

      // Identity
      'microsoft.managedidentity/userassignedidentities': 'Managed identities do not support region move. Create new identity and update RBAC role assignments.',

      // Integration
      'microsoft.logic/workflows': 'Logic Apps do not support region move. Export definition and redeploy in the target region.',
      'microsoft.logic/integrationaccounts': 'Integration accounts do not support region move. Create a new account in the target region.',
      'microsoft.automation/automationaccounts': 'Automation accounts do not support region move. Create a new account and migrate runbooks.',

      // API Management
      'microsoft.apimanagement/service': 'API Management supports region move via backup/restore. See https://learn.microsoft.com/en-us/azure/api-management/api-management-howto-disaster-recovery-backup-restore',

      // CDN
      'microsoft.cdn/profiles': 'CDN/Front Door profiles are global. No region move needed; update origins to the target region.',

      // Recovery / Backup
      'microsoft.recoveryservices/vaults': 'Recovery Services vaults do not support region move. Create a new vault and re-protect workloads.',
      'microsoft.dataprotection/backupvaults': 'Backup vaults do not support region move. Create a new vault and reconfigure backup policies.',

      // Other
      'microsoft.signalrservice/signalr': 'SignalR does not support region move. Create a new instance in the target region.',
      'microsoft.portal/dashboards': 'Dashboards are global resources. No region move needed.',
      'microsoft.appconfiguration/configurationstores': 'App Configuration does not support region move. Create a new store and export/import settings.',
      'microsoft.communication/communicationservices': 'Communication Services does not support region move. Create a new resource.',
      'microsoft.digitaltwins/digitaltwinsinstances': 'Digital Twins supports region move. See https://learn.microsoft.com/en-us/azure/digital-twins/how-to-move-regions',
      'microsoft.botservice/botservices': 'Bot Service does not support region move. Redeploy the bot in the target region.',
      'microsoft.eventgrid/topics': 'Event Grid topics do not support region move. Create new topics in the target region.',
      'microsoft.eventgrid/domains': 'Event Grid domains do not support region move. Create new domains in the target region.',
      'microsoft.eventgrid/systemtopics': 'System topics are auto-created. They will appear when resources are moved to the target region.',
      'microsoft.batch/batchaccounts': 'Batch accounts do not support region move. Create a new account in the target region.',
      'microsoft.maps/accounts': 'Azure Maps is a global service. No region move needed.',
      'microsoft.media/mediaservices': 'Media Services does not support region move. Create a new account in the target region.',
      'microsoft.migrate/projects': 'Azure Migrate projects do not support region move. Create a new project.',
      'microsoft.desktopvirtualization/hostpools': 'AVD host pools do not support region move. Create a new host pool and session hosts in the target region.',
      'microsoft.desktopvirtualization/applicationgroups': 'Application groups follow the host pool. Recreate in the target region.',
      'microsoft.desktopvirtualization/workspaces': 'AVD workspaces do not support region move. Create a new workspace.',
      'microsoft.netapp/netappaccounts': 'NetApp does not support region move. Use cross-region replication to migrate volumes.',
      'microsoft.appplatform/spring': 'Azure Spring Apps does not support region move. Create a new instance in the target region.',
      'microsoft.loadtestservice/loadtests': 'Azure Load Testing does not support region move. Create a new resource.',
      'microsoft.avs/privateclouds': 'AVS private clouds do not support region move. Create a new private cloud in the target region.',
      'microsoft.devtestlab/labs': 'DevTest Labs do not support region move. Create a new lab in the target region.',
      'microsoft.hybridcompute/machines': 'Arc machines do not support region move. Re-register in the target region.',
      'microsoft.redhatopenshift/openshiftclusters': 'ARO clusters do not support region move. Create a new cluster.',
      'microsoft.servicefabric/clusters': 'Service Fabric clusters do not support region move. Create a new cluster and redeploy applications.',
      'microsoft.storagesync/storagesyncservices': 'Storage Sync Service does not support region move. Create a new service.',
      'microsoft.powerbidedicated/capacities': 'Power BI Embedded does not support region move. Create a new capacity.',
      'microsoft.healthcareapis/services': 'Healthcare APIs do not support region move. Create a new service in the target region.',
      'microsoft.sqlvirtualmachine/sqlvirtualmachines': 'Move the underlying VM first, then re-register with the SQL VM resource provider in the target region.',
      'microsoft.domainregistration/domains': 'Domain registrations are global. No region move needed.',
      'microsoft.certificateregistration/certificateorders': 'Certificate orders are global. No region move needed.',
    };
    return this.__regionRemarks;
  }

  /**
   * Assess region move support for an array of resources.
   */
  assessRegionResources(resources) {
    if (regionAssessor) {
      return regionAssessor.assessResources(resources, this.normalizeType.bind(this));
    }
    // Fallback
    if (!resources || resources.length === 0) return [];
    const firstRow = resources[0];
    const typeCol = this._detectTypeColumn(firstRow);
    if (!typeCol) {
      throw new Error(`Could not find a "Resource Type" column in your file.`);
    }
    return resources.map(resource => {
      const clean = this._stripAssessmentColumns(resource);
      const typeValue = typeCol ? (resource[typeCol] || '') : '';
      const assessment = this.assessRegionType(typeValue);
      return {
        ...clean,
        'REGION MOVE SUPPORTED': assessment.regionMove,
        'NORMALIZED TYPE': assessment.normalizedType,
        'REMARKS': assessment.remarks
      };
    });
  }

  /**
   * Get summary for region move assessment.
   */
  getRegionSummary(assessedResources) {
    const total = assessedResources.length;
    const yes = assessedResources.filter(r => r['REGION MOVE SUPPORTED'] === 'Yes').length;
    const no = assessedResources.filter(r => r['REGION MOVE SUPPORTED'] === 'No').length;
    const review = assessedResources.filter(r => r['REGION MOVE SUPPORTED'] === 'Review').length;

    return { total, yes, no, review };
  }

  /**
   * Detect the column that contains the resource type from the Excel data.
   * Azure portal exports use various column names.
   */
  _detectTypeColumn(row) {
    // Priority-ordered list of possible column names (case-insensitive matching)
    const candidates = [
      // "Resource Type" variants first — Azure exports often have both "Type" (display name)
      // and "Resource Type" (ARM format). ARM format is what we need.
      'RESOURCE TYPE', 'Resource Type', 'Resource type', 'resource type',
      'Resource_Type', 'resource_type', 'RESOURCE_TYPE',
      'ResourceType', 'resourceType', 'resourcetype',
      'Azure Resource Type', 'azure resource type',
      'Resource Provider/Type',
      // "Type" last — may contain display names like "Virtual machine" instead of ARM types
      'TYPE', 'Type', 'type',
    ];

    for (const col of candidates) {
      if (row[col] !== undefined && row[col] !== null && row[col] !== '') {
        return col;
      }
    }

    // Fallback: find any column whose name contains "type" (case-insensitive)
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
   * Detect the column that contains the resource name.
   */
  _detectNameColumn(row) {
    const candidates = [
      'NAME', 'Name', 'name',
      'RESOURCE NAME', 'Resource Name', 'Resource name', 'resource name',
      'Resource_Name', 'resource_name', 'RESOURCE_NAME',
      'ResourceName', 'resourceName',
    ];

    for (const col of candidates) {
      if (row[col] !== undefined && row[col] !== null) {
        return col;
      }
    }

    const keys = Object.keys(row);
    for (const key of keys) {
      if (/\bname\b/i.test(key) && !/\b(group|subscription|location)\b/i.test(key)) {
        return key;
      }
    }

    return null;
  }

  /**
   * Strip any previously-added assessment columns so re-uploading an assessed file
   * produces a clean result with only the current assessment's columns.
   */
  _stripAssessmentColumns(resource) {
    const stripKeys = new Set([
      'SUBSCRIPTION MOVE SUPPORTED',
      'REGION MOVE SUPPORTED',
      'JIO REGION AVAILABLE',
      'JIO SERVICE NAME',
      'CURRENT REGION',
      'INDIA REGION',
      'NORMALIZED TYPE',
      'REMARKS'
    ]);
    const clean = {};
    for (const [key, value] of Object.entries(resource)) {
      if (!stripKeys.has(key.toUpperCase())) {
        clean[key] = value;
      }
    }
    return clean;
  }

  assessResources(resources) {
    if (subscriptionAssessor) {
      return subscriptionAssessor.assessResources(resources, this.normalizeType.bind(this));
    }
    // Fallback: inline logic if module failed to load
    if (!resources || resources.length === 0) return [];

    const firstRow = resources[0];
    const typeCol = this._detectTypeColumn(firstRow);

    if (!typeCol) {
      const availableCols = Object.keys(firstRow).join(', ');
      throw new Error(`Could not find a "Resource Type" column in your file. Available columns: [${availableCols}].`);
    }

    return resources.map(resource => {
      const clean = this._stripAssessmentColumns(resource);
      const typeValue = typeCol ? (resource[typeCol] || '') : '';
      const assessment = this.assessType(typeValue);
      return {
        ...clean,
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

  getRules() {
    return {
      metadata: this.metadata,
      totalRules: Object.keys(this.rules).length,
      totalRegionRules: Object.keys(this.regionRules).length,
      totalJioServices: Object.keys(this.jioServices).length,
      totalJioVMs: Object.keys(this.jioVMs).length,
      source: this.rulesSource,
      csvRuleCount: this.csvRuleCount,
      regionCsvRuleCount: this.regionCsvRuleCount,
      staticOverrides: Object.keys(this.staticRules).length,
      lastRefreshed: this.lastRefreshed,
      rules: this.rules
    };
  }

  /**
   * ARM type → Jio service display name mapping for lookup.
   */
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
   */
  assessJioType(rawType) {
    const normalized = this.normalizeType(rawType);

    // Try ARM → Jio name mapping first
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

    // Try fuzzy match against Jio service names using display name
    const displayName = rawType.trim().toLowerCase();
    const fuzzyMatch = this._fuzzyMatchJio(displayName) || this._fuzzyMatchJio(normalized);
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

  /**
   * Fuzzy match a name against Jio services list.
   */
  _fuzzyMatchJio(input) {
    if (!input) return null;
    const lower = input.toLowerCase().replace(/^microsoft\.\w+\//, '').replace(/\//g, ' ');

    // Direct key match
    if (this.jioServices[lower]) return this.jioServices[lower];

    // Try partial matches
    const jioKeys = Object.keys(this.jioServices);
    for (const key of jioKeys) {
      if (key.includes(lower) || lower.includes(key)) {
        return this.jioServices[key];
      }
    }

    // Try matching core words
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
  assessJioResources(resources) {
    if (jioAssessor) {
      return jioAssessor.assessResources(resources, this.normalizeType.bind(this));
    }
    // Fallback
    if (!resources || resources.length === 0) return [];
    const firstRow = resources[0];
    const typeCol = this._detectTypeColumn(firstRow);
    if (!typeCol) {
      throw new Error(`Could not find a "Resource Type" column in your file.`);
    }
    return resources.map(resource => {
      const clean = this._stripAssessmentColumns(resource);
      const typeValue = typeCol ? (resource[typeCol] || '') : '';
      const assessment = this.assessJioType(typeValue);
      return {
        ...clean,
        'JIO REGION AVAILABLE': assessment.jioAvailable,
        'NORMALIZED TYPE': assessment.normalizedType,
        'REMARKS': assessment.remarks
      };
    });
  }

  /**
   * Detect the location/region column in a resource row.
   */
  _detectLocationColumn(row) {
    const candidates = [
      'LOCATION', 'Location', 'location',
      'REGION', 'Region', 'region',
      'Resource Location', 'RESOURCE LOCATION', 'resource location',
      'Resource_Location', 'resource_location'
    ];
    for (const col of candidates) {
      if (row[col] !== undefined && row[col] !== null) return col;
    }
    const keys = Object.keys(row);
    for (const key of keys) {
      if (/\b(location|region)\b/i.test(key) && !/\b(group|resource group)\b/i.test(key)) return key;
    }
    return null;
  }

  /**
   * Get summary for Jio availability assessment.
   */
  getJioSummary(assessedResources) {
    const total = assessedResources.length;
    const yes = assessedResources.filter(r => r['JIO REGION AVAILABLE'] === 'Yes').length;
    const no = assessedResources.filter(r => r['JIO REGION AVAILABLE'] === 'No').length;
    const review = assessedResources.filter(r => r['JIO REGION AVAILABLE'] === 'Review').length;

    return { total, yes, no, review };
  }

  // ══════════════════════════════════════════════════════════════════════════
  // AWS TO AZURE MIGRATION COMPARISON
  // ══════════════════════════════════════════════════════════════════════════
  // Delegated to isolated assessors/awsAssessor.js module.
  // If the AWS module has errors, other assessment modes still work.

  assessAwsType(rawType, skuOrSize) {
    if (!awsAssessor) throw new Error('AWS Assessor module is not available. Check server logs for load errors.');
    return awsAssessor.assessType(rawType, skuOrSize);
  }

  assessAwsResources(resources) {
    if (!awsAssessor) throw new Error('AWS Assessor module is not available. Check server logs for load errors.');
    return awsAssessor.assessResources(resources);
  }

  getAwsSummary(assessedResources) {
    if (!awsAssessor) throw new Error('AWS Assessor module is not available. Check server logs for load errors.');
    return awsAssessor.getSummary(assessedResources);
  }

  reloadAwsMapping() {
    if (!awsAssessor) throw new Error('AWS Assessor module is not available.');
    awsAssessor.reloadMapping();
  }

  // ══════════════════════════════════════════════════════════════════════════
  // GCP TO AZURE MIGRATION COMPARISON
  // ══════════════════════════════════════════════════════════════════════════
  // Delegated to isolated assessors/gcpAssessor.js module.
  // If the GCP module has errors, other assessment modes still work.

  assessGcpType(rawType, skuOrSize) {
    if (!gcpAssessor) throw new Error('GCP Assessor module is not available. Check server logs for load errors.');
    return gcpAssessor.assessType(rawType, skuOrSize);
  }

  assessGcpResources(resources) {
    if (!gcpAssessor) throw new Error('GCP Assessor module is not available. Check server logs for load errors.');
    return gcpAssessor.assessResources(resources);
  }

  getGcpSummary(assessedResources) {
    if (!gcpAssessor) throw new Error('GCP Assessor module is not available. Check server logs for load errors.');
    return gcpAssessor.getSummary(assessedResources);
  }

  reloadGcpMapping() {
    if (!gcpAssessor) throw new Error('GCP Assessor module is not available.');
    gcpAssessor.reloadMapping();
  }
}

module.exports = new MigrationService();
