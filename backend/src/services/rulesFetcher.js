const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');

const LEARN_PAGE_URL = 'https://learn.microsoft.com/en-us/azure/azure-resource-manager/management/move-support-resources';
const CACHE_PATH = path.join(__dirname, '..', 'data', 'learn-rules-cache.json');

/**
 * Generic HTTPS fetcher with redirect support
 */
function fetchURL(url) {
  return new Promise((resolve, reject) => {
    const get = url.startsWith('https') ? https.get : http.get;

    get(url, { timeout: 30000, headers: { 'User-Agent': 'Mozilla/5.0 AzureCSPMigrationTool/1.0' } }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        const redirectUrl = res.headers.location;
        const redirectGet = redirectUrl.startsWith('https') ? https.get : http.get;
        redirectGet(redirectUrl, { timeout: 30000, headers: { 'User-Agent': 'Mozilla/5.0 AzureCSPMigrationTool/1.0' } }, (redirectRes) => {
          collectData(redirectRes, resolve, reject);
        }).on('error', reject);
        return;
      }

      if (res.statusCode !== 200) {
        reject(new Error(`HTTP ${res.statusCode} fetching ${url}`));
        return;
      }

      collectData(res, resolve, reject);
    }).on('error', reject);
  });
}

function collectData(res, resolve, reject) {
  let data = '';
  res.on('data', chunk => { data += chunk; });
  res.on('end', () => resolve(data));
  res.on('error', reject);
}

/**
 * Strip HTML tags from a string and decode entities.
 */
function stripHtml(html) {
  return html
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .trim();
}

/**
 * Determine Yes/No/Conditional from a table cell's HTML content.
 * Cells with <strong>Yes</strong> → "Yes"
 * Cells with just "No" → "No"
 * Cells with "Yes" plus extra text (conditions) → "Yes" (remarks captured separately)
 */
function parseCell(cellHtml) {
  const text = stripHtml(cellHtml);
  if (!text) return { value: 'No', remarks: '' };

  // Check for Yes (may be wrapped in <strong>)
  if (/\byes\b/i.test(text)) {
    // Extract conditional remarks (everything after "Yes" that isn't just whitespace)
    const remarksMatch = text.replace(/^yes\s*/i, '').trim();
    return { value: 'Yes', remarks: remarksMatch || '' };
  }

  // "pending" or other values → treat as No with remark
  if (/\bpending\b/i.test(text)) {
    return { value: 'No', remarks: 'Pending' };
  }

  return { value: 'No', remarks: '' };
}

/**
 * Parse the Microsoft Learn move-support-resources page HTML.
 * Extracts all provider/resource tables and returns both subscription and region rules.
 *
 * Returns: { subscriptionRules, regionRules }
 *   subscriptionRules: { "microsoft.xxx/yyy": { subscriptionMove: "Yes"|"No", remarks: "" } }
 *   regionRules:       { "microsoft.xxx/yyy": { regionMove: "Yes"|"No", remarks: "" } }
 */
function parseLearnPage(html) {
  const subscriptionRules = {};
  const regionRules = {};

  // Match each <h2> provider heading followed by its table(s)
  // Pattern: <h2 id="...">ProviderName</h2> ... <table>...</table>
  const providerPattern = /<h2[^>]*>([^<]+)<\/h2>([\s\S]*?)(?=<h2[^>]*>|<\/main>|$)/gi;
  let providerMatch;

  while ((providerMatch = providerPattern.exec(html)) !== null) {
    const providerName = stripHtml(providerMatch[1]).trim();
    const sectionHtml = providerMatch[2];

    // Skip non-provider sections (e.g., "Next steps", "Third-party services")
    if (!providerName.startsWith('Microsoft.') && !providerName.startsWith('microsoft.')) continue;

    // Extract all tables in this section
    const tablePattern = /<tbody>([\s\S]*?)<\/tbody>/gi;
    let tableMatch;

    while ((tableMatch = tablePattern.exec(sectionHtml)) !== null) {
      const tbodyHtml = tableMatch[1];

      // Extract rows
      const rowPattern = /<tr>([\s\S]*?)<\/tr>/gi;
      let rowMatch;

      while ((rowMatch = rowPattern.exec(tbodyHtml)) !== null) {
        const rowHtml = rowMatch[1];

        // Extract cells
        const cellPattern = /<td>([\s\S]*?)<\/td>/gi;
        const cells = [];
        let cellMatch;
        while ((cellMatch = cellPattern.exec(rowHtml)) !== null) {
          cells.push(cellMatch[1]);
        }

        if (cells.length < 3) continue;

        const resourceType = stripHtml(cells[0]).replace(/\(.*\)/, '').trim();
        if (!resourceType) continue;

        const fullType = `${providerName}/${resourceType}`;
        const key = fullType.toLowerCase();

        // Column 1: Resource group move (we don't use this)
        // Column 2: Subscription move
        const subCell = parseCell(cells[2]);

        // Don't overwrite "Yes" with "No" for duplicate entries
        if (!(subscriptionRules[key] && subscriptionRules[key].subscriptionMove === 'Yes' && subCell.value === 'No')) {
          subscriptionRules[key] = {
            subscriptionMove: subCell.value,
            remarks: subCell.remarks
          };
        }

        // Column 3: Region move (if present)
        if (cells.length >= 4) {
          const regionCell = parseCell(cells[3]);
          if (!(regionRules[key] && regionRules[key].regionMove === 'Yes' && regionCell.value === 'No')) {
            regionRules[key] = {
              regionMove: regionCell.value,
              remarks: regionCell.remarks
            };
          }
        }
      }
    }
  }

  return { subscriptionRules, regionRules };
}

/**
 * Fetch rules from the official Microsoft Learn documentation page (real-time).
 * Returns: { subscriptionRules, regionRules }
 */
async function fetchLearnPageRules() {
  const html = await fetchURL(LEARN_PAGE_URL);
  return parseLearnPage(html);
}

/**
 * Fetch subscription move rules — fetches from Learn page.
 * If unreachable, uses previously cached data from disk.
 * Returns: { subscriptionRules, regionRules, source }
 */
async function fetchAllRules() {
  try {
    const { subscriptionRules, regionRules } = await fetchLearnPageRules();
    const subCount = Object.keys(subscriptionRules).length;
    const regionCount = Object.keys(regionRules).length;

    if (subCount < 100) {
      throw new Error(`Learn page returned only ${subCount} subscription rules (expected 500+), data may be incomplete`);
    }

    console.log(`Fetched ${subCount} subscription + ${regionCount} region rules from Microsoft Learn (real-time)`);

    // Cache the fetched data to disk for offline use
    const cacheData = {
      subscriptionRules,
      regionRules,
      cachedAt: new Date().toISOString()
    };
    fs.writeFileSync(CACHE_PATH, JSON.stringify(cacheData, null, 2), 'utf-8');

    return { subscriptionRules, regionRules, source: 'microsoft-learn' };
  } catch (err) {
    console.warn('Microsoft Learn page fetch failed:', err.message);

    // Use cached data from disk if available
    if (fs.existsSync(CACHE_PATH)) {
      const cached = JSON.parse(fs.readFileSync(CACHE_PATH, 'utf-8'));
      console.log(`Using cached rules from ${cached.cachedAt} (${Object.keys(cached.subscriptionRules).length} subscription + ${Object.keys(cached.regionRules).length} region)`);
      return {
        subscriptionRules: cached.subscriptionRules,
        regionRules: cached.regionRules,
        source: `cached (${cached.cachedAt})`
      };
    }

    // No cache exists — throw so caller knows there's no data
    throw new Error('Microsoft Learn page unreachable and no cached data available. Please check internet connectivity.');
  }
}

module.exports = { fetchAllRules };
