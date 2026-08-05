911111// ─────────────────────────────────────────────
//  CONFIGURATION
// ─────────────────────────────────────────────
const CONFIG = {
  apolloApiKey:   null  // Replace with your Apollo API key
  claudeApiKey:   null  // Replace with your Claude API key
  leadsPerDay:    10,

 // web scraping criteria
    scraping: {
    maxAUM:         300,   // filter out companies with AUM above this (in millions)
    maxProperties:  50,    // filter out companies with more than this many properties
  },

  sheets: {
    criteria: "Criteria",
    leads:    "Leads",
    feedback: "Feedback",
  },

  // Leads sheet column order
  leadColumns: [
    "Date Added", "Company", "Website", "Location", "AUM", "AUM Source", "Properties",
    "Head of Acquisitions", "Email", "Phone",
    "AI Score", "AI Reasoning", "Status", "Feedback Notes"
  ],
};

// ─────────────────────────────────────────────
//  SCRAPING KEYWORD ARRAYS
// ─────────────────────────────────────────────
const AUM_KEYWORDS = [
  "AUM",
  "ASSETS UNDER MANAGEMENT",
  "INVESTED CAPITAL",
  "TOTAL ASSETS",
  "ASSETS MANAGED",
  "CAPITAL MANAGED",
  "PORTFOLIO VALUE",
  "TOTAL INVESTMENT VALUE",
];

const PROPERTY_KEYWORDS = [
  "PROPERTIES",
  "PORTFOLIO",
  "INVESTMENTS",
  "ASSETS",
  "DEVELOPMENTS",
  "PROJECTS",
  "BUILDINGS",
];

// ─────────────────────────────────────────────
//  MENU
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Prospecting Agent")
    .addItem("▶ Run now", "runAgent")
    .addItem("⚙ Setup sheets", "setupSheets")
    .addItem("📋 Open review panel", "showSidebar")
    .addItem("⏰ Set daily trigger", "createDailyTrigger")
    .addItem("☁ Export training data", "exportTrainingData")
    .addToUi();
    
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Lead Review")
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─────────────────────────────────────────────
//  SETUP — call once to initialize sheets
// ─────────────────────────────────────────────

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Criteria sheet ──
  let criteria = ss.getSheetByName(CONFIG.sheets.criteria);
  if (!criteria) criteria = ss.insertSheet(CONFIG.sheets.criteria);
  criteria.clearContents();
  criteria.getRange("A1:B1").setValues([["Field", "Value"]]);
  criteria.getRange("A2:B9").setValues([
    ["Target Role",        "Head of Acquisitions, Director of Acquisitions, VP Acquisitions, VP of Acquisitions, Acquisitions Manager, Acquisitions Director, Acquisitions Associate, Investment Manager, Director of Investments, VP of Investments, Chief Investment Officer, Vice President of Acquisitions"],
    ["Industry",           "Commercial Real Estate, Real Estate Development"],
    ["Property Types",     "Multifamily, Office, Industrial, Mixed-Use, Retail"],
    ["Min AUM (USD)",      ""],
    ["Max AUM (USD)",      ""],
    ["Locations",          "New York, New Jersey, North Dakota, Illinois, Ohio, Michigan, Indiana, Wisconsin, Minnesota, Iowa, Missouri, Kansas, Nebraska, South Dakota"],
    ["Keywords",           "real estate, commercial real estate, property development"],
    ["Exclude Keywords",   "residential only, single family, property management, bank"],
  ]);
  styleHeader(criteria, 2);

  // ── Leads sheet ──
  let leads = ss.getSheetByName(CONFIG.sheets.leads);
  if (!leads) leads = ss.insertSheet(CONFIG.sheets.leads);

  // ✅ Only reset header row — preserve existing lead data
  leads.getRange(1, 1, 1, CONFIG.leadColumns.length).setValues([CONFIG.leadColumns]);
  styleHeader(leads, CONFIG.leadColumns.length);

  // ── Feedback sheet ──
  let fb = ss.getSheetByName(CONFIG.sheets.feedback);
  if (!fb) fb = ss.insertSheet(CONFIG.sheets.feedback);
  fb.clearContents();
  fb.getRange("A1:F1").setValues([["Date", "Company", "Status", "AI Score", "Feedback Notes", "Criteria Snapshot"]]);
  styleHeader(fb, 6);

  // ── Settings sheet ──
  let settings = ss.getSheetByName("Settings");
  if (!settings) {
    settings = ss.insertSheet("Settings");
    settings.getRange("A1:B1").setValues([["Key", "Value"]]);
    settings.getRange("A2:B2").setValues([["rotationIndex", "0"]]);
    styleHeader(settings, 2);
  }

  SpreadsheetApp.getUi().alert("✅ Sheets are set up! Update your criteria in the Criteria sheet, then run the agent.");
}

// ─────────────────────────────────────────────
//  MAIN AGENT — run manually or on a trigger
// ─────────────────────────────────────────────
function runAgent() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const criteria = readCriteria(ss);
  const feedback = readFeedbackHistory(ss);
  const existing = readExistingCompanies(ss);

  Logger.log("Starting Apollo search...");
  const orgs = searchApollo(criteria, existing);
  Logger.log("Apollo returned: " + orgs.length + " orgs");

  // ✅ Filter by company name before scraping
  const nameFiltered = filterByName(orgs, criteria);
  Logger.log("After name filter: " + nameFiltered.length + " orgs remain");

  if (!orgs.length) {
    Logger.log("No new leads found, exiting.");
    SpreadsheetApp.getUi().alert("No new leads found from Apollo today.");
    return;
  }

  // Scrape and filter before enrichment ──
  Logger.log("Starting web scraping...");
  const filtered = scrapeAndFilter(orgs);
  Logger.log("After scraping: " + filtered.length + " orgs remain");

  if (!filtered.length) {
    SpreadsheetApp.getUi().alert("All leads filtered out by web scraping. Try adjusting scraping thresholds in CONFIG.");
    return;
  }

  Logger.log("Starting enrichment...");
  const enriched = orgs.map(org => enrichOrg(org, criteria));
  Logger.log("Enrichment complete.");

  Logger.log("Starting scoring...");
  const scored = enriched.map(lead => scoreLead(lead, criteria, feedback));
  Logger.log("Scoring complete.");

  writeLeads(ss, scored);
  SpreadsheetApp.getUi().alert(`✅ ${scored.length} new leads added to the Leads sheet.`);
}
// ─────────────────────────────────────────────
//  READ CRITERIA
// ─────────────────────────────────────────────
function readCriteria(ss) {
  const sheet = ss.getSheetByName(CONFIG.sheets.criteria);
  const data  = sheet.getDataRange().getValues();
  const out   = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) out[data[i][0]] = data[i][1];
  }
  return out;
}

// ─────────────────────────────────────────────
//  READ FEEDBACK HISTORY (last 100 entries)
// ─────────────────────────────────────────────
function readFeedbackHistory(ss) {
  const sheet   = ss.getSheetByName(CONFIG.sheets.feedback);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, Math.min(lastRow - 1, 100), 6).getValues();
  return rows
    .filter(r => r[2])
    .map(r => ({ company: r[1], status: r[2], score: r[3], notes: r[4] }));
}

// ─────────────────────────────────────────────
//  READ EXISTING COMPANIES (for deduplication)
// ─────────────────────────────────────────────
function readExistingCompanies(ss) {
  const sheet   = ss.getSheetByName(CONFIG.sheets.leads);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();

  const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  return new Set(names.map(n => normalizeCompanyName(n)));
}

// ─────────────────────────────────────────────
//  APOLLO — search organizations
// ─────────────────────────────────────────────

function searchApollo(criteria, existing) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getOrCreateSettings(ss);

  // ── Rotation sets ──
  const keywordSets = [
    ["multifamily investment", "apartment building", "commercial real estate"],
    ["industrial property development", "commercial real estate investment", "real estate"],
    ["commercial real estate refinance", "commercial loan brokering", "financial services"],
    ["hospitality investor", "student housing", "self storage"],
    ["capital markets", "investment banking", "investment management"],
    ["venture capital real estate", "construction", "architecture real estate"],
    ];

  const industries = [
    "commercial real estate",
    "capital markets",
    "investment banking",
    "venture capital & private equity",
    "architecture & planning",
    "construction",
    "financial services",
    "real estate",
    "investment management",
  ];

  const locations       = (criteria["Locations"] || "United States").split(",").map(l => l.trim());
  const excludeKeywords = (criteria["Exclude Keywords"] || "").split(",").map(k => k.trim()).filter(Boolean);

  // ── Pick keyword set for this run based on rotation index ──
  const rotationIndex = settings.rotationIndex || 0;
  const keywords      = keywordSets[rotationIndex % keywordSets.length];

  Logger.log(`Run #${rotationIndex + 1} — Keywords: ${keywords.join(", ")}`);

  const allOrgs    = [];
  const totalPages = 8;
  const startPage  = Math.floor(Math.random() * 100) + 1;
  Logger.log("Starting search from page: " + startPage + " to " + (startPage + totalPages - 1));

  for (let page = startPage; page < startPage + totalPages; page++) {
    const payload = {
      per_page: 10,
      page: page,
      organization_locations: locations,
      q_organization_keyword_tags: keywords,
    };

    const response = UrlFetchApp.fetch("https://api.apollo.io/v1/mixed_companies/search", {
      method: "post",
      contentType: "application/json",
      headers: { "X-Api-Key": CONFIG.apolloApiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const data = JSON.parse(response.getContentText());
    const orgs = data.accounts || data.organizations || [];

    Logger.log("Total entries for this search: " + data.pagination?.total_entries);
    Logger.log(`Page ${page} returned: ${orgs.length} results`);

    if (!orgs.length) break;

    orgs.forEach(org => {
      if (allOrgs.length >= CONFIG.leadsPerDay) return;
      const name       = org.name || "";
      const normalized = normalizeCompanyName(name);
      if (normalized && !existing.has(normalized)) {
        allOrgs.push(org);
        existing.add(normalized);
      }
    });

    if (allOrgs.length >= CONFIG.leadsPerDay) break;

    Utilities.sleep(500);
  }

  // ── Advance rotation index for next run ──
  advanceRotationIndex(ss, rotationIndex);

  Logger.log("Total orgs collected: " + allOrgs.length);
  return allOrgs;
}

// ─────────────────────────────────────────────
//  FILTER COMPANIES BY NAME
// ─────────────────────────────────────────────
function filterByName(orgs, criteria) {
  const excludeKeywords = (criteria["Exclude Keywords"] || "")
    .split(",")
    .map(k => k.trim().toLowerCase())
    .filter(Boolean);

  if (!excludeKeywords.length) return orgs;

  return orgs.filter(org => {
    const name = (org.name || "").toLowerCase();
    const excluded = excludeKeywords.find(k => name.includes(k));
    if (excluded) {
      Logger.log(`Name filter removed: ${org.name} — contains "${excluded}"`);
      return false;
    }
    return true;
  });
}

// ─────────────────────────────────────────────
//  WEB SCRAPER — filter by AUM and property count
// ─────────────────────────────────────────────

function scrapeAndFilter(orgs) {
  const maxAUM        = CONFIG.scraping.maxAUM;
  const maxProperties = CONFIG.scraping.maxProperties;
  const passed        = [];

  orgs.forEach(org => {
    const website = org.website_url || org.primary_domain;
    if (!website) {
      org._scrapedAUM        = "";
      org._scrapedPropCount  = null;
      passed.push(org);
      return;
    }

    try {
      const result = fetchWithUrlFallback(website);
      if (!result) {
        Logger.log(`All URL attempts failed for ${org.name}`);
        org._scrapedAUM       = "";
        org._aumSource        = "Unknown";
        org._scrapedPropCount = null;
        passed.push(org);
        return;
        }

      const { url: baseUrl, response } = result;
      const rawHtml = response.getContentText();
      const text    = rawHtml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").toUpperCase();

      // ── AUM check ──
      let scrapedAUM = "";
      let aumSource = "Unknown";
      
      if (AUM_KEYWORDS.some(k => text.includes(k))) {
        const aumValue = extractDollarAmount(text, AUM_KEYWORDS);
        if (aumValue !== null) {
          // ✅ Store the raw AUM string for the sheet
          scrapedAUM = aumValue >= 1000
            ? `$${(aumValue / 1000).toFixed(1)}B`
            : `$${aumValue}M`;
          aumSource = "Website";
          if (aumValue > maxAUM) {
            Logger.log(`Filtered out ${org.name} — AUM ${scrapedAUM} exceeds $${maxAUM}M`);
            return;
          }
        }
      }

      // ── Property count check from homepage text ──
      let propCount = null;
      if (PROPERTY_KEYWORDS.some(k => text.includes(k))) {
        propCount = extractPropertyCount(text);
        if (propCount !== null && propCount > maxProperties) {
          Logger.log(`Filtered out ${org.name} — ${propCount} properties exceeds ${maxProperties}`);
          return;
        }
      }

      // ── Portfolio page image count if no count found yet ──
      if (propCount === null) {
        const portfolioUrl = findPortfolioUrl(baseUrl, rawHtml);
        if (portfolioUrl) {
          const portfolioCount = countPortfolioImages(portfolioUrl);
          if (portfolioCount !== null) {
            propCount = portfolioCount;
            if (!scrapedAUM) aumSource = "Portfolio Page";
            if (portfolioCount > maxProperties) {
              Logger.log(`Filtered out ${org.name} — ${portfolioCount} portfolio images exceeds ${maxProperties}`);
              return;
            }
          }
        }
      }

      // ✅ Attach scraped data to org object before passing through
      org._scrapedAUM       = scrapedAUM;
      org._aumSource        = scrapedAUM ? aumSource: "Unknown";
      org._scrapedPropCount = propCount;
      passed.push(org);

    } catch(e) {
      Logger.log(`Scrape failed for ${org.name}: ${e.message}`);
      org._scrapedAUM       = "";
      org._scrapedPropCount = null;
      passed.push(org);
    }

    Utilities.sleep(300);
  });

  Logger.log(`Scraping complete — ${passed.length} of ${orgs.length} passed filters`);
  return passed;
}

// ─────────────────────────────────────────────
//  EXTRACT DOLLAR AMOUNT near a keyword (in millions)
// ─────────────────────────────────────────────
function extractDollarAmount(text, keywords) {
  for (const keyword of keywords) {
    const keywordIndex = text.indexOf(keyword);
    if (keywordIndex === -1) continue;

    const surrounding = text.substring(
      Math.max(0, keywordIndex - 150),
      Math.min(text.length, keywordIndex + 300)
    );

    const B = "(?:BILLION|BN|B[N+]?\\b)";  // matches B, BN, B+, BILLION
    const M = "(?:MILLION|MN|M[N+]?\\b)";  // matches M, MN, M+, MILLION
    const AMT = "(\\d+(?:\\.\\d+)?)\\+?";  // captures number, ignores optional trailing +

    const patterns = [
      new RegExp(`\\$\\s*${AMT}\\s*${B}`, "g"),                               // $1B / $1BN / $1B+
      new RegExp(`\\$\\s*${AMT}\\s*${M}`, "g"),                               // $300M / $300MN
      new RegExp(`${AMT}\\s*${B}\\s*(?:IN\\s*)?${keyword}`, "g"),             // 1B in [keyword]
      new RegExp(`${AMT}\\s*${M}\\s*(?:IN\\s*)?${keyword}`, "g"),             // 300M in [keyword]
      new RegExp(`${keyword}[^$\\d]*\\$?\\s*${AMT}\\s*${B}`, "g"),            // [keyword] of $1B
      new RegExp(`${keyword}[^$\\d]*\\$?\\s*${AMT}\\s*${M}`, "g"),            // [keyword] of $300M
    ];

    for (const pattern of patterns) {
      pattern.lastIndex = 0;
      const match = pattern.exec(surrounding);
      if (match) {
        const value     = parseFloat(match[1]);
        const matchText = match[0].toUpperCase();
        const isBillion = matchText.includes("BILLION") || matchText.includes("BN") || /B[N+]?\b/.test(matchText);
        const normalized = isBillion ? value * 1000 : value;
        Logger.log(`AUM match for "${keyword}": "${match[0]}" → $${normalized}M`);
        return normalized;
      }
    }
  }
  return null;
}

// ─────────────────────────────────────────────
//  EXTRACT PROPERTY COUNT near "properties" keyword
// ─────────────────────────────────────────────
function extractPropertyCount(text) {
  for (const keyword of PROPERTY_KEYWORDS) {
    const keywordIndex = text.indexOf(keyword);
    if (keywordIndex === -1) continue;

    // Search within 200 characters around the keyword
    const surrounding = text.substring(
      Math.max(0, keywordIndex - 100),
      Math.min(text.length, keywordIndex + 200)
    );

    // Build patterns dynamically using the keyword
    const patterns = [
      // 50 properties / 50+ properties
      new RegExp(`(\\d+)\\+?\\s*${keyword}`, "g"),
      // over 100 properties
      new RegExp(`OVER\\s*(\\d+)\\s*${keyword}`, "g"),
      // more than 50 properties
      new RegExp(`MORE\\s*THAN\\s*(\\d+)\\s*${keyword}`, "g"),
      // portfolio of 50
      new RegExp(`${keyword}\\s*OF\\s*(\\d+)`, "g"),
      // 50 assets in portfolio
      new RegExp(`(\\d+)\\s*${keyword}\\s*IN\\s*(?:OUR\\s*)?(?:PORTFOLIO|PIPELINE)`, "g"),
      // across 50 properties
      new RegExp(`ACROSS\\s*(\\d+)\\s*${keyword}`, "g"),
      // managing 50 properties
      new RegExp(`(?:MANAGING|OWNED?|OPERATED?)\\s*(\\d+)\\s*${keyword}`, "g"),
    ];

    for (const pattern of patterns) {
      pattern.lastIndex = 0;
      const match = pattern.exec(surrounding);
      if (match) {
        const count = parseInt(match[1]);
        Logger.log(`Property count match for "${keyword}": "${match[0]}" → ${count}`);
        return count;
      }
    }
  }

  return null;
}

// ─────────────────────────────────────────────
//  FIND PORTFOLIO/PROPERTIES PAGE URL
// ─────────────────────────────────────────────
function findPortfolioUrl(baseUrl, html) {
  const linkPattern = /href=["']([^"']*(?:portfolio|properties|our-properties|our-portfolio|projects|developments)[^"']*)["']/gi;
  const matches = [];
  let match;
  while ((match = linkPattern.exec(html)) !== null) {
    matches.push(match[1]);
  }
  if (!matches.length) return null;

  const href = matches[0];
  if (href.startsWith("http")) return href;
  if (href.startsWith("/")) {
    // Extract origin manually — new URL() doesn't exist in Apps Script
    const originMatch = baseUrl.match(/^(https?:\/\/[^\/]+)/);
    if (!originMatch) return null;
    return originMatch[1] + href;
  }
  return baseUrl.replace(/\/$/, "") + "/" + href;
}

// ─────────────────────────────────────────────
//  COUNT IMAGES ON PORTFOLIO PAGE
// ─────────────────────────────────────────────
function countPortfolioImages(url) {
  const MAX_PAGES = 10; // safety cap to avoid infinite loops / runaway fetches
  
  try {
    let totalCount  = 0;
    let currentUrl  = url;
    let pagesFetched = 0;
    const visitedUrls = new Set();

    while (currentUrl && pagesFetched < MAX_PAGES) {
      // Avoid re-visiting the same page
      if (visitedUrls.has(currentUrl)) break;
      visitedUrls.add(currentUrl);

      const result = fetchWithUrlFallback(currentUrl);
      if (!result) {
        Logger.log(`Failed to fetch portfolio page: ${currentUrl}`);
        break;
      }

      const html      = result.response.getContentText();
      const propImages = filterPropertyImages(html);
      totalCount += propImages;
      pagesFetched++;

      Logger.log(`Portfolio page ${currentUrl} (page ${pagesFetched}) — ${propImages} property-like images`);

      // Find the "next page" link and continue
      currentUrl = findNextPageUrl(currentUrl, html);
      if (currentUrl) Utilities.sleep(300); // be polite between pages
    }

    Logger.log(`Portfolio total — ${totalCount} property-like images across ${pagesFetched} page(s)`);
    return totalCount;

  } catch(e) {
    Logger.log(`Portfolio page scrape failed: ${e.message}`);
    return null;
  }
}

// ── Extracts property-like image count from an HTML string ──
function filterPropertyImages(html) {
  const imgPattern = /<img[^>]+>/gi;
  const allImages  = html.match(imgPattern) || [];
  const skip = ["icon", "logo", "sprite", "pixel", "tracking",
                "avatar", "arrow", "button", "badge", "1x1"];

  return allImages.filter(tag => {
    const src = (tag.match(/src=["']([^"']+)["']/i) || [])[1] || "";
    if (skip.some(s => src.toLowerCase().includes(s))) return false;
    if (src.startsWith("data:") || src.length < 10)   return false;
    return true;
  }).length;
}

// ── Finds the URL of the next paginated page, or returns null ──
function findNextPageUrl(currentUrl, html) {
  // Match common "next page" link patterns
  const nextPattern = /<a[^>]+href=["']([^"']+)["'][^>]*>[\s\S]*?(?:next|›|»|→|load more)[\s\S]*?<\/a>/gi;
  const match = nextPattern.exec(html);
  if (!match) return null;

  const href = match[1];

  // Resolve to absolute URL
  if (href.startsWith("http"))  return href;
  const originMatch = currentUrl.match(/^(https?:\/\/[^\/]+)/);
  if (!originMatch) return null;
  if (href.startsWith("/"))     return originMatch[1] + href;
                                return originMatch[1] + "/" + href;
}
// ─────────────────────────────────────────────
//  ROTATION HELPERS
// ─────────────────────────────────────────────
function getOrCreateSettings(ss) {
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.getRange("A1:B1").setValues([["Key", "Value"]]);
    sheet.getRange("A2:B2").setValues([["rotationIndex", "0"]]);
    styleHeader(sheet, 2);
  }

  const data     = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  settings.rotationIndex = parseInt(settings.rotationIndex || 0);
  return settings;
}

function advanceRotationIndex(ss, currentIndex) {
  const sheet   = ss.getSheetByName("Settings");
  const data    = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "rotationIndex") {
      sheet.getRange(i + 1, 2).setValue(currentIndex + 1);
      break;
    }
  }
}

// ─────────────────────────────────────────────
//  APOLLO — enrich org, find acquisitions contact
// ─────────────────────────────────────────────

function enrichOrg(org, criteria) {
  // ✅ Read target roles dynamically from Criteria sheet
  const rawRoles    = (criteria["Target Role"] || "").split(",").map(r => r.trim()).filter(Boolean);
  const searchTitles = rawRoles.length > 0 ? rawRoles : [
    "Head of Acquisitions", "Director of Acquisitions",
    "VP Acquisitions", "VP of Acquisitions", "Acquisitions Manager",
    "Acquisitions Director", "Acquisitions Associate", "Investment Manager",
    "Director of Investments", "VP of Investments",
    "Chief Investment Officer", "Vice President of Acquisitions, Investment Manager, Principal, Managing Partner, Partner",
  ];

  // Build priority list from Criteria sheet (lowercase for matching)
  const titlePriority = searchTitles.map(t => t.toLowerCase());

  const searchPayload = {
    per_page: 5,
    page: 1,
    organization_ids: [org.id],
    titles: searchTitles,
  };

  const searchResponse = UrlFetchApp.fetch("https://api.apollo.io/v1/mixed_people/api_search", {
    method: "post",
    contentType: "application/json",
    headers: { "X-Api-Key": CONFIG.apolloApiKey },
    payload: JSON.stringify(searchPayload),
    muteHttpExceptions: true,
  });

  const searchData = JSON.parse(searchResponse.getContentText());
  const people     = searchData.people || [];

  // Pick highest priority contact
  let bestContact = null;
  let bestRank    = Infinity;

  people.forEach(person => {
    const title         = (person.title || "").toLowerCase().trim();
    const rank          = titlePriority.findIndex(t => title.includes(t));
    const effectiveRank = rank === -1 ? Infinity : rank;
    if (effectiveRank < bestRank) {
      bestRank    = effectiveRank;
      bestContact = person;
    }
  });

  return {
    company:     org.name        || "",
    website:     org.website_url || "",
    location:    [org.city, org.state, org.country].filter(Boolean).join(", "),
    aum:         org._scrapedAUM  || "",        // ✅ website AUM → sheet
    aumSource:   org._aumSource || "Unknown",
    properties:  org._scrapedPropCount !== null ? org._scrapedPropCount : "Unknown",
    revenue:     org.organization_revenue_printed
                   || org.market_cap
                   || (org.organization_revenue
                       ? `$${(org.organization_revenue / 1e6).toFixed(1)}M`
                       : "Unknown"),            // ✅ Apollo revenue → Claude only
    propCount:   org._scrapedPropCount || "",   // ✅ property count → Claude only
    contactName: bestContact ? `${bestContact.first_name || ""} ${bestContact.last_name || ""}`.trim() : "",
    email:       bestContact?.email || "",
    phone:       bestContact?.phone_numbers?.[0]?.raw_number || "",
    title:       bestContact?.title || "",
  };
}

function debugDynamic() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const criteria = readCriteria(ss);
  
  Logger.log("Criteria loaded: " + JSON.stringify(criteria));

  const locations       = (criteria["Locations"] || "United States").split(",").map(l => l.trim());
  const keywords        = (criteria["Keywords"] || "real estate").split(",").map(k => k.trim());
  const excludeKeywords = (criteria["Exclude Keywords"] || "").split(",").map(k => k.trim()).filter(Boolean);

  Logger.log("Locations: " + JSON.stringify(locations));
  Logger.log("Keywords: " + JSON.stringify(keywords));
  Logger.log("Exclude Keywords: " + JSON.stringify(excludeKeywords));

  const payload = {
    per_page: 5,
    page: 1,
    organization_locations: locations,
    q_organization_keyword_tags: keywords,
    ...(excludeKeywords.length > 0 && { q_organization_keyword_tags_excluded: excludeKeywords }),
  };

  const response = UrlFetchApp.fetch("https://api.apollo.io/v1/mixed_companies/search", {
    method: "post",
    contentType: "application/json",
    headers: { "X-Api-Key": CONFIG.apolloApiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const data = JSON.parse(response.getContentText());
  Logger.log("Total results: " + data.pagination?.total_entries);
  Logger.log("Accounts returned: " + (data.accounts || data.organizations || []).length);
  Logger.log("First result: " + JSON.stringify((data.accounts || data.organizations || [])[0]?.name));
}

function debugEnrich() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const criteria = readCriteria(ss);

  const locations = (criteria["Locations"] || "United States").split(",").map(l => l.trim());
  const keywords  = (criteria["Keywords"] || "real estate").split(",").map(k => k.trim());

  const payload = {
    per_page: 5,
    page: 1,
    organization_locations: locations,
    q_organization_keyword_tags: keywords,
  };

  const response = UrlFetchApp.fetch("https://api.apollo.io/v1/mixed_companies/search", {
    method: "post",
    contentType: "application/json",
    headers: { "X-Api-Key": CONFIG.apolloApiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const data = JSON.parse(response.getContentText());
  const org  = (data.accounts || data.organizations || [])[0];

  Logger.log("Testing enrichment on: " + org.name);

  try {
    const enriched = enrichOrg(org, criteria);
    Logger.log("Enriched result: " + JSON.stringify(enriched));
  } catch(e) {
    Logger.log("Enrichment error: " + e.message);
  }
}

// ─────────────────────────────────────────────
//  CLAUDE — score lead
// ─────────────────────────────────────────────
function scoreLead(lead, criteria, feedbackHistory) {
  const feedbackSummary = feedbackHistory.length
    ? feedbackHistory.slice(0, 20).map(f =>
        `- ${f.company}: ${f.status} (score ${f.score})${f.notes ? " — " + f.notes : ""}`
      ).join("\n")
    : "No feedback history yet.";

  const prompt = `You are a prospecting agent for a real estate financing company. Score this lead 1-10 based on how well it matches the ideal client criteria. Return ONLY a JSON object with two fields: "score" (integer 1-10) and "reasoning" (one sentence max).

    IDEAL CLIENT CRITERIA:
    ${JSON.stringify(criteria, null, 2)}

    PAST FEEDBACK TO LEARN FROM:
    ${feedbackSummary}

    LEAD TO SCORE:
    Company: ${lead.company}
    Website: ${lead.website}
    Location: ${lead.location}
    AUM (from website): ${lead.aum || "Unknown"}
    Annual Revenue (from Apollo): ${lead.revenue || "Unknown"}
    Property Count: ${lead.propCount || "Unknown"}
    Contact: ${lead.contactName} (${lead.title})
    Email: ${lead.email}
    Phone: ${lead.phone}

    Return only valid JSON, no markdown, no explanation.`;

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": CONFIG.claudeApiKey,
      "anthropic-version": "2023-06-01",
    },
    payload: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 200,
      messages: [{ role: "user", content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  let score     = 5;
  let reasoning = "Could not score.";

  try {
    const data   = JSON.parse(response.getContentText());
    const text   = data.content?.[0]?.text || "{}";
    const parsed = JSON.parse(text);
    score        = parsed.score     || 5;
    reasoning    = parsed.reasoning || "No reasoning provided.";
  } catch (e) {
    reasoning = "Scoring error: " + e.message;
  }

  return { ...lead, score, reasoning };
}

// ─────────────────────────────────────────────
//  WRITE LEADS TO SHEET
// ─────────────────────────────────────────────
function writeLeads(ss, leads) {
  const sheet = ss.getSheetByName(CONFIG.sheets.leads);
  const today = new Date().toLocaleDateString();

  const rows = leads.map(l => [
    today,
    l.company,
    l.website,
    l.location,
    l.aum,
    l.aumSource,
    l.properties,
    l.contactName,
    l.email,
    l.phone,
    l.score,
    l.reasoning,
    "", "", // Status, Feedback Notes
  ]);

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, CONFIG.leadColumns.length).setValues(rows);

  // Dropdown for Status column (col 11)
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["✅ Good", "❌ Bad", "🔄 Maybe"], true)
    .build();
  sheet.getRange(startRow, 11, rows.length, 1).setDataValidation(rule);

  // Color-code rows by score
  rows.forEach((row, i) => {
    const score    = row[8];
    const rowRange = sheet.getRange(startRow + i, 1, 1, CONFIG.leadColumns.length);
    if      (score >= 8) rowRange.setBackground("#e6f4ea"); // green
    else if (score >= 5) rowRange.setBackground("#fff8e1"); // yellow
    else                 rowRange.setBackground("#fce8e6"); // red
  });
}

// ─────────────────────────────────────────────
//  SIDEBAR: get pending leads for review
// ─────────────────────────────────────────────
function getPendingLeads() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(CONFIG.sheets.leads);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, CONFIG.leadColumns.length).getValues();
  return data
    .map((row, i) => ({
      rowIndex:  i + 2,
      date:      row[0],  company:   row[1],  website:   row[2],
      location:  row[3],  aum:       row[4],  contact:   row[5],
      email:     row[6],  phone:     row[7],  score:     row[8],
      reasoning: row[9],  status:    row[10], notes:     row[11],
    }))
    .filter(r => !r.status); // only unreviewed
}

// ─────────────────────────────────────────────
//  SIDEBAR: save feedback
// ─────────────────────────────────────────────
function saveFeedback(rowIndex, status, notes) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const leads  = ss.getSheetByName(CONFIG.sheets.leads);
  const fb     = ss.getSheetByName(CONFIG.sheets.feedback);

  leads.getRange(rowIndex, 11).setValue(status);
  leads.getRange(rowIndex, 12).setValue(notes || "");

  const row = leads.getRange(rowIndex, 1, 1, CONFIG.leadColumns.length).getValues()[0];
  fb.appendRow([
    new Date().toLocaleDateString(),
    row[1], status, row[8], notes || "",
    JSON.stringify(readCriteria(ss)),
  ]);

  return { success: true };
}

// ─────────────────────────────────────────────
//  DAILY TRIGGER — run once to activate
// ─────────────────────────────────────────────
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "runAgent") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("runAgent")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  SpreadsheetApp.getUi().alert("✅ Daily trigger set! The agent will run every day at 8am.");
}

// ─────────────────────────────────────────────
//  EXPORT FEEDBACK TO S3 AS TRAINING DATA
// ─────────────────────────────────────────────
function exportTrainingData() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(CONFIG.sheets.feedback);
  const lastRow  = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No feedback data yet. Review some leads first.");
    return;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // Filter to only Good/Bad decisions (skip Maybe for cleaner training data)
  const labeled = rows.filter(r => r[2] === "✅ Good" || r[2] === "❌ Bad");

  if (labeled.length < 10) {
    SpreadsheetApp.getUi().alert(`Only ${labeled.length} labeled leads found. Aim for 500+ before fine-tuning.`);
    return;
  }

  // Format as JSONL (one training example per line)
  // This is the format SageMaker and most fine-tuning services expect
  const jsonl = labeled.map(row => {
    const company  = row[1];
    const status   = row[2];
    const score    = row[3];
    const notes    = row[4];
    const criteria = row[5];

    const prompt = `You are a lead scoring agent for a commercial real estate financing company. Score this lead.
Criteria: ${criteria}
Company: ${company}
Return only JSON with "score" (1-10) and "reasoning" (one sentence).`;

    const completion = JSON.stringify({
      score:     score,
      reasoning: notes || (status === "✅ Good"
        ? "Strong match for target criteria."
        : "Does not meet target criteria."),
    });

    return JSON.stringify({
      prompt:     prompt,
      completion: completion,
      label:      status,
    });
  }).join("\n");

  // Upload to S3
  const timestamp = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd-HH-mm");
  const fileName  = `training-data/leads-training-${timestamp}.jsonl`;
  const rawExport = `feedback-exports/raw-export-${timestamp}.jsonl`;

  uploadToS3(jsonl, fileName);
  uploadToS3(jsonl, rawExport);

  SpreadsheetApp.getUi().alert(
    `✅ Exported ${labeled.length} training examples to S3.\n\nFile: ${fileName}\n\nYou need ${Math.max(0, 500 - labeled.length)} more labeled leads before fine-tuning.`
  );
}

// ─────────────────────────────────────────────
//  HELPERS
// ─────────────────────────────────────────────

//  AWS SIGNING HELPERS
function computeSHA256(message) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    message,
    Utilities.Charset.UTF_8
  );
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function computeHMACSHA256(key, message) {
  return Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    message,
    key,
    Utilities.Charset.UTF_8
  );
}

function computeHMACSHA256Hex(key, message) {
  const sig = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256,
    message,
    key,
    Utilities.Charset.UTF_8
  );
  return sig.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function getSigningKey(secretKey, dateStamp, region, service) {
  const kDate    = computeHMACSHA256("AWS4" + secretKey, dateStamp);
  const kRegion  = computeHMACSHA256(kDate,    region);
  const kService = computeHMACSHA256(kRegion,  service);
  return           computeHMACSHA256(kService, "aws4_request");
}

function styleHeader(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground("#3C3489");
  range.setFontColor("#ffffff");
  range.setFontWeight("bold");
}

//  S3 UPLOAD HELPER (AWS Signature V4)

function uploadToS3(content, filePath) {
  const aws        = CONFIG.aws;
  const service    = "s3";
  const host       = `${aws.bucket}.s3.${aws.region}.amazonaws.com`;
  const endpoint   = `https://${host}/${filePath}`;
  const now        = new Date();
  const amzDate    = Utilities.formatDate(now, "UTC", "yyyyMMdd'T'HHmmss'Z'");
  const dateStamp  = Utilities.formatDate(now, "UTC", "yyyyMMdd");
  const payloadHash = computeSHA256(content);

  // Build canonical request
  const canonicalHeaders = `host:${host}\nx-amz-content-sha256:${payloadHash}\nx-amz-date:${amzDate}\n`;
  const signedHeaders    = "host;x-amz-content-sha256;x-amz-date";
  const canonicalRequest = [
    "PUT",
    "/" + filePath,
    "",
    canonicalHeaders,
    signedHeaders,
    payloadHash,
  ].join("\n");

  // Build string to sign
  const credentialScope = `${dateStamp}/${aws.region}/${service}/aws4_request`;
  const stringToSign    = [
    "AWS4-HMAC-SHA256",
    amzDate,
    credentialScope,
    computeSHA256(canonicalRequest),
  ].join("\n");

  // Calculate signature
  const signingKey  = getSigningKey(aws.secretAccessKey, dateStamp, aws.region, service);
  const signature   = computeHMACSHA256Hex(signingKey, stringToSign);
  const authHeader  = `AWS4-HMAC-SHA256 Credential=${aws.accessKeyId}/${credentialScope}, SignedHeaders=${signedHeaders}, Signature=${signature}`;

  // Make request
  const response = UrlFetchApp.fetch(endpoint, {
    method: "put",
    contentType: "application/x-ndjson",
    headers: {
      "Authorization":          authHeader,
      "x-amz-date":             amzDate,
      "x-amz-content-sha256":   payloadHash,
    },
    payload: content,
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("S3 upload failed: " + response.getContentText());
  }
}

//  NORMALIZE COMPANY NAME FOR DEDUPLICATION
function normalizeCompanyName(name) {
  return name
    .toString()
    .toLowerCase()
    .trim()
    // Remove common legal suffixes
    .replace(/\b(llc|inc|corp|ltd|lp|llp|co|company|group|holdings|partners|properties|realty|capital|management|advisors|associates|enterprises|international|global)\b/g, "")
    // Remove punctuation and special characters
    .replace(/[^a-z0-9\s]/g, "")
    // Collapse multiple spaces
    .replace(/\s+/g, " ")
    .trim();
}

//  NORMALIZE AND FETCH URL WITH FALLBACK
function fetchWithUrlFallback(website) {
  const options = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { "User-Agent": "Mozilla/5.0" },
  };

  // Build list of URLs to try in order
  const base    = website.replace(/^https?:\/\//i, "").replace(/\/+$/, "");
  const hasWww  = base.toLowerCase().startsWith("www.");
  const urlsToTry = [
    "https://" + base,                                      // original
    hasWww
      ? "https://" + base.replace(/^www\./i, "")           // remove www
      : "https://www." + base,                             // add www
    "http://" + base,                                       // try http
  ];

  for (const url of urlsToTry) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code     = response.getResponseCode();
      if (code === 200) {
        Logger.log(`✅ Connected to ${url}`);
        return { url, response };
      } 
      Logger.log(`⚠️ ${url} returned ${code}`);
    } catch(e) {
      Logger.log(`❌ Failed to fetch ${url}: ${e.message}`);
    }
  }

  return null; // all attempts failed
}
