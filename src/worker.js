const MAX_POINTS = 10;
const NUMBER_PATTERN = /(\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?:[kKmMbBtT%])?/g;

const DEFAULT_META = {
  eyebrow: "Globalite Macro Brief",
  main_title: "WEEKLY TOP 10 ARGUMENTS",
  subtitle: "A clear weekly macro summary with the key arguments that matter.",
  block_height: "925000",
  max_supply_btc: "21000000",
  circulating_supply_btc: "19960000",
  hashrate_eh_s: "820",
  hashrate_scale_eh_s: "1000",
  snapshot_title: "At The Time Of Writing",
  snapshot_intro:
    "At the time of writing, these on-chain supply anchors provide the baseline context.",
  snapshot_note: "Figures are rounded and updated with each issue.",
  tldr_title: "TL;DR",
  tldr_content:
    "Leverage reset first, liquidity expanded next, and structural adoption kept building.",
  conclusion_title: "GLOBALITE CONCLUSION",
  conclusion_content:
    "For deeper context on these points, visit globalite.co.\nOur team tracks macro shifts, liquidity, and positioning every week.",
  cta_url: "https://globalite.co",
  cta_label: "globalite.co",
  address_line: "Globalite, Lugano, Piazza dell'Indipendenza 3, CAP 6901",
  footer_line: "Globalite Macro Brief - For internal distribution.",
  hero_image_url: "/hero.png",
  footer_logo_url: "/logotosite.png",
  footer_instagram_icon: "/instagram.png",
  footer_x_icon: "/x:twitter.png",
  footer_linkedin_icon: "/linkedin.png",
  image_dir: ".",
  auto_image_by_order: "true",
  logo_url: "/brand_orange_bg_transparent@2xSite.svg",
};

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname.startsWith("/img/")) {
      return serveR2Image(url, env);
    }

    if (url.pathname === "/health") {
      return new Response(
        JSON.stringify({ status: "ok", timestamp: new Date().toISOString() }),
        {
          headers: {
            "content-type": "application/json; charset=utf-8",
            "cache-control": "no-store",
          },
        }
      );
    }

    const pagePaths = new Set(["/", "/newsletter", "/newsletter.html"]);
    if (!pagePaths.has(url.pathname)) {
      if (env.ASSETS) {
        return env.ASSETS.fetch(request);
      }
      return new Response("Not found", { status: 404 });
    }

    const ttl = normalizeTtl(env.CACHE_TTL_SECONDS);
    const forceRefresh =
      url.searchParams.get("force") === "1" ||
      request.headers.get("cache-control") === "no-cache";
    const cacheKey = new Request(`${url.origin}/__newsletter_cache_v1__`);

    if (!forceRefresh) {
      const cached = await caches.default.match(cacheKey);
      if (cached) {
        return cached;
      }
    }

    try {
      const sheetId = normalizeText(env.GOOGLE_SHEET_ID);
      if (!sheetId) {
        throw new Error("GOOGLE_SHEET_ID is missing in Worker vars.");
      }

      const metaTab = normalizeText(env.GOOGLE_META_TAB) || "meta";
      const pointsTab = normalizeText(env.GOOGLE_POINTS_TAB) || "points";
      const livePricesTab =
        normalizeText(env.GOOGLE_LIVE_PRICES_TAB) || "live_prices";
      const btcPriceTab =
        normalizeText(env.GOOGLE_BTC_PRICE_TAB) || "BTC Price";
      const treasuriesTab =
        normalizeText(env.GOOGLE_TREASURIES_TAB) || "Treasuries";
      const circulatingTab =
        normalizeText(env.GOOGLE_CIRCULATING_TAB) || "Circulating BTC";
      const liquidationsTab =
        normalizeText(env.GOOGLE_LIQUIDATIONS_TAB) || "Liquidations";
      const ownershipTab =
        normalizeText(env.GOOGLE_OWNERSHIP_TAB) || "Distribution";
      const graphSettingsTab =
        normalizeText(env.GOOGLE_GRAPH_SETTINGS_TAB) || "graph_settings";

      const [
        metaRows,
        pointsRows,
        livePriceRows,
        btcPriceRows,
        treasuriesRows,
        circulatingRows,
        liquidationsRows,
        ownershipRows,
        graphSettingsRows,
      ] =
        await Promise.all([
          fetchGoogleTabRows(sheetId, metaTab, true),
          fetchGoogleTabRows(sheetId, pointsTab, true),
          fetchGoogleLiveBtcRows(sheetId, livePricesTab, false),
          fetchGoogleTabRows(sheetId, btcPriceTab, false),
          fetchGoogleTabRows(sheetId, treasuriesTab, false),
          fetchGoogleTabRows(sheetId, circulatingTab, false),
          fetchGoogleTabRows(sheetId, liquidationsTab, false),
          fetchGoogleTabRows(sheetId, ownershipTab, false),
          fetchGoogleTabRows(sheetId, graphSettingsTab, false),
        ]);

      const meta = readMeta(metaRows);
      const points = readPoints(pointsRows);
      const graphSettingsMap = graphSettingsRows.length
        ? safeOptionalParse(graphSettingsTab, () => readGraphSettingsMap(graphSettingsRows), {})
        : {};
      const treasuriesTopN = null;
      const liquidationsTopN = toPositiveInt(
        graphSettingsMap.liquidations && graphSettingsMap.liquidations.top_n,
        6
      );
      const liveBtc = livePriceRows.length
        ? safeOptionalParse(livePricesTab, () => readLiveBtcPrice(livePriceRows), null)
        : null;
      let btcPricePoints = btcPriceRows.length
        ? safeOptionalParse(btcPriceTab, () => readBtcPricePoints(btcPriceRows), [])
        : [];
      if (!btcPricePoints.length && livePriceRows.length) {
        btcPricePoints = safeOptionalParse(
          livePricesTab,
          () => readBtcPricePointsFromLivePrices(livePriceRows),
          []
        );
      }
      const treasuryBars = treasuriesRows.length
        ? safeOptionalParse(
            treasuriesTab,
            () => readTreasuryBars(treasuriesRows, treasuriesTopN),
            []
          )
        : [];
      let circulatingMetric = circulatingRows.length
        ? safeOptionalParse(circulatingTab, () => readCirculatingMetric(circulatingRows), null)
        : null;
      const liquidationBars = liquidationsRows.length
        ? safeOptionalParse(
            liquidationsTab,
            () => readLiquidationBars(liquidationsRows, liquidationsTopN),
            []
          )
        : [];
      const ownershipSegments = ownershipRows.length
        ? safeOptionalParse(ownershipTab, () => readOwnershipSegments(ownershipRows), [])
        : [];
      if (!circulatingMetric) {
        const maxSupply = parseNumber(meta.max_supply_btc, 21_000_000);
        const circulatingSupply = parseNumber(meta.circulating_supply_btc, 0);
        if (circulatingSupply > 0 && maxSupply > 0) {
          circulatingMetric = {
            as_of_date: "",
            circulating_supply_btc: circulatingSupply,
            max_supply_btc: maxSupply,
            note: "",
          };
        }
      }

      const html = renderHtml(
        meta,
        points,
        btcPricePoints,
        treasuryBars,
        circulatingMetric,
        liquidationBars,
        ownershipSegments,
        graphSettingsMap,
        liveBtc,
        {
        useR2Images: Boolean(env.IMAGES),
        r2ImagePrefix: normalizeText(env.R2_IMAGE_PREFIX) || "image",
        r2ImageExt: normalizeText(env.R2_IMAGE_EXT) || "jpg",
        }
      );
      const response = new Response(html, {
        headers: {
          "content-type": "text/html; charset=utf-8",
          "cache-control": `public, max-age=${ttl}, s-maxage=${ttl}`,
        },
      });
      ctx.waitUntil(caches.default.put(cacheKey, response.clone()));
      return response;
    } catch (error) {
      return new Response(`Newsletter render error: ${error.message}`, {
        status: 500,
        headers: {
          "content-type": "text/plain; charset=utf-8",
          "cache-control": "no-store",
        },
      });
    }
  },
};

async function serveR2Image(url, env) {
  if (!env.IMAGES) {
    return new Response("R2 binding not configured.", { status: 404 });
  }

  const key = decodeURIComponent(url.pathname.slice("/img/".length));
  if (!key || key.includes("..")) {
    return new Response("Invalid image key.", { status: 400 });
  }

  const object = await env.IMAGES.get(key);
  if (!object) {
    return new Response("Not found", { status: 404 });
  }

  const headers = new Headers();
  if (typeof object.writeHttpMetadata === "function") {
    object.writeHttpMetadata(headers);
  }
  headers.set("etag", object.httpEtag);
  if (!headers.has("content-type")) {
    headers.set("content-type", guessContentType(key));
  }
  headers.set("cache-control", "public, max-age=300");

  return new Response(object.body, { headers });
}

function guessContentType(key) {
  const lower = key.toLowerCase();
  if (lower.endsWith(".png")) return "image/png";
  if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
  if (lower.endsWith(".webp")) return "image/webp";
  if (lower.endsWith(".gif")) return "image/gif";
  if (lower.endsWith(".svg")) return "image/svg+xml";
  return "application/octet-stream";
}

function normalizeTtl(value) {
  const parsed = Number.parseInt(String(value ?? "120"), 10);
  if (!Number.isFinite(parsed)) {
    return 120;
  }
  return Math.max(30, parsed);
}

function safeOptionalParse(tabName, parser, fallback) {
  try {
    return parser();
  } catch (error) {
    console.warn(`Skipping tab '${tabName}': ${error.message}`);
    return fallback;
  }
}

async function fetchGoogleTabRows(sheetId, tabName, required) {
  const url =
    `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?` +
    `tqx=out:csv&sheet=${encodeURIComponent(tabName)}`;
  return fetchGoogleCsvRows(url, tabName, required);
}

async function fetchGoogleLiveBtcRows(sheetId, tabName, required) {
  const query =
    "select A,B,I,J where upper(J) contains 'BTC' or upper(J) contains 'BITCOIN' or upper(J) contains 'XBT' order by A desc limit 600";
  const url =
    `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?` +
    `tqx=out:csv&sheet=${encodeURIComponent(tabName)}&tq=${encodeURIComponent(query)}`;
  return fetchGoogleCsvRows(url, tabName, required);
}

async function fetchGoogleCsvRows(url, tabName, required) {
  const response = await fetch(url, {
    headers: { "accept": "text/csv,text/plain;q=0.9,*/*;q=0.1" },
    cf: { cacheEverything: false },
  });

  if (!response.ok) {
    if (!required && (response.status === 400 || response.status === 404)) {
      return [];
    }
    throw new Error(
      `Could not load Google Sheet tab '${tabName}' (HTTP ${response.status}).`
    );
  }

  const raw = (await response.text()).replace(/^\uFEFF/, "").trim();
  if (!raw) {
    if (required) {
      throw new Error(`Google Sheet tab '${tabName}' is empty.`);
    }
    return [];
  }

  const lowered = raw.toLowerCase();
  if (lowered.startsWith("<!doctype html") || lowered.startsWith("<html")) {
    if (!required) {
      return [];
    }
    throw new Error(
      `Could not read tab '${tabName}'. Confirm sheet link-sharing is enabled.`
    );
  }
  if (
    lowered.includes("google.visualization.query.setresponse") &&
    lowered.includes('"status":"error"')
  ) {
    if (!required) {
      return [];
    }
    throw new Error(`Google Sheets query error for tab '${tabName}'.`);
  }

  return normalizeTableRows(parseCsv(raw));
}

function parseCsv(input) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];

    if (inQuotes) {
      if (ch === '"') {
        if (input[i + 1] === '"') {
          value += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        value += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(value);
      value = "";
    } else if (ch === "\n") {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else if (ch === "\r") {
      // Ignore CR; LF handles row ending.
    } else {
      value += ch;
    }
  }

  row.push(value);
  if (row.length > 1 || row[0] !== "") {
    rows.push(row);
  }

  return rows;
}

function normalizeTableRows(rows) {
  if (!rows.length) {
    return [];
  }
  const width = rows.reduce((max, row) => Math.max(max, row.length), 0);
  return rows.map((row) => row.concat(Array(width - row.length).fill("")));
}

function readMeta(rows) {
  const meta = {};
  for (let i = 1; i < rows.length; i += 1) {
    const key = normalizeText(rows[i][0]);
    const value = normalizeText(rows[i][1]);
    if (key) {
      meta[key] = value;
    }
  }

  for (const [key, value] of Object.entries(DEFAULT_META)) {
    if (!Object.prototype.hasOwnProperty.call(meta, key) || !normalizeText(meta[key])) {
      meta[key] = value;
    }
  }
  return meta;
}

function normalizeGraphKey(value) {
  const key = normalizeText(value).toLowerCase().replace(/-/g, "_").replace(/\s+/g, "_");
  const aliases = {
    btc: "btc_price",
    btcprice: "btc_price",
    btc_price: "btc_price",
    price: "btc_price",
    liquidation: "liquidations",
    liquidations: "liquidations",
    treasury: "treasuries",
    treasuries: "treasuries",
    circulating: "circulating_btc",
    circulating_btc: "circulating_btc",
    circulatingbtc: "circulating_btc",
    ownership: "ownership",
    distribution: "ownership",
  };
  return aliases[key] || key;
}

function readGraphSettingsMap(rows) {
  if (!rows.length) {
    return {};
  }
  const mapping = headerIndexMap(rows, []);
  const keyIdx =
    "graph_key" in mapping ? mapping.graph_key : ("key" in mapping ? mapping.key : -1);
  if (keyIdx < 0) {
    return {};
  }

  const result = {};
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const graphKey = normalizeGraphKey(row[keyIdx]);
    if (!graphKey) {
      continue;
    }
    const showRaw = "show" in mapping ? row[mapping.show] : "yes";
    const show = isRowVisible(showRaw);
    const title = "title" in mapping ? normalizeText(row[mapping.title]) : "";
    const topN =
      "top_n" in mapping
        ? toPositiveInt(row[mapping.top_n], null)
        : ("topn" in mapping ? toPositiveInt(row[mapping.topn], null) : null);
    const maxBars =
      "max_bars" in mapping
        ? toPositiveInt(row[mapping.max_bars], null)
        : ("maxbars" in mapping ? toPositiveInt(row[mapping.maxbars], null) : null);
    const comment = "comment" in mapping ? normalizeText(row[mapping.comment]) : "";
    result[graphKey] = {
      show,
      title,
      top_n: topN,
      max_bars: maxBars,
      comment,
    };
  }
  return result;
}

function headerIndexMap(rows, required) {
  const headers = rows[0] ?? [];
  const mapping = {};
  headers.forEach((header, idx) => {
    const key = normalizeText(header).toLowerCase();
    if (key) {
      mapping[key] = idx;
    }
  });

  const missing = required.filter((key) => !(key in mapping));
  if (missing.length) {
    throw new Error(`Missing required columns: ${missing.join(", ")}`);
  }
  return mapping;
}

function readPoints(rows) {
  if (!rows.length) {
    throw new Error("Points tab is empty.");
  }
  const mapping = headerIndexMap(rows, [
    "order",
    "title",
    "content",
    "image_path",
    "image_caption",
  ]);

  const points = [];
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const orderText = normalizeText(row[mapping.order]);
    const title = normalizeText(row[mapping.title]);
    const content = normalizeText(row[mapping.content]);
    const imagePath = normalizeText(row[mapping.image_path]);
    const imageCaption = normalizeText(row[mapping.image_caption]);
    const source = "source" in mapping ? normalizeText(row[mapping.source]) : "";

    if (!orderText && !title && !content && !imagePath && !imageCaption && !source) {
      continue;
    }

    const order = parseOrder(orderText, i + 1);
    if (!title) {
      throw new Error(`Missing title at points row ${i + 1}.`);
    }
    if (!content) {
      throw new Error(`Missing content at points row ${i + 1}.`);
    }

    points.push({
      order,
      title,
      content,
      image_path: imagePath,
      image_caption: imageCaption,
      source,
    });
  }

  if (!points.length) {
    throw new Error("No points found. Add at least 1 point.");
  }

  points.sort((a, b) => a.order - b.order);
  const orders = points.map((item) => item.order);
  const duplicates = [...new Set(orders.filter((value, idx) => orders.indexOf(value) !== idx))];
  if (duplicates.length) {
    throw new Error(`Duplicate order values found: ${duplicates.join(", ")}`);
  }
  if (points.length > MAX_POINTS) {
    throw new Error(`Found ${points.length} points. Max allowed is ${MAX_POINTS}.`);
  }

  return points;
}

function readLiveBtcPrice(rows) {
  if (!rows.length) {
    return null;
  }
  const mapping = headerIndexMap(rows, ["date", "asset"]);
  const closeIdx = "close" in mapping ? mapping.close : ("price" in mapping ? mapping.price : -1);
  if (closeIdx < 0) {
    return null;
  }
  const currencyIdx = "currency" in mapping ? mapping.currency : -1;
  let latest = null;
  let latestDate = null;

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (!isBtcAsset(row[mapping.asset])) {
      continue;
    }
    const price = parseNumber(row[closeIdx], Number.NaN);
    if (!Number.isFinite(price) || price <= 0) {
      continue;
    }
    const dateRaw = row[mapping.date];
    const dateValue = parseDateValue(dateRaw);
    const candidate = {
      price,
      date: renderDateLabel(dateRaw),
      currency: currencyIdx >= 0 ? normalizeText(row[currencyIdx]) || "USD" : "USD",
    };

    if (!latest) {
      latest = candidate;
      latestDate = dateValue;
      continue;
    }

    if (dateValue && (!latestDate || dateValue >= latestDate)) {
      latest = candidate;
      latestDate = dateValue;
    } else if (!dateValue && !latestDate) {
      latest = candidate;
    }
  }

  return latest;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }
  const text = normalizeText(value);
  if (!text) {
    return null;
  }
  const clean = text.endsWith("Z") ? text.slice(0, -1) : text;
  const direct = new Date(clean);
  if (!Number.isNaN(direct.getTime())) {
    return direct;
  }
  const mdy = clean.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (mdy) {
    const month = Number(mdy[1]) - 1;
    const day = Number(mdy[2]);
    let year = Number(mdy[3]);
    if (year < 100) {
      year += 2000;
    }
    const parsed = new Date(year, month, day);
    if (!Number.isNaN(parsed.getTime())) {
      return parsed;
    }
  }
  return null;
}

function renderDateLabel(value) {
  const parsed = parseDateValue(value);
  if (parsed) {
    const yyyy = parsed.getFullYear();
    const mm = String(parsed.getMonth() + 1).padStart(2, "0");
    const dd = String(parsed.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }
  return normalizeText(value);
}

function cleanEntityLabel(value, maxChars = 18) {
  const cleaned = normalizeText(value).replace(/\s*\(.*?\)/g, "");
  if (cleaned.length <= maxChars) {
    return cleaned;
  }
  return `${cleaned.slice(0, maxChars - 1).trimEnd()}…`;
}

function treasuryLogoFallback(entity) {
  const source = cleanEntityLabel(entity, 64).replace(/[^a-zA-Z0-9\s]/g, " ");
  const rawWords = source
    .split(/\s+/)
    .map((word) => word.trim())
    .filter(Boolean);
  const stopWords = new Set([
    "the",
    "inc",
    "incorporated",
    "corp",
    "corporation",
    "company",
    "holdings",
    "group",
    "fund",
    "trust",
    "bitcoin",
    "etf",
    "limited",
    "ltd",
    "plc",
    "sa",
    "ag",
  ]);
  const words = rawWords.filter((word) => !stopWords.has(word.toLowerCase()));
  const pick = words.length ? words : rawWords;
  if (!pick.length) {
    return "BT";
  }
  if (pick.length === 1) {
    return pick[0].slice(0, 2).toUpperCase();
  }
  return `${pick[0][0] || ""}${pick[1][0] || ""}`.toUpperCase();
}

function readBtcPricePoints(rows, limit = 60) {
  if (!rows.length) {
    return [];
  }
  const mapping = headerIndexMap(rows, ["date", "price"]);
  const points = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const price = parseNumber(row[mapping.price], Number.NaN);
    if (!(price > 0)) {
      continue;
    }
    const dateRaw = row[mapping.date];
    points.push({
      idx: i,
      dateValue: parseDateValue(dateRaw),
      point: {
        date_label: renderDateLabel(dateRaw),
        price,
      },
    });
  }

  if (!points.length) {
    return [];
  }

  if (points.some((item) => item.dateValue !== null)) {
    points.sort((a, b) => {
      if (a.dateValue === null && b.dateValue === null) return a.idx - b.idx;
      if (a.dateValue === null) return 1;
      if (b.dateValue === null) return -1;
      return a.dateValue - b.dateValue;
    });
  }

  const byDateLabel = new Map();
  points.forEach((item) => {
    byDateLabel.set(item.point.date_label, item.point);
  });
  const ordered = Array.from(byDateLabel.values());
  return ordered.slice(-limit);
}

function readBtcPricePointsFromLivePrices(rows, limit = 60) {
  if (!rows.length) {
    return [];
  }
  const mapping = headerIndexMap(rows, ["date", "asset"]);
  const priceIdx = "price" in mapping ? mapping.price : ("close" in mapping ? mapping.close : -1);
  if (priceIdx < 0) {
    return [];
  }

  const points = [];
  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (!isBtcAsset(row[mapping.asset])) {
      continue;
    }
    const price = parseNumber(row[priceIdx], Number.NaN);
    if (!(price > 0)) {
      continue;
    }
    const dateRaw = row[mapping.date];
    points.push({
      idx: i,
      dateValue: parseDateValue(dateRaw),
      point: {
        date_label: renderDateLabel(dateRaw),
        price,
      },
    });
  }

  if (!points.length) {
    return [];
  }

  if (points.some((item) => item.dateValue !== null)) {
    points.sort((a, b) => {
      if (a.dateValue === null && b.dateValue === null) return a.idx - b.idx;
      if (a.dateValue === null) return 1;
      if (b.dateValue === null) return -1;
      return a.dateValue - b.dateValue;
    });
  }

  const byDateLabel = new Map();
  points.forEach((item) => {
    byDateLabel.set(item.point.date_label, item.point);
  });
  const ordered = Array.from(byDateLabel.values());
  return ordered.slice(-limit);
}

function readTreasuryBars(rows, limit = null) {
  if (!rows.length) {
    return [];
  }
  const mapping = headerIndexMap(rows, ["entity", "btc"]);
  const rowTypeIdx = "row_type" in mapping ? mapping.row_type : -1;
  const holderGroupIdx = "holder_group" in mapping ? mapping.holder_group : -1;
  const logoIdx =
    "logo" in mapping
      ? mapping.logo
      : ("logo_path" in mapping
          ? mapping.logo_path
          : ("icon" in mapping
              ? mapping.icon
              : ("image" in mapping
                  ? mapping.image
                  : ("png" in mapping
                      ? mapping.png
                      : ("file" in mapping
                          ? mapping.file
                          : ("filename" in mapping ? mapping.filename : -1))))));
  const showIdx = "show" in mapping ? mapping.show : -1;
  const bars = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (showIdx >= 0 && !isRowVisible(row[showIdx])) {
      continue;
    }
    const rowType = rowTypeIdx >= 0 ? normalizeText(row[rowTypeIdx]).toLowerCase() : "";
    if (rowType && rowType !== "entity") {
      continue;
    }
    const entity = normalizeText(row[mapping.entity]);
    const btc = parseNumber(row[mapping.btc], 0);
    if (!entity || btc <= 0) {
      continue;
    }
    bars.push({
      entity,
      btc,
      holder_group: holderGroupIdx >= 0 ? normalizeText(row[holderGroupIdx]) : "",
      logo: logoIdx >= 0 ? normalizeText(row[logoIdx]) : "",
      show: showIdx >= 0 ? normalizeText(row[showIdx]) : "yes",
    });
  }

  bars.sort((a, b) => b.btc - a.btc);
  if (Number.isFinite(limit) && limit > 0) {
    return bars.slice(0, limit);
  }
  return bars;
}

function readCirculatingMetric(rows) {
  if (!rows.length) {
    return null;
  }
  const mapping = headerIndexMap(rows, ["circulating_supply_btc", "max_supply_btc"]);
  const asOfIdx = "as_of_date" in mapping ? mapping.as_of_date : -1;
  const noteIdx = "note" in mapping ? mapping.note : -1;
  const showIdx = "show" in mapping ? mapping.show : -1;
  const candidates = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (showIdx >= 0 && !isRowVisible(row[showIdx])) {
      continue;
    }
    const circulating = parseNumber(row[mapping.circulating_supply_btc], 0);
    let maxSupply = parseNumber(row[mapping.max_supply_btc], 21_000_000);
    if (maxSupply <= 0) {
      maxSupply = 21_000_000;
    }
    if (circulating <= 0 && maxSupply <= 0) {
      continue;
    }

    const asOfRaw = asOfIdx >= 0 ? row[asOfIdx] : "";
    candidates.push({
      idx: i,
      dateValue: parseDateValue(asOfRaw),
      metric: {
        as_of_date: renderDateLabel(asOfRaw),
        circulating_supply_btc: circulating,
        max_supply_btc: maxSupply,
        note: noteIdx >= 0 ? normalizeText(row[noteIdx]) : "",
      },
    });
  }

  if (!candidates.length) {
    return null;
  }

  if (candidates.some((item) => item.dateValue !== null)) {
    candidates.sort((a, b) => {
      if (a.dateValue === null && b.dateValue === null) return a.idx - b.idx;
      if (a.dateValue === null) return 1;
      if (b.dateValue === null) return -1;
      return a.dateValue - b.dateValue;
    });
  }
  return candidates[candidates.length - 1].metric;
}

function readLiquidationBars(rows, limit = 6) {
  if (!rows.length) {
    return [];
  }
  const mapping = headerIndexMap(rows, ["label", "longs", "shorts"]);
  const periodTypeIdx = "period_type" in mapping ? mapping.period_type : -1;
  const periodKeyIdx = "period_key" in mapping ? mapping.period_key : -1;
  const totalIdx = "total" in mapping ? mapping.total : -1;
  const showIdx = "show" in mapping ? mapping.show : -1;

  const collect = (monthlyOnly) => {
    const bars = [];
    for (let i = 1; i < rows.length; i += 1) {
      const row = rows[i];
      if (showIdx >= 0 && !isRowVisible(row[showIdx])) {
        continue;
      }
      const periodType = periodTypeIdx >= 0 ? normalizeText(row[periodTypeIdx]).toLowerCase() : "";
      if (monthlyOnly && periodType && periodType !== "monthly") {
        continue;
      }
      const label = normalizeText(row[mapping.label]) || (periodKeyIdx >= 0 ? normalizeText(row[periodKeyIdx]) : "");
      const longs = Math.max(0, parseNumber(row[mapping.longs], 0));
      const shorts = Math.max(0, parseNumber(row[mapping.shorts], 0));
      let total = totalIdx >= 0 ? parseNumber(row[totalIdx], longs + shorts) : longs + shorts;
      if (total <= 0) {
        total = longs + shorts;
      }
      if (!label || total <= 0) {
        continue;
      }
      bars.push({
        label,
        longs,
        shorts,
        total,
        show: showIdx >= 0 ? normalizeText(row[showIdx]) : "yes",
      });
    }
    return bars;
  };

  const bars = collect(true);
  const finalBars = bars.length ? bars : collect(false);
  return finalBars.slice(-limit);
}

function readOwnershipSegments(rows) {
  if (!rows.length) {
    return [];
  }
  const mapping = headerIndexMap(rows, ["category", "amount_btc"]);
  const colorIdx = "color" in mapping ? mapping.color : -1;
  const percentIdx = "percent" in mapping ? mapping.percent : -1;
  const showIdx = "show" in mapping ? mapping.show : -1;
  const segments = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    if (showIdx >= 0 && !isRowVisible(row[showIdx])) {
      continue;
    }
    const category = normalizeText(row[mapping.category]);
    const amount = Math.max(0, parseNumber(row[mapping.amount_btc], 0));
    const percent = percentIdx >= 0 ? Math.max(0, parseNumber(row[percentIdx], 0)) : 0;
    const color = colorIdx >= 0 ? safeColor(row[colorIdx]) : "";
    if (!category && amount <= 0) {
      continue;
    }
    if (!category) {
      continue;
    }
    segments.push({
      category,
      amount_btc: amount,
      percent,
      color: color || "rgb(255, 66, 2)",
      show: showIdx >= 0 ? normalizeText(row[showIdx]) : "yes",
    });
  }

  if (!segments.length) {
    return [];
  }

  const totalPercent = segments.reduce((acc, item) => acc + Math.max(0, item.percent), 0);
  if (totalPercent > 0) {
    segments.forEach((item) => {
      item.percent = (Math.max(0, item.percent) / totalPercent) * 100;
    });
  } else {
    const totalAmount = segments.reduce((acc, item) => acc + Math.max(0, item.amount_btc), 0);
    if (totalAmount > 0) {
      segments.forEach((item) => {
        item.percent = (Math.max(0, item.amount_btc) / totalAmount) * 100;
      });
    }
  }

  segments.sort((a, b) => b.amount_btc - a.amount_btc);
  return segments;
}

function renderHtml(
  meta,
  points,
  btcPricePoints,
  treasuryBars,
  circulatingMetric,
  liquidationBars,
  ownershipSegments,
  graphSettingsMap,
  liveBtc,
  imageOptions
) {
  const title = escapeHtml(meta.main_title);
  const subtitle = escapeHtml(meta.subtitle);
  const eyebrow = escapeHtml(meta.eyebrow);
  const blockHeight = escapeHtml(renderBlockHeight(meta.block_height));
  const tldrTitle = escapeHtml(meta.tldr_title);
  const tldrContent = indentBlock(renderContentBlocks(meta.tldr_content), 16);
  const conclusionTitle = escapeHtml(meta.conclusion_title);
  const conclusionContent = indentBlock(renderContentBlocks(meta.conclusion_content), 16);
  const addressLine = escapeHtml(meta.address_line);
  const footerLine = escapeHtml(meta.footer_line);
  const heroImageUrl = escapeHtml(resolveHeroImage(meta), true);
  const resolvedLogo = resolveLogo(meta);
  const logoUrl = escapeHtml(resolvedLogo, true);
  const footerLogoUrl = escapeHtml(resolveAssetPath(meta.footer_logo_url, "/logotosite.png"), true);
  const footerInstagramIcon = "/instagram.png";
  const footerXIcon = "/x:twitter.png";
  const footerLinkedinIcon = "/linkedin.png";

  const pointsHtml = points
    .map((point) => renderPoint(point, meta, imageOptions))
    .join("");
  const marketHtml = renderMarketSection(
    meta,
    treasuryBars,
    circulatingMetric,
    liquidationBars,
    ownershipSegments,
    graphSettingsMap,
    liveBtc
  );

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${title} - Globalite Macro Brief</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;700&display=swap" rel="stylesheet">
    <style>
      body {
        margin: 0;
        padding: 0;
        background: #f5f5f5;
        font-family: "Poppins", Arial, sans-serif;
        color: #1f1f1f;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      table { border-collapse: collapse; }
      img { border: 0; display: block; max-width: 100%; height: auto; }
      a { color: #ff4202; text-decoration: none; }
      .toolbar { width: 100%; max-width: 650px; margin: 0 auto; display: flex; justify-content: flex-end; padding: 12px 0 8px; }
      .download-pdf-btn { border: 1px solid #ff4202; border-radius: 999px; padding: 8px 14px; background: #ffffff; color: #ff4202; font: 600 12px/1 "Poppins", Arial, sans-serif; cursor: pointer; }
      .download-pdf-btn:hover { background: #fff4ef; }
      .wrapper { width: 100%; background: #f5f5f5; padding: 32px 0; }
      .container { width: 650px; max-width: 650px; background: #ffffff; border: 1px solid #e6e6e6; border-radius: 16px; overflow: hidden; }
      .divider { height: 4px; background: #ff4202; line-height: 4px; }
      .hero { position: relative; overflow: hidden; background: #0a0a0a; min-height: 260px; display: flex; flex-direction: column; justify-content: flex-end; }
      .hero-bg { position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; object-position: center top; opacity: 0.82; }
      .hero-gradient { position: absolute; inset: 0; background: linear-gradient(to bottom, rgba(0,0,0,0.1) 0%, rgba(0,0,0,0.72) 100%); }
      .hero-content { position: relative; z-index: 2; padding: 28px 32px; }
      .hero-logo img { width: 150px; height: auto; display: block; margin: 0 0 14px; }
      .hero-eyebrow { color: #ff4202; font-weight: 700; font-size: 11px; letter-spacing: 1.4px; text-transform: uppercase; margin: 0 0 6px; }
      .hero-title { margin: 0 0 6px; font-size: 28px; font-weight: 700; color: #ffffff; }
      .hero-subtitle { margin: 0 0 14px; color: rgba(255,255,255,0.6); font-size: 14px; }
      .hero-badge { display: inline-block; font-size: 11px; color: #ffcfb8; background: rgba(255,66,2,0.25); border: 1px solid rgba(255,66,2,0.45); border-radius: 999px; padding: 5px 12px; }
      .section { padding: 16px 32px; border-top: 1px solid #f0f0f0; }
      .section h2 { margin: 0 0 20px; font-size: 18px; font-weight: 700; }
      .section p { margin: 0; font-size: 14px; line-height: 1.6; }
      .section p + p { margin-top: 12px; }
      .section ul { margin: 20px 0 20px 18px; padding: 0; font-size: 14px; line-height: 1.6; }
      .section li { margin-bottom: 8px; }
      .section .point-source { margin-top: 14px; font-size: 11px; line-height: 1.5; color: #8a8a8a; }
      .image { margin: 20px 0; }
      .image img { width: 100%; border-radius: 12px; border: 1px solid #e6e6e6; }
      .caption { font-size: 12px; color: #7a7a7a; margin-top: 6px; }
      .extra-images { margin: 14px 0 0; display: grid; gap: 10px; }
      .extra-images img { width: 100%; border-radius: 12px; border: 1px solid #e6e6e6; }
      .market { background: #0f0f0f; color: #f4f4f4; border-top: 1px solid #171717; }
      .market-live { margin: 0 0 12px; display: inline-flex; gap: 6px; align-items: baseline; font-size: 12px; color: #ffcfb8; background: rgba(255,66,2,0.16); border: 1px solid rgba(255,66,2,0.35); border-radius: 999px; padding: 4px 10px; }
      .market-live strong { color: #ffffff; }
      .market-live span { color: #ffcfb8; }
      /* Snapshot / bottom graphs */
      .snapshot-section { width:100%; margin-top:28px; padding-top:18px; border-top:1px solid #242424; }
      .snapshot-title { font-size:1.08em; font-weight:500; color:#fff; margin-bottom:14px; letter-spacing:.01em; }
      .snapshot-grid { display:grid; gap:14px; grid-template-columns:1fr; }
      .snapshot-card { border:1px solid #232323; border-radius:12px; background:#141414; padding:14px; }
      .snapshot-card h3 { font-size:.95em; font-weight:500; color:#fff; margin-bottom:10px; }
      .snapshot-caption { margin-top:8px; color:#9a9a9a; font-size:.84em; }
      .snapshot-empty { margin: 0; color: #9a9a9a; font-size: .84em; }
      .snapshot-distribution-bar { width:100%; height:26px; display:flex; border-radius:8px; overflow:hidden; border:1px solid #2a2a2a; background:#1a1a1a; }
      .snapshot-distribution-segment { height:100%; min-width:2px; border-right:1px solid rgba(255,255,255,.15); }
      .snapshot-distribution-segment:last-child { border-right:none; }
      .snapshot-legend { margin-top:10px; display:grid; grid-template-columns:1fr; gap:6px; }
      .snapshot-legend-item { display:flex; align-items:center; justify-content:space-between; gap:10px; font-size:.84em; color:#e6e6e6; }
      .snapshot-legend-left { display:inline-flex; align-items:center; gap:8px; }
      .snapshot-legend-dot { width:10px; height:10px; border-radius:999px; flex:0 0 10px; }
      .snapshot-legend-name { overflow:hidden; white-space:nowrap; text-overflow:ellipsis; color:#e6e6e6; }
      .snapshot-legend-value { color:#9a9a9a; white-space:nowrap; font-variant-numeric:tabular-nums; }
      .snapshot-circ-value { font-size:1.55em; color:#fff; font-weight:500; margin-bottom:10px; font-variant-numeric:tabular-nums; }
      .snapshot-circ-bar { width:100%; height:18px; border-radius:999px; background:#1a1a1a; overflow:hidden; border:1px solid #2a2a2a; }
      .snapshot-circ-fill { height:100%; background:linear-gradient(90deg,#ff4202 0%,#ff8f60 100%); width:0%; }
      .snapshot-circ-note { margin-top:8px; color:#9a9a9a; font-size:.84em; }
      .liq-bar-wrap { display:flex; flex-direction:column; gap:6px; margin-top:8px; }
      .liq-row { display:grid; grid-template-columns:80px 1fr 1fr 60px; gap:8px; align-items:center; font-size:.82em; color:#e6e6e6; }
      .liq-label { color:#9a9a9a; white-space:nowrap; }
      .liq-bar-track { background:#1f1f1f; border-radius:4px; height:10px; overflow:hidden; position:relative; }
      .liq-bar-longs { height:100%; background:#ff4202; border-radius:4px; }
      .liq-bar-shorts { height:100%; background:#6699ff; border-radius:4px; }
      .liq-total { text-align:right; color:#9a9a9a; }
      .liq-legend { display:flex; gap:14px; margin-bottom:8px; font-size:.82em; }
      .liq-dot { width:8px; height:8px; border-radius:999px; display:inline-block; margin-right:4px; }
      .snapshot-treas-bars { display:flex; gap:12px; align-items:flex-start; overflow-x:auto; padding:2px 2px 8px; width:100%; max-width:100%; box-sizing:border-box; }
      .snapshot-treas-item { display:flex; flex-direction:column; align-items:center; justify-content:flex-start; gap:6px; flex:1 0 100px; min-width:100px; max-width:140px; }
      .snapshot-treas-track { width:100%; height:160px; border-radius:8px; border:1px solid #2a2a2a; background:#1a1a1a; overflow:hidden; display:flex; align-items:flex-end; }
      .snapshot-treas-fill { width:100%; background:linear-gradient(180deg,#ff8f60 0%,#ff4202 100%); }
      .snapshot-treas-logo-wrap { width:26px; height:26px; border-radius:999px; background:#0f0f0f; border:1px solid #2a2a2a; display:flex; align-items:center; justify-content:center; overflow:hidden; }
      .snapshot-treas-logo { width:20px; height:20px; object-fit:contain; }
      .snapshot-treas-logo-fallback { width:20px; height:20px; display:flex; align-items:center; justify-content:center; font-size:10px; font-weight:700; color:#ffcfb8; letter-spacing:.3px; text-transform:uppercase; }
      .snapshot-treas-label { font-size:.78em; color:#e6e6e6; text-align:center; line-height:1.25; min-height:32px; max-height:32px; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; }
      .snapshot-treas-value { color:#ff8f60; white-space:nowrap; font-variant-numeric:tabular-nums; font-size:.8em; min-height:18px; }
      .snapshot-treas-group { color:#888; font-size:.7em; text-align:center; min-height:28px; max-height:28px; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; }
      .snapshot-treas-table-wrap { margin-top:12px; border:1px solid #2a2a2a; border-radius:10px; overflow:hidden; }
      .snapshot-treas-table { width:100%; border-collapse:collapse; font-size:.78em; }
      .snapshot-treas-table th { text-align:left; padding:8px 10px; color:#9f9f9f; font-weight:500; border-bottom:1px solid #2a2a2a; background:#121212; white-space:nowrap; }
      .snapshot-treas-table td { padding:7px 10px; border-bottom:1px solid #202020; color:#e6e6e6; vertical-align:middle; }
      .snapshot-treas-table tr:last-child td { border-bottom:none; }
      .snapshot-treas-cell-name { max-width:220px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
      .snapshot-treas-cell-num { text-align:right; color:#ffcfb8; white-space:nowrap; font-variant-numeric:tabular-nums; }
      .tldr { background: #fff8ec; border-top: 2px solid #ff4202; }
      .conclusion { background: #fff7f3; border-top: 2px solid #ff4202; }
      .footer { padding: 0; font-size: 12px; color: #7a7a7a; }
      .footer-legal { padding: 16px 32px 12px; }
      .footer-dark { background: #0f0f0f; padding: 22px 32px; }
      .footer-bar { display: flex; align-items: center; justify-content: space-between; }
      .footer-site-link { display: inline-flex; align-items: center; gap: 10px; text-decoration: none; }
      .footer-site-link img { width: 36px; height: 36px; border-radius: 10px; object-fit: contain; }
      .footer-site-name { display: block; color: #ffffff; font-size: 13px; font-weight: 700; }
      .footer-site-url { display: block; color: #ff4202; font-size: 11px; }
      .footer-socials { display: flex; gap: 8px; }
      .footer-social-btn { display: inline-flex; align-items: center; justify-content: center; width: 36px; height: 36px; background: #ffffff; border-radius: 10px; border: 1px solid #efefef; }
      .footer-social-btn img { width: 20px; height: 20px; object-fit: contain; }
      .footer-copy { margin: 14px 0 0; padding-top: 14px; border-top: 1px solid rgba(255,255,255,0.08); font-size: 11px; color: rgba(255,255,255,0.3); text-align: center; }
      @media (max-width: 720px) {
        .toolbar { padding: 10px 16px 6px; box-sizing: border-box; }
        .wrapper { padding: 16px 0; }
        .container { width: 100%; max-width: 100%; border-radius: 0; }
        .section { padding: 18px 20px; }
        .hero-content { padding: 20px; }
        .hero-title { font-size: 22px; }
        .footer-legal { padding: 14px 20px 10px; }
        .footer-dark { padding: 18px 20px; }
        .footer-bar { flex-direction: column; align-items: flex-start; gap: 14px; }
      }
      @media print {
        .no-print { display: none !important; }
        body { background: #ffffff; }
        .wrapper { background: #ffffff; padding: 0; }
        .container { border: 0; border-radius: 0; }
      }
    </style>
  </head>
  <body>
    <div class="toolbar no-print">
      <button class="download-pdf-btn" type="button" onclick="window.print()">Download PDF</button>
    </div>
    <table class="wrapper" role="presentation" width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center">
          <table class="container" role="presentation" cellpadding="0" cellspacing="0">
            <tr>
              <td class="divider">&nbsp;</td>
            </tr>
            <tr>
              <td class="hero">
                <img class="hero-bg" src="${heroImageUrl}" alt="">
                <div class="hero-gradient"></div>
                <div class="hero-content">
                  <div class="hero-logo"><img src="${logoUrl}" alt="Globalite"></div>
                  <p class="hero-eyebrow">${eyebrow}</p>
                  <h1 class="hero-title">${title}</h1>
                  <p class="hero-subtitle">${subtitle}</p>
                  <span class="hero-badge">Written at block height: <strong>${blockHeight}</strong></span>
                </div>
              </td>
            </tr>
${pointsHtml}
            <tr>
              <td class="section tldr">
                <h2>${tldrTitle}</h2>
${tldrContent}
              </td>
            </tr>
            <tr>
              <td class="section conclusion">
                <h2>${conclusionTitle}</h2>
${conclusionContent}
              </td>
            </tr>
${marketHtml}
            <tr>
              <td class="footer">
                <div class="footer-legal">
                  <p>${footerLine}</p>
                  <p>${addressLine}</p>
                </div>
                <div class="footer-dark">
                  <div class="footer-bar">
                    <a class="footer-site-link" href="https://globalite.co" target="_blank">
                      <img src="${footerLogoUrl}" alt="Globalite">
                      <span>
                        <span class="footer-site-name">Globalite</span>
                        <span class="footer-site-url">globalite.co</span>
                      </span>
                    </a>
                    <div class="footer-socials">
                      <a class="footer-social-btn" href="https://www.instagram.com/globalite.sa/"><img src="${footerInstagramIcon}" alt="Instagram"></a>
                      <a class="footer-social-btn" href="https://x.com/globalite_sa"><img src="${footerXIcon}" alt="X"></a>
                      <a class="footer-social-btn" href="https://www.linkedin.com/company/globalite-sa"><img src="${footerLinkedinIcon}" alt="LinkedIn"></a>
                    </div>
                  </div>
                  <p class="footer-copy">© 2026 Globalite SA. All rights reserved.</p>
                </div>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    <script>
      (function () {
        var params = new URLSearchParams(window.location.search);
        var refreshSeconds = Number(params.get("refresh"));
        if (!Number.isFinite(refreshSeconds) || refreshSeconds < 5) {
          return;
        }
        window.setInterval(function () {
          window.location.reload();
        }, refreshSeconds * 1000);
      })();
    </script>
  </body>
</html>`;
}

function renderPoint(point, meta, imageOptions) {
  const imageSources = resolveImageSources(point, meta, imageOptions);
  const imageBlock = renderImageBlock(point, imageSources);
  const extraImageSources = resolveExtraImagePaths(point, meta, imageOptions);
  const extraImagesBlock = renderExtraImagesBlock(point, extraImageSources);

  let output = "            <tr>\n";
  output += "              <td class=\"section\">\n";
  output += `                <h2>${point.order}. ${escapeHtml(point.title)}</h2>\n`;
  if (imageBlock) {
    output += `${indentBlock(imageBlock, 16)}\n`;
  }
  output += `${indentBlock(renderContentBlocks(point.content), 16)}\n`;
  if (point.source) {
    output += `                <p class=\"point-source\">${escapeHtml(point.source)}</p>\n`;
  }
  if (extraImagesBlock) {
    output += `${indentBlock(extraImagesBlock, 16)}\n`;
  }
  output += "              </td>\n";
  output += "            </tr>\n";
  return output;
}

function renderImageBlock(point, imageSources) {
  const primarySrc = imageSources[0] || "";
  if (!primarySrc) {
    return "";
  }
  const fallbacks = imageSources.slice(1).join("|");
  const caption = point.image_caption || point.title;
  return (
    '<div class="image">\n' +
    `  <img src="${escapeHtml(primarySrc, true)}" data-fallbacks="${escapeHtml(fallbacks, true)}" alt="${escapeHtml(point.title)}" onerror="const list=(this.dataset.fallbacks||'').split('|').filter(Boolean);if(list.length){this.src=list.shift();this.dataset.fallbacks=list.join('|');}else{this.closest('.image').style.display='none';}">\n` +
    `  <div class="caption">${escapeHtml(caption)}</div>\n` +
    "</div>"
  );
}

function renderExtraImagesBlock(point, imageSources) {
  if (!imageSources.length) {
    return "";
  }
  const imageTags = imageSources
    .map(
      (src, index) =>
        `  <img src="${escapeHtml(src, true)}" alt="${escapeHtml(point.title)} - extra ${index + 1}" onerror="this.style.display='none'">`
    )
    .join("\n");
  return '<div class="extra-images">\n' + imageTags + "\n</div>";
}

function formatUsd(value) {
  const absValue = Math.abs(value);
  if (absValue >= 1000) {
    return `$${Math.round(value).toLocaleString("en-US")}`;
  }
  return `$${Number(value).toFixed(2)}`.replace(/\.00$/, "");
}

function formatUsdLarge(value) {
  const abs = Math.abs(value);
  const trim = (num) => num.toFixed(2).replace(/\.0+$/, "").replace(/(\.\d*?)0+$/, "$1");
  if (abs >= 1e12) {
    return `$${trim(value / 1e12)}T`;
  }
  if (abs >= 1e9) {
    return `$${trim(value / 1e9)}B`;
  }
  if (abs >= 1e6) {
    return `$${trim(value / 1e6)}M`;
  }
  if (abs >= 1e3) {
    return `$${trim(value / 1e3)}K`;
  }
  return formatUsd(value);
}

function buildBtcPriceChartSvg(points) {
  if (!points.length) {
    return '<p class="market-empty">No BTC price points found.</p>';
  }

  const width = 620;
  const height = 230;
  const padLeft = 54;
  const padRight = 14;
  const padTop = 14;
  const padBottom = 34;
  const plotWidth = width - padLeft - padRight;
  const plotHeight = height - padTop - padBottom;
  const plotBottom = padTop + plotHeight;

  let minPrice = Math.min(...points.map((point) => point.price));
  let maxPrice = Math.max(...points.map((point) => point.price));
  if (maxPrice <= minPrice) {
    maxPrice = minPrice + 1;
  }

  const padding = Math.max((maxPrice - minPrice) * 0.08, maxPrice * 0.01);
  minPrice = Math.max(0, minPrice - padding);
  maxPrice += padding;
  const span = maxPrice - minPrice;
  const denominator = Math.max(1, points.length - 1);

  const coords = points.map((point, idx) => {
    const x = padLeft + (plotWidth * idx) / denominator;
    const y = padTop + ((maxPrice - point.price) / span) * plotHeight;
    return { x, y };
  });

  const linePath = `M ${coords.map((item) => `${item.x.toFixed(2)} ${item.y.toFixed(2)}`).join(" L ")}`;
  const areaPath =
    `M ${coords[0].x.toFixed(2)} ${plotBottom.toFixed(2)} ` +
    `${coords.map((item) => `L ${item.x.toFixed(2)} ${item.y.toFixed(2)}`).join(" ")} ` +
    `L ${coords[coords.length - 1].x.toFixed(2)} ${plotBottom.toFixed(2)} Z`;

  const gridLines = [];
  const yLabels = [];
  for (let step = 0; step < 5; step += 1) {
    const y = padTop + (plotHeight * step) / 4;
    const value = maxPrice - ((maxPrice - minPrice) * step) / 4;
    gridLines.push(
      `<line x1="${padLeft.toFixed(2)}" y1="${y.toFixed(2)}" x2="${(width - padRight).toFixed(2)}" y2="${y.toFixed(2)}" stroke="#262626" stroke-width="1" />`
    );
    yLabels.push(
      `<text x="${(padLeft - 8).toFixed(2)}" y="${(y + 4).toFixed(2)}" text-anchor="end" fill="#8b8b8b" font-size="10">${escapeHtml(formatUsd(value))}</text>`
    );
  }

  const firstLabel = escapeHtml(points[0].date_label);
  const midLabel = escapeHtml(points[Math.floor(points.length / 2)].date_label);
  const lastLabel = escapeHtml(points[points.length - 1].date_label);
  const xLabels =
    `<text x="${padLeft.toFixed(2)}" y="${(height - 8).toFixed(2)}" text-anchor="start" fill="#8b8b8b" font-size="10">${firstLabel}</text>` +
    `<text x="${(padLeft + plotWidth / 2).toFixed(2)}" y="${(height - 8).toFixed(2)}" text-anchor="middle" fill="#8b8b8b" font-size="10">${midLabel}</text>` +
    `<text x="${(width - padRight).toFixed(2)}" y="${(height - 8).toFixed(2)}" text-anchor="end" fill="#8b8b8b" font-size="10">${lastLabel}</text>`;

  const last = coords[coords.length - 1];
  return `
<svg class="market-price-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Bitcoin price trend chart">
  <defs>
    <linearGradient id="priceAreaGradient" x1="0" x2="0" y1="0" y2="1">
      <stop offset="0%" stop-color="#ff4202" stop-opacity="0.42" />
      <stop offset="100%" stop-color="#ff4202" stop-opacity="0" />
    </linearGradient>
  </defs>
  ${gridLines.join("")}
  ${yLabels.join("")}
  <path d="${areaPath}" fill="url(#priceAreaGradient)" />
  <path d="${linePath}" fill="none" stroke="#ff4202" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round" />
  <circle cx="${last.x.toFixed(2)}" cy="${last.y.toFixed(2)}" r="4" fill="#ff4202" stroke="#ffffff" stroke-width="1" />
  ${xLabels}
</svg>`;
}

function renderTreasuryBars(treasuryBars) {
  if (!treasuryBars.length) {
    return '<p class="market-empty">No treasury rows found.</p>';
  }
  let maxBtc = Math.max(...treasuryBars.map((item) => item.btc));
  if (!(maxBtc > 0)) {
    maxBtc = 1;
  }
  const barsHtml = treasuryBars
    .map((item) => {
      const height = Math.max(8, (item.btc / maxBtc) * 100);
      return (
        '<div class="market-bar-item">' +
        `<div class="market-bar-track"><div class="market-bar-fill" style="height:${height.toFixed(6)}%;"></div></div>` +
        `<p class="market-bar-label">${escapeHtml(cleanEntityLabel(item.entity))}</p>` +
        `<p class="market-bar-value">${escapeHtml(formatBtcCompact(item.btc))}</p>` +
        "</div>"
      );
    })
    .join("");
  return `<div class="market-bars">${barsHtml}</div>`;
}

function renderLiquidationsBars(liquidationBars) {
  if (!liquidationBars.length) {
    return '<p class="market-empty">No liquidation rows found.</p>';
  }
  let maxTotal = Math.max(...liquidationBars.map((item) => item.total));
  if (!(maxTotal > 0)) {
    maxTotal = 1;
  }
  const barsHtml = liquidationBars
    .map((item) => {
      const longHeight = item.longs > 0 ? Math.max(4, (item.longs / maxTotal) * 100) : 0;
      const shortHeight = item.shorts > 0 ? Math.max(4, (item.shorts / maxTotal) * 100) : 0;
      const totalLabel = `$${Number(item.total).toFixed(2).replace(/\\.00$/, "")}B`;
      return (
        '<div class="market-bar-item">' +
        '<div class="market-bar-track market-liq-track">' +
        `<div class="market-liq-short" style="height:${shortHeight.toFixed(6)}%;"></div>` +
        `<div class="market-liq-long" style="height:${longHeight.toFixed(6)}%;"></div>` +
        "</div>" +
        `<p class="market-bar-label">${escapeHtml(item.label)}</p>` +
        `<p class="market-bar-value">${escapeHtml(totalLabel)}</p>` +
        "</div>"
      );
    })
    .join("");

  return (
    '<div class="market-legend">' +
    '<span><i class="dot dot-long"></i>Longs</span>' +
    '<span><i class="dot dot-short"></i>Shorts</span>' +
    "</div>" +
    `<div class="market-bars">${barsHtml}</div>`
  );
}

function renderCirculatingCard(circulatingMetric) {
  if (!circulatingMetric) {
    return '<p class="market-empty">No circulating supply row found.</p>';
  }
  const maxSupply =
    circulatingMetric.max_supply_btc > 0 ? circulatingMetric.max_supply_btc : 21_000_000;
  const pct = Math.max(
    0,
    Math.min(
      100,
      maxSupply > 0 ? (circulatingMetric.circulating_supply_btc / maxSupply) * 100 : 0
    )
  );
  const asOfHtml = circulatingMetric.as_of_date
    ? `<p class="market-subnote">As of ${escapeHtml(circulatingMetric.as_of_date)}</p>`
    : "";
  const noteHtml = circulatingMetric.note
    ? `<p class="market-note">${escapeHtml(circulatingMetric.note)}</p>`
    : "";
  return `
<div class="market-circ-value">${escapeHtml(formatBtcInteger(circulatingMetric.circulating_supply_btc))} BTC</div>
<div class="market-progress-track"><div class="market-progress-fill" style="width:${pct.toFixed(6)}%;"></div></div>
<p class="market-subnote">${escapeHtml(formatPercent(pct))} of ${escapeHtml(formatBtcInteger(maxSupply))} BTC max supply</p>
${asOfHtml}
${noteHtml}`;
}

function renderOwnershipCard(ownershipSegments, maxItems = 8) {
  if (!ownershipSegments.length) {
    return '<p class="market-empty">No ownership rows found.</p>';
  }

  const barSegments = ownershipSegments
    .map(
      (item) =>
        `<div class="market-own-segment" style="width:${Math.max(0, item.percent).toFixed(6)}%;background:${escapeHtml(item.color, true)};" title="${escapeHtml(item.category, true)}: ${escapeHtml(formatBtcCompact(item.amount_btc), true)} (${escapeHtml(formatPercent(item.percent), true)})"></div>`
    )
    .join("");

  const legendRows = ownershipSegments
    .slice(0, maxItems)
    .map(
      (item) =>
        '<div class="market-own-item">' +
        `<span class="market-own-dot" style="background:${escapeHtml(item.color, true)};"></span>` +
        `<span class="market-own-name">${escapeHtml(item.category)}</span>` +
        `<span class="market-own-value">${escapeHtml(formatPercent(item.percent))}</span>` +
        "</div>"
    )
    .join("");

  return (
    `<div class="market-own-bar">${barSegments}</div>` +
    `<div class="market-own-legend">${legendRows}</div>`
  );
}

function renderMarketSection(
  meta,
  treasuryBars,
  circulatingMetric,
  liquidationBars,
  ownershipSegments,
  graphSettingsMap,
  liveBtc
) {
  const snapshotHtml = renderSnapshotSection({
    distribution: ownershipSegments,
    circulating: circulatingMetric,
    liquidations: liquidationBars,
    treasuries: treasuryBars,
    live_btc_price: liveBtc && Number.isFinite(Number(liveBtc.price)) ? Number(liveBtc.price) : 0,
    settings: graphSettingsMap || {},
    title: normalizeText(meta.snapshot_title) || "Bitcoin Data",
  });
  if (!snapshotHtml) {
    return "";
  }
  const liveChip = liveBtc
    ? `<p class="market-live">Live BTC: <strong>${escapeHtml(formatUsd(liveBtc.price))}</strong>${
        liveBtc.date ? ` <span>(${escapeHtml(liveBtc.date)})</span>` : ""
      }</p>`
    : "";
  return `
            <tr>
              <td class="section market">
                ${liveChip}
                ${snapshotHtml}
              </td>
            </tr>
`;
}

function renderSnapshotSection(data) {
  const cards = [];
  const settings = data.settings || {};

  const ownershipSetting = {
    show: !settings.ownership || settings.ownership.show !== false,
    title:
      (settings.ownership && normalizeText(settings.ownership.title)) || "Supply Ownership",
    top_n: toPositiveInt(settings.ownership && settings.ownership.top_n, 8),
  };
  const circulatingSetting = {
    show: !settings.circulating_btc || settings.circulating_btc.show !== false,
    title:
      (settings.circulating_btc && normalizeText(settings.circulating_btc.title)) ||
      "Circulating BTC",
  };
  const liquidationsSetting = {
    show: !settings.liquidations || settings.liquidations.show !== false,
    title:
      (settings.liquidations && normalizeText(settings.liquidations.title)) || "Liquidations",
    top_n: toPositiveInt(settings.liquidations && settings.liquidations.top_n, 6),
  };
  const treasuriesSetting = {
    show: !settings.treasuries || settings.treasuries.show !== false,
    title: (settings.treasuries && normalizeText(settings.treasuries.title)) || "Treasuries",
    max_bars: toPositiveInt(settings.treasuries && settings.treasuries.max_bars, 10),
  };

  if (ownershipSetting.show && data.distribution && data.distribution.length) {
    const segments = data.distribution
      .filter((row) => isRowVisible(row.show))
      .slice(0, ownershipSetting.top_n || 8);
    const total = segments.reduce((sum, row) => sum + Number(row.amount_btc || 0), 0);
    const bar = segments
      .map((row) => {
        const pct =
          row.percent && row.percent > 0
            ? Number(row.percent)
            : total > 0
              ? (Number(row.amount_btc || 0) / total) * 100
              : 0;
        return `<div class="snapshot-distribution-segment" style="width:${pct.toFixed(4)}%;background:${escapeHtml(safeColor(row.color), true)}" title="${escapeHtml(`${row.category || ""}: ${pct.toFixed(1)}%`, true)}"></div>`;
      })
      .join("");
    const legend = segments
      .map((row) => {
        const pct =
          row.percent && row.percent > 0
            ? Number(row.percent)
            : total > 0
              ? (Number(row.amount_btc || 0) / total) * 100
              : 0;
        const btc = Number(row.amount_btc || 0);
        const btcFmt = btc >= 1e6 ? `${(btc / 1e6).toFixed(2)}M` : btc >= 1e3 ? `${(btc / 1e3).toFixed(0)}K` : btc.toFixed(0);
        return `<div class="snapshot-legend-item">
        <span class="snapshot-legend-left">
          <span class="snapshot-legend-dot" style="background:${escapeHtml(safeColor(row.color), true)}"></span>
          <span class="snapshot-legend-name">${escapeHtml(row.category || "")}</span>
        </span>
        <span class="snapshot-legend-value">${escapeHtml(btcFmt)} BTC (${escapeHtml(pct.toFixed(1))}%)</span>
      </div>`;
      })
      .join("");
    cards.push(`<article class="snapshot-card">
      <h3>${escapeHtml(ownershipSetting.title)}</h3>
      <div class="snapshot-distribution-bar">${bar}</div>
      <div class="snapshot-legend">${legend}</div>
      <p class="snapshot-caption">Largest holders shown first from left to right.</p>
    </article>`);
  }

  if (circulatingSetting.show && data.circulating) {
    const c = data.circulating;
    const circ = Number(c.circulating_supply_btc || 0);
    const max = Number(c.max_supply_btc || 21000000);
    const pct = max > 0 ? (circ / max) * 100 : 0;
    const circFmt = Math.round(circ).toLocaleString("en-US");
    cards.push(`<article class="snapshot-card">
      <h3>${escapeHtml(circulatingSetting.title)}</h3>
      <div class="snapshot-circ-value">${escapeHtml(circFmt)} BTC</div>
      <div class="snapshot-circ-bar"><div class="snapshot-circ-fill" style="width:${pct.toFixed(2)}%"></div></div>
      <p class="snapshot-circ-note">${escapeHtml(pct.toFixed(2))}% of the 21,000,000 BTC maximum supply has been mined.</p>
    </article>`);
  }

  if (liquidationsSetting.show && data.liquidations && data.liquidations.length) {
    const rows = data.liquidations
      .filter((row) => isRowVisible(row.show))
      .slice(-(liquidationsSetting.top_n || 6));
    const maxTotal = Math.max(
      ...rows.map((row) => Number(row.total || row.longs || 0) + Number(row.shorts || 0)),
      0
    );
    const rowsHtml = rows
      .map((row) => {
        const longs = Number(row.longs || 0);
        const shorts = Number(row.shorts || 0);
        const total = Number(row.total || 0) || longs + shorts;
        const lPct = maxTotal > 0 ? (longs / maxTotal) * 100 : 0;
        const sPct = maxTotal > 0 ? (shorts / maxTotal) * 100 : 0;
        const totalLabel = total >= 1 ? `${total.toFixed(1)}B` : `${(total * 1000).toFixed(0)}M`;
        return `<div class="liq-row">
        <span class="liq-label">${escapeHtml(row.label || "")}</span>
        <div class="liq-bar-track"><div class="liq-bar-longs" style="width:${lPct.toFixed(1)}%"></div></div>
        <div class="liq-bar-track"><div class="liq-bar-shorts" style="width:${sPct.toFixed(1)}%"></div></div>
        <span class="liq-total">${escapeHtml(totalLabel)}</span>
      </div>`;
      })
      .join("");
    cards.push(`<article class="snapshot-card">
      <h3>${escapeHtml(liquidationsSetting.title)}</h3>
      <div class="liq-legend">
        <span><span class="liq-dot" style="background:#ff4202"></span>Longs</span>
        <span><span class="liq-dot" style="background:#6699ff"></span>Shorts</span>
      </div>
      <div class="liq-bar-wrap">
        <div class="liq-row"><span class="liq-label"></span><span style="font-size:.75em;color:#555">LONGS</span><span style="font-size:.75em;color:#555">SHORTS</span><span></span></div>
        ${rowsHtml}
      </div>
    </article>`);
  }

  if (treasuriesSetting.show && data.treasuries && data.treasuries.length) {
    const visibleRows = data.treasuries
      .filter((row) => isRowVisible(row.show))
      .sort((a, b) => Number(b.btc || 0) - Number(a.btc || 0));
    const selectedCount = visibleRows.length;
    const barsLimit =
      treasuriesSetting.max_bars && treasuriesSetting.max_bars > 0
        ? treasuriesSetting.max_bars
        : 10;
    const barRows = visibleRows.slice(0, barsLimit);
    const remainingRows = visibleRows.slice(barsLimit);
    const totalBtc = visibleRows.reduce((sum, row) => sum + Number(row.btc || 0), 0);
    const liveBtcPrice = Number(data.live_btc_price || 0);
    const maxBtc = Math.max(...barRows.map((row) => Number(row.btc || 0)), 1);
    const rowsHtml = barRows
      .map((row) => {
        const btc = Number(row.btc || 0);
        const btcFmt = Math.round(btc).toLocaleString("en-US");
        const height = maxBtc > 0 ? Math.max(6, (btc / maxBtc) * 100) : 0;
        const logoCandidates = resolveTreasuryLogoCandidates(row);
        const logoSrc = logoCandidates[0] || "";
        const logoFallbacks = logoCandidates.slice(1).join("|");
        const logoFallbackText = escapeHtml(treasuryLogoFallback(row.entity || ""));
        const logoImg = logoSrc
          ? `<img class="snapshot-treas-logo" src="${escapeHtml(logoSrc, true)}" data-fallbacks="${escapeHtml(logoFallbacks, true)}" alt="${escapeHtml(row.entity || "", true)}" onload="const fb=this.nextElementSibling;if(fb){fb.style.display='none';}" onerror="const list=(this.dataset.fallbacks||'').split('|').filter(Boolean);if(list.length){this.src=list.shift();this.dataset.fallbacks=list.join('|');}else{this.style.display='none';const fb=this.nextElementSibling;if(fb){fb.style.display='flex';}}">`
          : "";
        const logoHtml = `<span class="snapshot-treas-logo-wrap">${logoImg}<span class="snapshot-treas-logo-fallback"${logoImg ? ' style="display:none"' : ""}>${logoFallbackText}</span></span>`;
        const groupHtml = row.holder_group
          ? `<div class="snapshot-treas-group">${escapeHtml(row.holder_group)}</div>`
          : '<div class="snapshot-treas-group"></div>';
        return `<div class="snapshot-treas-item">
        <div class="snapshot-treas-track"><div class="snapshot-treas-fill" style="height:${height.toFixed(2)}%"></div></div>
        ${logoHtml}
        <div class="snapshot-treas-label">${escapeHtml(cleanEntityLabel(row.entity || "", 15))}</div>
        ${groupHtml}
        <div class="snapshot-treas-value">${escapeHtml(btcFmt)} BTC</div>
      </div>`;
      })
      .join("");
    const remainderTable = remainingRows.length
      ? `<div class="snapshot-treas-table-wrap">
      <table class="snapshot-treas-table" role="presentation">
        <thead>
          <tr>
            <th>${escapeHtml(`Top: ${selectedCount}`)}</th>
            <th class="snapshot-treas-cell-num">BTC</th>
            <th class="snapshot-treas-cell-num">% of total</th>
            <th class="snapshot-treas-cell-num">Value (USD)</th>
          </tr>
        </thead>
        <tbody>
          ${remainingRows
            .map((row) => {
              const btc = Number(row.btc || 0);
              const pctTotal = totalBtc > 0 ? (btc / totalBtc) * 100 : 0;
              const usdValue = liveBtcPrice > 0 ? btc * liveBtcPrice : 0;
              return `<tr>
            <td class="snapshot-treas-cell-name">${escapeHtml(row.entity || "")}</td>
            <td class="snapshot-treas-cell-num">${escapeHtml(Math.round(btc).toLocaleString("en-US"))}</td>
            <td class="snapshot-treas-cell-num">${escapeHtml(formatPercent(pctTotal))}</td>
            <td class="snapshot-treas-cell-num">${escapeHtml(liveBtcPrice > 0 ? formatUsdLarge(usdValue) : "-")}</td>
          </tr>`;
            })
            .join("")}
        </tbody>
      </table>
    </div>`
      : "";
    cards.push(`<article class="snapshot-card">
      <h3>${escapeHtml(treasuriesSetting.title)}</h3>
      <div class="snapshot-treas-bars">${rowsHtml}</div>
      ${remainderTable}
    </article>`);
  }

  if (!cards.length) {
    return "";
  }

  return `<section class="snapshot-section">
    <h2 class="snapshot-title">${escapeHtml(data.title || "Bitcoin Data")}</h2>
    <div class="snapshot-grid">${cards.join("")}</div>
  </section>`;
}

function resolveExtraImagePaths(point, meta, imageOptions) {
  if (!parseBool(meta.auto_image_by_order)) {
    return [];
  }

  const useR2Images = Boolean(imageOptions?.useR2Images);
  const r2ImagePrefix = normalizeText(imageOptions?.r2ImagePrefix) || "image";
  const r2ImageExt = normalizeText(imageOptions?.r2ImageExt) || "jpg";
  const imageBaseUrl = normalizeText(meta.image_base_url);
  const maxExtraRaw = parseNumber(meta.max_extra_images, 6);
  const maxExtraImages = Math.max(0, Math.min(20, Math.floor(maxExtraRaw || 6)));
  const sources = [];

  for (let index = 1; index <= maxExtraImages; index += 1) {
    const candidate = useR2Images
      ? `${r2ImagePrefix}${point.order}.${index}.${r2ImageExt}`
      : `${point.order}.${index}.png`;

    if (imageBaseUrl) {
      sources.push(`${imageBaseUrl.replace(/\/+$/, "")}/${candidate.replace(/^\/+/, "")}`);
    } else if (useR2Images) {
      sources.push(`/img/${candidate.replace(/^\/+/, "")}`);
    } else {
      sources.push(`/${candidate.replace(/^\/+/, "")}`);
    }
  }

  return sources;
}

function resolveImageSources(point, meta, imageOptions) {
  const useR2Images = Boolean(imageOptions?.useR2Images);
  const r2ImagePrefix = normalizeText(imageOptions?.r2ImagePrefix) || "image";
  const r2ImageExt = normalizeText(imageOptions?.r2ImageExt) || "jpg";
  const imagePath = normalizeText(point.image_path);
  const imageBaseUrl = normalizeText(meta.image_base_url);
  const candidates = [];
  const seen = new Set();

  const add = (value) => {
    const normalized = normalizeText(value);
    if (!normalized) {
      return;
    }
    const key = normalized.toLowerCase();
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    candidates.push(normalized);
  };

  const pushCandidateWithFallbacks = (candidate, explicitLocalPreferred = false) => {
    const clean = normalizeText(candidate).replace(/^\.\/+/, "");
    if (!clean) {
      return;
    }
    if (looksLikeRemoteImageSource(clean)) {
      add(clean);
      return;
    }

    const localPath = resolveAssetPath(clean, "");
    const basePath = clean.replace(/^\/+/, "");
    if (explicitLocalPreferred) {
      add(localPath);
      if (imageBaseUrl) {
        add(`${imageBaseUrl.replace(/\/+$/, "")}/${basePath}`);
      }
      return;
    }

    if (imageBaseUrl) {
      add(`${imageBaseUrl.replace(/\/+$/, "")}/${basePath}`);
    }
    if (useR2Images) {
      add(`/img/${basePath}`);
    }
    add(localPath);
  };

  if (imagePath) {
    const isExplicitLocal = /^(\/|\.\/|public\/)/i.test(imagePath);
    pushCandidateWithFallbacks(imagePath, isExplicitLocal);
    if (!/\.[a-z0-9]{2,5}$/i.test(imagePath)) {
      ["png", "jpg", "jpeg", "webp"].forEach((ext) =>
        pushCandidateWithFallbacks(`${imagePath}.${ext}`, isExplicitLocal)
      );
    }
    return candidates;
  }

  if (!parseBool(meta.auto_image_by_order)) {
    return candidates;
  }

  if (useR2Images) {
    pushCandidateWithFallbacks(`${r2ImagePrefix}${point.order}.${r2ImageExt}`);
    return candidates;
  }

  const baseNames = [`${point.order}`, `image${point.order}`];
  const exts = ["png", "jpg", "jpeg", "webp"];
  baseNames.forEach((base) => {
    exts.forEach((ext) => pushCandidateWithFallbacks(`${base}.${ext}`));
  });
  return candidates;
}

function resolveAssetPath(rawValue, fallbackPath) {
  const raw = normalizeText(rawValue) || normalizeText(fallbackPath);
  if (!raw) {
    return "";
  }
  if (looksLikeRemoteImageSource(raw)) {
    return raw;
  }
  const trimmed = raw.replace(/^\.\//, "").replace(/^\/+/, "");
  const withoutPublic = trimmed.replace(/^public\//i, "");
  return `/${withoutPublic.replace(/^\/+/, "")}`;
}

function resolveHeroImage(meta) {
  return resolveAssetPath(meta.hero_image_url, "/hero.png");
}

function resolveLogo(meta) {
  return resolveAssetPath(meta.logo_url, "/brand_orange_bg_transparent@2xSite.svg");
}

function looksLikeRemoteImageSource(path) {
  return /^(https?:\/\/|data:|cid:)/i.test(path);
}

function renderContentBlocks(rawValue) {
  const lines = normalizeText(rawValue).replace(/\r\n/g, "\n").split("\n");
  const blocks = [];
  let listOpen = false;

  const closeList = () => {
    if (listOpen) {
      blocks.push("</ul>");
      listOpen = false;
    }
  };

  for (const sourceLine of lines) {
    const line = sourceLine.trim();
    if (!line) {
      closeList();
      continue;
    }

    if (line.startsWith("- ") || line.startsWith("* ")) {
      const item = emphasizeNumbers(line.slice(2).trim());
      if (!listOpen) {
        blocks.push("<ul>");
        listOpen = true;
      }
      blocks.push(`<li>${item}</li>`);
    } else {
      closeList();
      blocks.push(`<p>${emphasizeNumbers(line)}</p>`);
    }
  }

  closeList();
  return blocks.join("\n");
}

function emphasizeNumbers(text) {
  let output = "";
  let start = 0;
  const matches = text.matchAll(NUMBER_PATTERN);
  for (const match of matches) {
    const index = match.index ?? 0;
    output += escapeHtml(text.slice(start, index));
    output += `<strong>${escapeHtml(match[0])}</strong>`;
    start = index + match[0].length;
  }
  output += escapeHtml(text.slice(start));
  return output;
}

function indentBlock(text, spaces) {
  if (!text) {
    return "";
  }
  const prefix = " ".repeat(spaces);
  return text
    .split("\n")
    .map((line) => (line ? `${prefix}${line}` : ""))
    .join("\n");
}

function parseOrder(value, rowNumber) {
  const text = normalizeText(value);
  if (!text) {
    throw new Error(`Missing order value at points row ${rowNumber}.`);
  }
  const parsed = Number(text);
  if (!Number.isInteger(parsed)) {
    throw new Error(`Order value must be a whole number at points row ${rowNumber}.`);
  }
  if (parsed < 1 || parsed > MAX_POINTS) {
    throw new Error(
      `Order value must be between 1 and ${MAX_POINTS} at points row ${rowNumber}.`
    );
  }
  return parsed;
}

function renderBlockHeight(value) {
  const clean = normalizeText(value);
  if (!clean) {
    return "n/a";
  }
  const numeric = parseNumber(clean, Number.NaN);
  if (Number.isFinite(numeric)) {
    return formatBtcInteger(numeric);
  }
  return clean;
}

function parseBool(value) {
  return ["1", "true", "yes", "y", "on"].includes(normalizeText(value).toLowerCase());
}

function isRowVisible(value) {
  const normalized = normalizeText(value).toLowerCase();
  if (!normalized) {
    return true;
  }
  return !["0", "false", "no", "n", "off"].includes(normalized);
}

function toPositiveInt(value, fallback = null) {
  const parsed = Number.parseInt(String(value ?? ""), 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return fallback;
  }
  return parsed;
}

function safeColor(value, fallback = "#ff4202") {
  const color = normalizeText(value);
  if (!color) {
    return fallback;
  }
  if (/^#[0-9a-f]{3,8}$/i.test(color)) return color;
  if (/^rgb(a)?\(/i.test(color)) return color;
  if (/^hsl(a)?\(/i.test(color)) return color;
  return fallback;
}

function normalizeLogoSlug(value) {
  return normalizeText(value)
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function entityToLogoPath(entity) {
  const slug = normalizeLogoSlug(entity);
  if (!slug) {
    return "";
  }
  return `/${slug}.png`;
}

function resolveTreasuryLogoCandidates(row) {
  const candidates = [];
  const seen = new Set();

  const addCandidate = (value) => {
    const resolved = resolveAssetPath(value, "");
    if (!resolved) {
      return;
    }
    const key = resolved.toLowerCase();
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    candidates.push(resolved);
  };

  const addFilenameFamily = (value) => {
    const raw = normalizeText(value);
    if (!raw) {
      return;
    }
    const trimmed = raw.replace(/^\.\/+/, "").replace(/^public\//i, "").replace(/^\/+/, "");
    if (!trimmed) {
      return;
    }
    const names = new Set([trimmed, trimmed.toLowerCase()]);
    names.forEach((name) => {
      addCandidate(`/${name}`);
      if (!name.includes("/")) {
        addCandidate(`/logos/${name}`);
      }
      if (!/\.[a-z0-9]{2,5}$/i.test(name)) {
        [".png", ".webp", ".jpg", ".jpeg", ".svg"].forEach((ext) => {
          addCandidate(`/${name}${ext}`);
          if (!name.includes("/")) {
            addCandidate(`/logos/${name}${ext}`);
          }
        });
      }
      const slug = normalizeLogoSlug(name);
      if (slug) {
        [".png", ".webp", ".jpg", ".jpeg", ".svg"].forEach((ext) => {
          addCandidate(`/${slug}${ext}`);
          addCandidate(`/logos/${slug}${ext}`);
          const hyphenSlug = slug.replace(/_/g, "-");
          if (hyphenSlug !== slug) {
            addCandidate(`/${hyphenSlug}${ext}`);
            addCandidate(`/logos/${hyphenSlug}${ext}`);
          }
        });
      }
    });
  };

  const addSlugVariants = (value) => {
    const slug = normalizeLogoSlug(value);
    if (!slug) {
      return;
    }
    [".png", ".webp", ".jpg", ".jpeg", ".svg"].forEach((ext) => {
      addCandidate(`/${slug}${ext}`);
      addCandidate(`/logos/${slug}${ext}`);
      const hyphenSlug = slug.replace(/_/g, "-");
      if (hyphenSlug !== slug) {
        addCandidate(`/${hyphenSlug}${ext}`);
        addCandidate(`/logos/${hyphenSlug}${ext}`);
      }
    });
    const parts = slug.split("_").filter(Boolean);
    if (parts.length > 1) {
      const shortSlug = parts[0];
      [".png", ".webp", ".jpg", ".jpeg", ".svg"].forEach((ext) => {
        addCandidate(`/${shortSlug}${ext}`);
        addCandidate(`/logos/${shortSlug}${ext}`);
      });
    }
  };

  const explicit = normalizeText(row && row.logo);
  if (explicit) {
    explicit
      .split(/[|,;]/)
      .map((part) => part.trim())
      .filter(Boolean)
      .forEach((part) => addFilenameFamily(part));
  }

  const entity = normalizeText(row && row.entity);
  if (entity) {
    addSlugVariants(entity);
    const withoutParens = normalizeText(entity.replace(/\(.*?\)/g, " "));
    if (withoutParens && withoutParens !== entity) {
      addSlugVariants(withoutParens);
    }
    const words = (withoutParens || entity).split(/\s+/).filter(Boolean);
    if (words.length >= 1) {
      addSlugVariants(words.slice(0, 2).join(" "));
    }
  }

  return candidates;
}

function normalizeAssetKey(value) {
  return normalizeText(value).toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function isBtcAsset(value) {
  const key = normalizeAssetKey(value);
  return key === "BTC" || key === "BITCOIN" || key === "BTCUSD" || key === "XBT" || key === "XBTUSD";
}

function parseNumber(value, defaultValue = 0) {
  const text = normalizeText(value).replace(/,/g, "");
  if (!text) {
    return defaultValue;
  }
  const parsed = Number(text);
  if (!Number.isFinite(parsed)) {
    return defaultValue;
  }
  return parsed;
}

function normalizeText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function formatBtcInteger(value) {
  return Math.round(value).toLocaleString("en-US");
}

function formatBtcCompact(value) {
  const absValue = Math.abs(value);
  if (absValue >= 1_000_000) {
    const rendered = (value / 1_000_000).toFixed(2).replace(/\.0+$/, "").replace(/(\.\d*?)0+$/, "$1");
    return `${rendered}M BTC`;
  }
  if (absValue >= 1_000) {
    const rendered = Math.round(value / 1_000).toLocaleString("en-US");
    return `${rendered}K BTC`;
  }
  return `${formatBtcInteger(value)} BTC`;
}

function formatPercent(value) {
  const rendered = Number(value).toFixed(1).replace(/\.0$/, "");
  return `${rendered}%`;
}

function escapeHtml(value, escapeQuotes = false) {
  const text = String(value ?? "");
  const escaped = text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  if (!escapeQuotes) {
    return escaped;
  }
  return escaped.replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
