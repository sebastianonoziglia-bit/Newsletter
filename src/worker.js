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

      const [
        metaRows,
        pointsRows,
        livePriceRows,
        btcPriceRows,
        treasuriesRows,
        circulatingRows,
        liquidationsRows,
      ] =
        await Promise.all([
          fetchGoogleTabRows(sheetId, metaTab, true),
          fetchGoogleTabRows(sheetId, pointsTab, true),
          fetchGoogleTabRows(sheetId, livePricesTab, false),
          fetchGoogleTabRows(sheetId, btcPriceTab, false),
          fetchGoogleTabRows(sheetId, treasuriesTab, false),
          fetchGoogleTabRows(sheetId, circulatingTab, false),
          fetchGoogleTabRows(sheetId, liquidationsTab, false),
        ]);

      const meta = readMeta(metaRows);
      const points = readPoints(pointsRows);
      const liveBtc = livePriceRows.length
        ? safeOptionalParse(livePricesTab, () => readLiveBtcPrice(livePriceRows), null)
        : null;
      const btcPricePoints = btcPriceRows.length
        ? safeOptionalParse(btcPriceTab, () => readBtcPricePoints(btcPriceRows), [])
        : [];
      const treasuryBars = treasuriesRows.length
        ? safeOptionalParse(treasuriesTab, () => readTreasuryBars(treasuriesRows), [])
        : [];
      let circulatingMetric = circulatingRows.length
        ? safeOptionalParse(circulatingTab, () => readCirculatingMetric(circulatingRows), null)
        : null;
      const liquidationBars = liquidationsRows.length
        ? safeOptionalParse(liquidationsTab, () => readLiquidationBars(liquidationsRows), [])
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

  for (let i = rows.length - 1; i >= 1; i -= 1) {
    const row = rows[i];
    const asset = normalizeText(row[mapping.asset]).toUpperCase();
    if (!["BITCOIN", "BTC-USD", "BTC"].includes(asset)) {
      continue;
    }
    const price = parseNumber(row[closeIdx], Number.NaN);
    if (!Number.isFinite(price)) {
      continue;
    }
    return {
      price,
      date: normalizeText(row[mapping.date]),
      currency: currencyIdx >= 0 ? normalizeText(row[currencyIdx]) || "USD" : "USD",
    };
  }
  return null;
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

  const ordered = points.map((item) => item.point);
  return ordered.slice(-limit);
}

function readTreasuryBars(rows, limit = 6) {
  if (!rows.length) {
    return [];
  }
  const mapping = headerIndexMap(rows, ["entity", "btc"]);
  const rowTypeIdx = "row_type" in mapping ? mapping.row_type : -1;
  const bars = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const rowType = rowTypeIdx >= 0 ? normalizeText(row[rowTypeIdx]).toLowerCase() : "";
    if (rowType && rowType !== "entity") {
      continue;
    }
    const entity = normalizeText(row[mapping.entity]);
    const btc = parseNumber(row[mapping.btc], 0);
    if (!entity || btc <= 0) {
      continue;
    }
    bars.push({ entity, btc });
  }

  bars.sort((a, b) => b.btc - a.btc);
  return bars.slice(0, limit);
}

function readCirculatingMetric(rows) {
  if (!rows.length) {
    return null;
  }
  const mapping = headerIndexMap(rows, ["circulating_supply_btc", "max_supply_btc"]);
  const asOfIdx = "as_of_date" in mapping ? mapping.as_of_date : -1;
  const noteIdx = "note" in mapping ? mapping.note : -1;
  const candidates = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
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

  const collect = (monthlyOnly) => {
    const bars = [];
    for (let i = 1; i < rows.length; i += 1) {
      const row = rows[i];
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
      bars.push({ label, longs, shorts, total });
    }
    return bars;
  };

  const bars = collect(true);
  const finalBars = bars.length ? bars : collect(false);
  return finalBars.slice(-limit);
}

function renderHtml(
  meta,
  points,
  btcPricePoints,
  treasuryBars,
  circulatingMetric,
  liquidationBars,
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
  const footerInstagramIcon = escapeHtml(resolveAssetPath(meta.footer_instagram_icon, "/instagram.png"), true);
  const footerXIcon = escapeHtml(resolveAssetPath(meta.footer_x_icon, "/x:twitter.png"), true);
  const footerLinkedinIcon = escapeHtml(resolveAssetPath(meta.footer_linkedin_icon, "/linkedin.png"), true);

  const pointsHtml = points
    .map((point) => renderPoint(point, meta, imageOptions))
    .join("");
  const marketHtml = renderMarketSection(
    meta,
    btcPricePoints,
    treasuryBars,
    circulatingMetric,
    liquidationBars,
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
      .toolbar { width: 100%; max-width: 680px; margin: 0 auto; display: flex; justify-content: flex-end; padding: 12px 0 8px; }
      .download-pdf-btn { border: 1px solid #ff4202; border-radius: 999px; padding: 8px 14px; background: #ffffff; color: #ff4202; font: 600 12px/1 "Poppins", Arial, sans-serif; cursor: pointer; }
      .download-pdf-btn:hover { background: #fff4ef; }
      .wrapper { width: 100%; background: #f5f5f5; padding: 32px 0; }
      .container { width: 680px; max-width: 680px; background: #ffffff; border: 1px solid #e6e6e6; border-radius: 16px; overflow: hidden; }
      .divider { height: 4px; background: #ff4202; line-height: 4px; }
      .hero { position: relative; overflow: hidden; background: #0a0a0a; min-height: 260px; display: flex; flex-direction: column; justify-content: flex-end; }
      .hero-bg { position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; opacity: 0.82; }
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
      .market { background: #070707; color: #f4f4f4; border-top: 1px solid #171717; }
      .market h2 { color: #ffffff; margin-bottom: 8px; }
      .market-intro { margin: 0 0 12px; color: #b8b8b8; font-size: 13px; }
      .market-live { margin: 0 0 12px; display: inline-flex; gap: 6px; align-items: baseline; font-size: 12px; color: #ffcfb8; background: rgba(255,66,2,0.16); border: 1px solid rgba(255,66,2,0.35); border-radius: 999px; padding: 4px 10px; }
      .market-live strong { color: #ffffff; }
      .market-live span { color: #ffcfb8; }
      .market-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
      .market-card { border: 1px solid #1e1e1e; border-radius: 12px; padding: 12px; background: #101010; }
      .market-card h3 { margin: 0 0 10px; color: #ffffff; font-size: 14px; font-weight: 700; }
      .market-price-card { grid-column: 1 / -1; }
      .market-price-svg { width: 100%; height: auto; display: block; }
      .market-bars { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; align-items: end; min-height: 180px; }
      .market-bar-item { display: flex; flex-direction: column; align-items: center; gap: 6px; }
      .market-bar-track { width: 100%; max-width: 72px; height: 128px; border-radius: 8px; border: 1px solid #2a2a2a; background: linear-gradient(180deg, #141414 0%, #0b0b0b 100%); overflow: hidden; display: flex; flex-direction: column; justify-content: flex-end; }
      .market-bar-fill { width: 100%; background: linear-gradient(180deg, #ff8b61 0%, #ff4202 100%); }
      .market-liq-track { justify-content: flex-end; }
      .market-liq-short { width: 100%; background: #8f3a1d; }
      .market-liq-long { width: 100%; background: #ff4202; }
      .market-bar-label { margin: 0; font-size: 11px; color: #d0d0d0; text-align: center; line-height: 1.3; }
      .market-bar-value { margin: 0; font-size: 11px; color: #ffcfb8; }
      .market-legend { display: flex; gap: 14px; margin: 0 0 8px; font-size: 11px; color: #bcbcbc; }
      .market-legend .dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 6px; }
      .market-legend .dot-long { background: #ff4202; }
      .market-legend .dot-short { background: #8f3a1d; }
      .market-circ-card .market-circ-value { margin: 0 0 8px; font-size: 22px; font-weight: 700; color: #ffffff; }
      .market-progress-track { width: 100%; height: 14px; border: 1px solid #2a2a2a; border-radius: 999px; background: #0c0c0c; overflow: hidden; }
      .market-progress-fill { height: 100%; background: linear-gradient(90deg, #ff4202 0%, #ff8b61 100%); }
      .market-subnote { margin: 8px 0 0; font-size: 12px; color: #b9b9b9; }
      .market-note { margin: 8px 0 0; font-size: 11px; color: #8e8e8e; }
      .market-empty { margin: 0; color: #9a9a9a; font-size: 12px; }
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
      .footer-social-btn { display: inline-flex; align-items: center; justify-content: center; width: 36px; height: 36px; background: rgba(255,255,255,0.08); border-radius: 10px; }
      .footer-social-btn img { width: 18px; height: 18px; filter: brightness(0) invert(1); }
      .footer-copy { margin: 14px 0 0; padding-top: 14px; border-top: 1px solid rgba(255,255,255,0.08); font-size: 11px; color: rgba(255,255,255,0.3); text-align: center; }
      @media (max-width: 720px) {
        .toolbar { padding: 10px 16px 6px; box-sizing: border-box; }
        .wrapper { padding: 16px 0; }
        .container { width: 100%; max-width: 100%; border-radius: 0; }
        .section { padding: 18px 20px; }
        .hero-content { padding: 20px; }
        .hero-title { font-size: 22px; }
        .market-grid { grid-template-columns: 1fr; }
        .market-price-card { grid-column: auto; }
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
  const imageSrc = resolveImagePath(point, meta, imageOptions);
  const imageBlock = renderImageBlock(point, imageSrc);
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

function renderImageBlock(point, imageSrc) {
  if (!imageSrc) {
    return "";
  }
  const caption = point.image_caption || point.title;
  return (
    '<div class="image">\n' +
    `  <img src="${escapeHtml(imageSrc, true)}" alt="${escapeHtml(point.title)}" onerror="this.closest('.image').style.display='none'">\n` +
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

function renderMarketSection(
  meta,
  btcPricePoints,
  treasuryBars,
  circulatingMetric,
  liquidationBars,
  liveBtc
) {
  if (
    !btcPricePoints.length &&
    !treasuryBars.length &&
    !circulatingMetric &&
    !liquidationBars.length
  ) {
    return "";
  }

  const sectionTitle =
    normalizeText(meta.market_section_title) || "Bitcoin Market Dashboard";
  const sectionIntro =
    normalizeText(meta.market_section_intro) ||
    "Auto-rendered from BTC Price, Liquidations, Treasuries, and Circulating BTC tabs.";

  const liveChip = liveBtc
    ? `<p class="market-live">Live BTC: <strong>${escapeHtml(formatUsd(liveBtc.price))}</strong>${
        liveBtc.date ? ` <span>(${escapeHtml(liveBtc.date)})</span>` : ""
      }</p>`
    : "";

  return `
            <tr>
              <td class="section market">
                <h2>${escapeHtml(sectionTitle)}</h2>
                <p class="market-intro">${escapeHtml(sectionIntro)}</p>
                ${liveChip}
                <div class="market-grid">
                  <div class="market-card market-price-card">
                    <h3>BTC Price</h3>
                    ${buildBtcPriceChartSvg(btcPricePoints)}
                  </div>
                  <div class="market-card">
                    <h3>Liquidations</h3>
                    ${renderLiquidationsBars(liquidationBars)}
                  </div>
                  <div class="market-card">
                    <h3>Treasuries (Top Holders)</h3>
                    ${renderTreasuryBars(treasuryBars)}
                  </div>
                  <div class="market-card market-circ-card">
                    <h3>Circulating BTC</h3>
                    ${renderCirculatingCard(circulatingMetric)}
                  </div>
                </div>
              </td>
            </tr>
`;
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

function resolveImagePath(point, meta, imageOptions) {
  const useR2Images = Boolean(imageOptions?.useR2Images);
  const r2ImagePrefix = normalizeText(imageOptions?.r2ImagePrefix) || "image";
  const r2ImageExt = normalizeText(imageOptions?.r2ImageExt) || "jpg";
  const imagePath = normalizeText(point.image_path);
  let candidate = "";

  if (imagePath) {
    candidate = imagePath;
  } else if (parseBool(meta.auto_image_by_order)) {
    candidate = useR2Images
      ? `${r2ImagePrefix}${point.order}.${r2ImageExt}`
      : `${point.order}.png`;
  }

  if (!candidate) {
    return "";
  }
  if (looksLikeRemoteImageSource(candidate)) {
    return candidate;
  }

  const imageBaseUrl = normalizeText(meta.image_base_url);
  if (imageBaseUrl) {
    return `${imageBaseUrl.replace(/\/+$/, "")}/${candidate.replace(/^\/+/, "")}`;
  }

  if (useR2Images) {
    return `/img/${candidate.replace(/^\/+/, "")}`;
  }

  return `/${candidate.replace(/^\/+/, "")}`;
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
