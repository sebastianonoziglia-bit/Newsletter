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
  footer_logo_url: "/logotosite.png",
  footer_instagram_icon: "/instagram.png",
  footer_x_icon: "/x:twitter.png",
  footer_linkedin_icon: "/linkedin.png",
  image_dir: ".",
  auto_image_by_order: "true",
  logo_url: "/brand_orange_bg_transparent@2xSite.svg",
};

const DEFAULT_DISTRIBUTION = [
  ["Individuals", 13660000, "rgb(255, 66, 2)"],
  ["Lost Bitcoin", 1570000, "rgb(153, 153, 153)"],
  ["Funds & ETFs", 1490000, "rgb(255, 140, 90)"],
  ["Businesses", 1390000, "rgb(255, 107, 61)"],
  ["To Be Mined", 1040000, "rgb(204, 204, 204)"],
  ["Satoshi / Patoshi", 968000, "rgb(255, 200, 150)"],
  ["Governments", 432000, "rgb(255, 173, 120)"],
  ["Other Entities", 421000, "rgb(255, 227, 180)"],
];

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
      const distributionTab =
        normalizeText(env.GOOGLE_DISTRIBUTION_TAB) || "distribution";
      const livePricesTab =
        normalizeText(env.GOOGLE_LIVE_PRICES_TAB) || "live_prices";

      const [metaRows, pointsRows, distributionRows, livePriceRows] =
        await Promise.all([
          fetchGoogleTabRows(sheetId, metaTab, true),
          fetchGoogleTabRows(sheetId, pointsTab, true),
          fetchGoogleTabRows(sheetId, distributionTab, false),
          fetchGoogleTabRows(sheetId, livePricesTab, false),
        ]);

      const meta = readMeta(metaRows);
      const points = readPoints(pointsRows);
      const maxSupply = parseNumber(meta.max_supply_btc, 21_000_000);
      const distribution = distributionRows.length
        ? readDistribution(distributionRows, maxSupply)
        : defaultDistributionSegments(maxSupply);
      const liveBtc = livePriceRows.length ? readLiveBtcPrice(livePriceRows) : null;

      const html = renderHtml(meta, points, distribution, liveBtc, {
        useR2Images: Boolean(env.IMAGES),
        r2ImagePrefix: normalizeText(env.R2_IMAGE_PREFIX) || "image",
        r2ImageExt: normalizeText(env.R2_IMAGE_EXT) || "jpg",
      });
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

function defaultDistributionSegments(maxSupplyBtc) {
  return finalizeDistributionSegments(
    DEFAULT_DISTRIBUTION.map(([category, amountBtc, color]) => ({
      category,
      amount_btc: Number(amountBtc),
      percent: 0,
      color,
    })),
    maxSupplyBtc
  );
}

function finalizeDistributionSegments(segments, maxSupplyBtc) {
  if (!segments.length) {
    return [];
  }

  let denominator = maxSupplyBtc > 0 ? maxSupplyBtc : segments.reduce((acc, s) => acc + s.amount_btc, 0);
  if (denominator <= 0) {
    denominator = 1;
  }

  let normalized = segments.map((segment) => ({
    ...segment,
    percent: segment.percent > 0 ? segment.percent : (segment.amount_btc / denominator) * 100,
  }));

  const totalPercent = normalized.reduce((acc, s) => acc + s.percent, 0);
  if (totalPercent > 0) {
    normalized = normalized.map((segment) => ({
      ...segment,
      percent: (segment.percent / totalPercent) * 100,
    }));
  }

  normalized.sort((a, b) => b.amount_btc - a.amount_btc);
  return normalized;
}

function readDistribution(rows, maxSupplyBtc) {
  if (!rows.length) {
    return defaultDistributionSegments(maxSupplyBtc);
  }

  const mapping = headerIndexMap(rows, ["category", "amount_btc", "color"]);
  const segments = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const category = normalizeText(row[mapping.category]);
    const amountBtc = parseNumber(row[mapping.amount_btc], 0);
    const color = normalizeText(row[mapping.color]) || "rgb(255, 66, 2)";
    const percent = "percent" in mapping ? parseNumber(row[mapping.percent], 0) : 0;

    if (!category && amountBtc <= 0) {
      continue;
    }
    if (!category) {
      throw new Error("Distribution row is missing category.");
    }
    if (amountBtc < 0) {
      throw new Error(`Distribution amount cannot be negative for category '${category}'.`);
    }

    segments.push({
      category,
      amount_btc: amountBtc,
      percent,
      color,
    });
  }

  if (!segments.length) {
    return defaultDistributionSegments(maxSupplyBtc);
  }
  return finalizeDistributionSegments(segments, maxSupplyBtc);
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

function renderHtml(meta, points, distribution, liveBtc, imageOptions) {
  const title = escapeHtml(meta.main_title);
  const subtitle = escapeHtml(meta.subtitle);
  const eyebrow = escapeHtml(meta.eyebrow);
  const blockHeight = escapeHtml(renderBlockHeight(meta.block_height));
  const tldrTitle = escapeHtml(meta.tldr_title);
  const tldrContent = indentBlock(renderContentBlocks(meta.tldr_content), 16);
  const conclusionTitle = escapeHtml(meta.conclusion_title);
  const conclusionContent = indentBlock(renderContentBlocks(meta.conclusion_content), 16);
  const ctaUrl = escapeHtml(meta.cta_url, true);
  const ctaLabel = escapeHtml(meta.cta_label);
  const addressLine = escapeHtml(meta.address_line);
  const footerLine = escapeHtml(meta.footer_line);
  const resolvedLogo = resolveLogo(meta);
  const logoUrl = escapeHtml(resolvedLogo, true);
  const footerLogoUrl = escapeHtml(normalizeText(meta.footer_logo_url) || "/logotosite.png", true);
  const footerInstagramIcon = escapeHtml(normalizeText(meta.footer_instagram_icon) || "/instagram.png", true);
  const footerXIcon = escapeHtml(normalizeText(meta.footer_x_icon) || "/x:twitter.png", true);
  const footerLinkedinIcon = escapeHtml(normalizeText(meta.footer_linkedin_icon) || "/linkedin.png", true);

  const pointsHtml = points
    .map((point) => renderPoint(point, meta, imageOptions))
    .join("");
  const snapshotHtml = renderSnapshotSection(meta, distribution);
  const liveBtcHtml = renderLiveBtcChip(liveBtc);

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
      .header { padding: 28px 32px 18px; }
      .logo { margin: 0 0 16px; text-align: center; }
      .logo img { width: 190px; margin: 0 auto; }
      .eyebrow { color: #ff4202; font-weight: 700; font-size: 12px; letter-spacing: 1px; text-transform: uppercase; margin: 0 0 6px; }
      h1 { margin: 6px 0 6px; font-size: 28px; line-height: 1.2; font-weight: 700; }
      .subtitle { margin: 0; color: #5f5f5f; font-size: 14px; line-height: 1.6; }
      .block-height { margin: 12px 0 0; display: inline-block; font-size: 12px; line-height: 1.4; color: #8f3a1a; background: #fff1eb; border: 1px solid #ffd6c8; border-radius: 999px; padding: 6px 10px; }
      .live-price { margin: 10px 0 0; font-size: 12px; color: #4f4f4f; }
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
      .snapshot { background: #fcfcfc; }
      .snapshot-intro { margin: 0 0 14px; font-size: 14px; color: #4f4f4f; }
      .snapshot-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }
      .snapshot-card { border: 1px solid #ececec; border-radius: 12px; padding: 12px; background: #ffffff; }
      .snapshot-card h3 { margin: 0 0 10px; font-size: 14px; font-weight: 700; color: #1f1f1f; }
      .snapshot-ownership-viz { display: flex; justify-content: center; margin: 0 0 10px; }
      .snapshot-donut { width: 132px; height: 132px; display: block; }
      .snapshot-donut-segment { fill: none; stroke-width: 24; transform: rotate(-90deg); transform-origin: 60px 60px; }
      .snapshot-donut-label { font-size: 10px; fill: #8a8a8a; }
      .snapshot-donut-value { font-size: 10px; fill: #5f5f5f; }
      .snapshot-bar { width: 100%; height: 24px; border: 1px solid #e6e6e6; border-radius: 8px; overflow: hidden; display: flex; }
      .snapshot-bar-segment { height: 100%; min-width: 2px; }
      .snapshot-legend { margin-top: 10px; display: grid; gap: 6px; }
      .snapshot-legend-item { display: grid; grid-template-columns: 12px 1fr auto; align-items: center; gap: 8px; }
      .snapshot-dot { width: 10px; height: 10px; border-radius: 999px; display: inline-block; }
      .snapshot-name { font-size: 12px; color: #3f3f3f; }
      .snapshot-value { font-size: 12px; color: #5f5f5f; white-space: nowrap; }
      .snapshot-circ-value { margin: 0 0 8px; font-size: 22px; font-weight: 700; color: #1f1f1f; }
      .snapshot-progress-track { width: 100%; height: 14px; border: 1px solid #e0e0e0; border-radius: 999px; background: #f0f0f0; overflow: hidden; }
      .snapshot-progress-fill { height: 100%; background: linear-gradient(90deg, #ff4202 0%, #ff8b61 100%); }
      .snapshot-circ-note { margin: 8px 0 0; font-size: 12px; color: #666666; }
      .snapshot-footnote { margin: 12px 0 0; font-size: 12px; color: #7a7a7a; }
      .tldr { background: #fff8ec; border-top: 2px solid #ff4202; }
      .conclusion { background: #fff7f3; border-top: 2px solid #ff4202; }
      .footer { padding: 18px 32px 28px; font-size: 12px; color: #7a7a7a; }
      .footer p { margin: 0; }
      .footer p + p { margin-top: 6px; }
      .footer-links { margin-top: 14px; padding-top: 12px; border-top: 1px solid #ececec; display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
      .footer-panel { border: 1px solid #ececec; border-radius: 12px; background: #fafafa; padding: 10px; }
      .footer-panel-title { margin: 0 0 8px; font-size: 11px; font-weight: 700; letter-spacing: 0.6px; text-transform: uppercase; color: #8a8a8a; font-family: "Poppins", Arial, sans-serif; }
      .footer-logo-link { display: inline-flex; align-items: center; gap: 8px; color: #1f1f1f; font-size: 12px; font-weight: 600; }
      .footer-logo-link img { width: 52px; height: 52px; object-fit: cover; border-radius: 20px; border: 1px solid #e0e0e0; }
      .footer-social { display: flex; flex-direction: column; gap: 6px; }
      .footer-social a { display: inline-flex; align-items: center; gap: 8px; color: #1f1f1f; font-size: 12px; font-weight: 600; font-family: "Poppins", Arial, sans-serif; }
      .footer-social img { width: 30px; height: 30px; object-fit: contain; border-radius: 8px; }
      @media (max-width: 720px) {
        .toolbar { padding: 10px 16px 6px; box-sizing: border-box; }
        .wrapper { padding: 16px 0; }
        .container { width: 100%; max-width: 100%; border-radius: 0; }
        .header, .section, .footer { padding: 18px 20px; }
        .snapshot-grid { grid-template-columns: 1fr; }
        .footer-links { grid-template-columns: 1fr; }
        h1 { font-size: 24px; }
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
              <td class="header">
                <div class="logo">
                  <img src="${logoUrl}" alt="Globalite">
                </div>
                <p class="eyebrow">${eyebrow}</p>
                <h1>${title}</h1>
                <p class="subtitle">${subtitle}</p>
                <p class="block-height">This article was written at block height: <strong>${blockHeight}</strong></p>
                ${liveBtcHtml}
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
${snapshotHtml}
            <tr>
              <td class="footer">
                <p>${footerLine}</p>
                <p>${addressLine}</p>
                <div class="footer-links">
                  <div class="footer-panel">
                    <p class="footer-panel-title">Site</p>
                    <a class="footer-logo-link" href="https://globalite.co" target="_blank" rel="noopener noreferrer">
                      <img src="${footerLogoUrl}" alt="Globalite logo">
                      <span>globalite.co</span>
                    </a>
                  </div>
                  <div class="footer-panel">
                    <p class="footer-panel-title">Socials</p>
                    <div class="footer-social">
                      <a href="https://www.instagram.com/globalite.sa/" target="_blank" rel="noopener noreferrer"><img src="${footerInstagramIcon}" alt="Instagram"><span>Instagram</span></a>
                      <a href="https://x.com/globalite_sa" target="_blank" rel="noopener noreferrer"><img src="${footerXIcon}" alt="X"><span>X</span></a>
                      <a href="https://www.linkedin.com/company/globalite-sa" target="_blank" rel="noopener noreferrer"><img src="${footerLinkedinIcon}" alt="LinkedIn"><span>LinkedIn</span></a>
                    </div>
                  </div>
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

function renderSnapshotSection(meta, distribution) {
  const snapshotTitle = escapeHtml(meta.snapshot_title || "At The Time Of Writing");
  const snapshotIntro = escapeHtml(meta.snapshot_intro || "");
  const snapshotNote = escapeHtml(meta.snapshot_note || "");
  let maxSupplyBtc = parseNumber(meta.max_supply_btc, 21_000_000);
  if (maxSupplyBtc <= 0) {
    maxSupplyBtc = 21_000_000;
  }

  let circulatingBtc = parseNumber(meta.circulating_supply_btc, 0);
  if (circulatingBtc <= 0) {
    circulatingBtc = Math.max(
      0,
      maxSupplyBtc -
        distribution
          .filter((segment) => segment.category.trim().toLowerCase() === "to be mined")
          .reduce((acc, segment) => acc + segment.amount_btc, 0)
    );
  }

  const circulationPct = maxSupplyBtc > 0 ? (circulatingBtc / maxSupplyBtc) * 100 : 0;
  const hashrateEhS = parseNumber(meta.hashrate_eh_s, 0);
  let hashrateScale = parseNumber(meta.hashrate_scale_eh_s, 1000);
  if (hashrateScale <= 0) {
    hashrateScale = 1000;
  }
  const hashratePct = hashrateScale > 0 ? (hashrateEhS / hashrateScale) * 100 : 0;

  const barSegments = distribution
    .map(
      (segment) =>
        `                    <div class="snapshot-bar-segment" style="width:${Math.max(0, segment.percent).toFixed(6)}%;background:${escapeHtml(segment.color, true)};" title="${escapeHtml(segment.category)}: ${escapeHtml(formatBtcCompact(segment.amount_btc))} (${escapeHtml(formatPercent(segment.percent))})"></div>`
    )
    .join("\n");

  const circumference = 2 * Math.PI * 45;
  let consumed = 0;
  const donutSegments = [];
  for (const segment of distribution) {
    const arc = circumference * (Math.max(0, segment.percent) / 100);
    donutSegments.push(
      `                      <circle class="snapshot-donut-segment" cx="60" cy="60" r="45" stroke="${escapeHtml(segment.color, true)}" stroke-dasharray="${arc.toFixed(6)} ${circumference.toFixed(6)}" stroke-dashoffset="${(-consumed).toFixed(6)}"><title>${escapeHtml(segment.category)}: ${escapeHtml(formatBtcCompact(segment.amount_btc))} (${escapeHtml(formatPercent(segment.percent))})</title></circle>`
    );
    consumed += arc;
  }

  const legendRows = distribution
    .map(
      (segment) =>
        "                    <div class=\"snapshot-legend-item\">" +
        `<span class="snapshot-dot" style="background:${escapeHtml(segment.color, true)}"></span>` +
        `<span class="snapshot-name">${escapeHtml(segment.category)}</span>` +
        `<span class="snapshot-value">${escapeHtml(formatBtcCompact(segment.amount_btc))} (${escapeHtml(formatPercent(segment.percent))})</span>` +
        "</div>"
    )
    .join("\n");

  return `
            <tr>
              <td class="section snapshot">
                <h2>${snapshotTitle}</h2>
                <p class="snapshot-intro">${snapshotIntro}</p>
                <div class="snapshot-grid">
                  <div class="snapshot-card">
                    <h3>Ownership Distribution</h3>
                    <div class="snapshot-ownership-viz">
                      <svg class="snapshot-donut" viewBox="0 0 120 120" aria-label="Ownership distribution donut chart">
                        <circle cx="60" cy="60" r="45" fill="none" stroke="#ececec" stroke-width="24"></circle>
${donutSegments.join("\n")}
                        <circle cx="60" cy="60" r="30" fill="#ffffff"></circle>
                        <text x="60" y="56" text-anchor="middle" class="snapshot-donut-label">Supply</text>
                        <text x="60" y="72" text-anchor="middle" class="snapshot-donut-value">${escapeHtml(formatBtcInteger(maxSupplyBtc))}</text>
                      </svg>
                    </div>
                    <div class="snapshot-bar">
${barSegments}
                    </div>
                    <div class="snapshot-legend">
${legendRows}
                    </div>
                  </div>
                  <div class="snapshot-card">
                    <h3>Bitcoin In Circulation At Write Time</h3>
                    <p class="snapshot-circ-value">${escapeHtml(formatBtcInteger(circulatingBtc))} BTC</p>
                    <div class="snapshot-progress-track">
                      <div class="snapshot-progress-fill" style="width:${Math.max(0, Math.min(100, circulationPct)).toFixed(6)}%;"></div>
                    </div>
                    <p class="snapshot-circ-note">${escapeHtml(formatPercent(circulationPct))} of ${escapeHtml(formatBtcInteger(maxSupplyBtc))} BTC max supply</p>
                  </div>
                  <div class="snapshot-card">
                    <h3>Network Hashrate (Daily)</h3>
                    <p class="snapshot-circ-value">${escapeHtml(formatBtcInteger(hashrateEhS))} EH/s</p>
                    <div class="snapshot-progress-track">
                      <div class="snapshot-progress-fill" style="width:${Math.max(0, Math.min(100, hashratePct)).toFixed(6)}%;"></div>
                    </div>
                    <p class="snapshot-circ-note">${escapeHtml(formatPercent(hashratePct))} of ${escapeHtml(formatBtcInteger(hashrateScale))} EH/s reference scale</p>
                  </div>
                </div>
                <p class="snapshot-footnote">${snapshotNote}</p>
              </td>
            </tr>
`;
}

function renderLiveBtcChip(liveBtc) {
  if (!liveBtc) {
    return "";
  }
  const formatted = formatCurrency(liveBtc.price, liveBtc.currency || "USD");
  const updatedText = liveBtc.date ? ` | updated ${escapeHtml(liveBtc.date)}` : "";
  return `<p class="live-price">BTC live: <strong>${escapeHtml(formatted)}</strong>${updatedText}</p>`;
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

function resolveLogo(meta) {
  const logo = normalizeText(meta.logo_url);
  if (!logo) {
    return "/brand_orange_bg_transparent@2xSite.svg";
  }
  return logo;
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

function formatCurrency(value, currency) {
  const maybeCurrency = normalizeText(currency) || "USD";
  const maximumFractionDigits = Math.abs(value) >= 1000 ? 0 : 2;
  try {
    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: maybeCurrency,
      maximumFractionDigits,
    }).format(value);
  } catch {
    return `$${Number(value).toLocaleString("en-US")}`;
  }
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
