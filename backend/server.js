const express = require("express");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT_DIR = path.resolve(__dirname, "..");
const FRONTEND_DIR = path.join(ROOT_DIR, "frontend");
const UPLOAD_DIR = path.join(ROOT_DIR, "uploads");
const LATEST_FILE = path.join(UPLOAD_DIR, "blinkit-master.xlsx");

fs.mkdirSync(UPLOAD_DIR, { recursive: true });

app.use(cors());
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true }));
app.use("/uploads", express.static(UPLOAD_DIR));

const upload = multer({ dest: UPLOAD_DIR });
const sessions = new Map();

const FIELD_ALIASES = {
  status: ["status"],
  weightKg: ["product weight kg", "product weight (kg)", "weight"],
  itemId: ["item id"],
  asin: ["asin"],
  sku: ["sku", "sku code", "item code", "product code", "article code"],
  model: ["model", "model no", "model number", "model name"],
  productName: ["product name", "item name", "title", "product title", "name", "description"],
  brand: ["brand", "brand name"],
  expansionLevel: ["expansion level"],
  category: ["category", "category l1", "l1 category", "vertical"],
  subCategory: ["sub category", "category l2", "category level 2", "subcategory"],
  variant: ["variant", "size", "pack size", "color"],
  hsn: ["hsn", "hsn code"],
  mrp: ["mrp", "list price"],
  blinkitSp: ["blinkit pricing", "blinkit sp", "blinkit selling price", "selling price", "sale price", "sp", "our price"],
  dealPrice: ["deal price", "deal sp", "offer price"],
  bauPrice: ["amazon bau", "bau price", "bau sp", "regular price", "base price"],
  dpValue: ["dp"],
  nlcValue: ["nlc"],
  adCostValue: ["ad cost"],
  commissionValue: ["comission", "commission"],
  fulfillmentValue: ["fulfillment"],
  returnValue: ["return"],
  storage2Value: ["storage 2 m", "storage (2 m)"],
  storage3Value: ["storage 3 m", "storage (3 m)"],
  shippingValue: ["shipping 15 kg", "shipping (15/kg)", "shipping"],
  inwardingValue: ["inwarding"],
  ohValue: ["oh"],
  financeValue: ["finance"],
  gstValue: ["gst"],
  totalCostValue: ["total cost"],
  marginValue: ["margin"],
  marginPctValue: ["margin %"],
  commissionPctValue: ["comission %", "commission %"],
  quantityValue: ["quantity"],
  valueValue: ["value"],
  marginValueTotal: ["margin value"],
  remark: ["remark"],
  purchaseCost: ["purchase cost", "buying price", "buy price", "landed cost", "cogs", "procurement cost", "dp"],
  inwardCost: ["inward cost", "inwarding cost", "freight", "freight cost", "logistics inward", "freight and handling"],
  packagingCost: ["packaging cost", "packaging", "packing cost", "pack cost"],
  warehousingCost: ["warehouse cost", "warehousing cost", "storage cost"],
  shippingFee: ["shipping fee", "last mile", "delivery fee", "shipping cost"],
  pickPackFee: ["pick pack fee", "pick and pack", "pick-pack fee", "fulfillment fee", "handling fee"],
  fixedFee: ["fixed fee", "platform fee", "closing fee"],
  otherCost: ["other cost", "misc cost", "additional cost", "other charges"],
  commissionPct: ["commission %", "commission pct", "referral fee %", "referral %", "trade margin %"],
  collectionPct: ["collection %", "payment gateway %", "collection fee %"],
  adsPct: ["ads %", "advertising %", "marketing %", "ads spend %"],
  returnsPct: ["returns %", "return loss %", "returns allowance %"],
  promoPct: ["promo %", "discount funding %", "promotion %"],
  promoFlat: ["promo flat", "discount funding", "promo amount"],
  gstPct: ["gst %", "gst", "tax %", "gst slab"],
};

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/%/g, " pct ")
    .replace(/[₹$()\-_/.,]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKey(value) {
  return normalizeHeader(value).replace(/\s+/g, "");
}

function isBlank(value) {
  return value === undefined || value === null || String(value).trim() === "";
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  if (isBlank(value)) {
    return 0;
  }
  const cleaned = String(value).replace(/,/g, "").replace(/[₹$%]/g, "").trim();
  const parsed = parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function toText(value) {
  return isBlank(value) ? "" : String(value).trim();
}

function getFirstValue(row, aliases) {
  const rowKeys = Object.keys(row);
  const normalizedRow = new Map(rowKeys.map((key) => [normalizeKey(key), key]));
  for (const alias of aliases) {
    const exactKey = normalizedRow.get(normalizeKey(alias));
    if (exactKey && !isBlank(row[exactKey])) {
      return { value: row[exactKey], source: exactKey };
    }
  }
  return { value: "", source: "" };
}

function getCellMeta(sheet, rowNumber, column) {
  const cell = sheet[`${column}${rowNumber}`];
  return {
    formula: cell && cell.f ? cell.f : "",
    value: cell ? cell.v : "",
  };
}

function deriveRow(rawRow, index, sheet, rowNumber) {
  const matchedColumns = {};

  function readText(field) {
    const picked = getFirstValue(rawRow, FIELD_ALIASES[field] || []);
    matchedColumns[field] = picked.source;
    return toText(picked.value);
  }

  function readNumber(field) {
    const picked = getFirstValue(rawRow, FIELD_ALIASES[field] || []);
    matchedColumns[field] = picked.source;
    return toNumber(picked.value);
  }

  const sku = readText("sku");
  const model = readText("model");
  const productName = readText("productName");

  return {
    id: `row-${index + 1}`,
    rowNumber,
    status: readText("status"),
    itemId: readText("itemId"),
    asin: readText("asin"),
    sku: sku || `SKU-${index + 1}`,
    model: model || sku || `MODEL-${index + 1}`,
    productName: productName || sku || `Product ${index + 1}`,
    brand: readText("brand"),
    expansionLevel: readText("expansionLevel"),
    category: readText("category"),
    subCategory: readText("subCategory"),
    variant: readText("variant"),
    hsn: readText("hsn"),
    weightKg: readNumber("weightKg"),
    dpValue: readNumber("dpValue"),
    nlcValue: readNumber("nlcValue"),
    adCostValue: readNumber("adCostValue"),
    commissionValue: readNumber("commissionValue"),
    fulfillmentValue: readNumber("fulfillmentValue"),
    returnValue: readNumber("returnValue"),
    storage2Value: readNumber("storage2Value"),
    storage3Value: readNumber("storage3Value"),
    shippingValue: readNumber("shippingValue"),
    inwardingValue: readNumber("inwardingValue"),
    ohValue: readNumber("ohValue"),
    financeValue: readNumber("financeValue"),
    gstValue: readNumber("gstValue"),
    totalCostValue: readNumber("totalCostValue"),
    marginValue: readNumber("marginValue"),
    marginPctValue: readNumber("marginPctValue"),
    commissionPctValue: readNumber("commissionPctValue"),
    mrp: readNumber("mrp"),
    blinkitSp: readNumber("blinkitSp"),
    dealPrice: readNumber("dealPrice"),
    bauPrice: readNumber("bauPrice"),
    amazonBau: readNumber("bauPrice"),
    quantityValue: readNumber("quantityValue"),
    valueValue: readNumber("valueValue"),
    marginValueTotal: readNumber("marginValueTotal"),
    purchaseCost: readNumber("purchaseCost"),
    inwardCost: readNumber("inwardCost"),
    packagingCost: readNumber("packagingCost"),
    warehousingCost: readNumber("warehousingCost"),
    shippingFee: readNumber("shippingFee"),
    pickPackFee: readNumber("pickPackFee"),
    fixedFee: readNumber("fixedFee"),
    otherCost: readNumber("otherCost"),
    commissionPct: readNumber("commissionPct"),
    collectionPct: readNumber("collectionPct"),
    adsPct: readNumber("adsPct"),
    returnsPct: readNumber("returnsPct"),
    promoPct: readNumber("promoPct"),
    promoFlat: readNumber("promoFlat"),
    gstPct: readNumber("gstPct"),
    remark: readText("remark"),
    formulas: {
      adCost: getCellMeta(sheet, rowNumber, "Q").formula,
      commission: getCellMeta(sheet, rowNumber, "R").formula,
      fulfillment: getCellMeta(sheet, rowNumber, "S").formula,
      returnCost: getCellMeta(sheet, rowNumber, "T").formula,
      storage2: getCellMeta(sheet, rowNumber, "U").formula,
      storage3: getCellMeta(sheet, rowNumber, "V").formula,
      shipping: getCellMeta(sheet, rowNumber, "W").formula,
      inwarding: getCellMeta(sheet, rowNumber, "X").formula,
      oh: getCellMeta(sheet, rowNumber, "Y").formula,
      finance: getCellMeta(sheet, rowNumber, "Z").formula,
      gst: getCellMeta(sheet, rowNumber, "AA").formula,
      totalCost: getCellMeta(sheet, rowNumber, "AB").formula,
      margin: getCellMeta(sheet, rowNumber, "AC").formula,
      marginPct: getCellMeta(sheet, rowNumber, "AD").formula,
      quantity: getCellMeta(sheet, rowNumber, "AH").formula,
      value: getCellMeta(sheet, rowNumber, "AI").formula,
      marginValue: getCellMeta(sheet, rowNumber, "AJ").formula,
    },
    matchedColumns,
    raw: rawRow,
  };
}

function parseWorkbook(filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: false, cellFormula: true });
  const sheetName = workbook.SheetNames.includes("Blinkit") ? "Blinkit" : workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const normalizedRows = rows
    .map((row, index) => deriveRow(row, index, sheet, index + 2))
    .filter((row) => row.sku || row.model || row.productName);

  return {
    workbook,
    sheetName,
    rows: normalizedRows,
    columns: rows[0] ? Object.keys(rows[0]) : [],
  };
}

function calculateMetrics(row, globals) {
  const sellingPrice = toNumber(row.blinkitSp || row.dealPrice || row.bauPrice || row.mrp);
  const gstPct = toNumber(row.gstPct || globals.gstDefaultPct);
  const procurement =
    toNumber(row.purchaseCost) +
    toNumber(row.inwardCost || globals.inwardDefault) +
    toNumber(row.packagingCost || globals.packagingDefault) +
    toNumber(row.warehousingCost) +
    toNumber(row.otherCost);

  const flatFees =
    toNumber(row.shippingFee) +
    toNumber(row.pickPackFee) +
    toNumber(row.fixedFee) +
    toNumber(row.promoFlat);

  const variablePct =
    toNumber(row.commissionPct) / 100 +
    toNumber(row.collectionPct || globals.collectionPct) / 100 +
    toNumber(row.adsPct || globals.adsDefaultPct) / 100 +
    toNumber(row.returnsPct || globals.returnsDefaultPct) / 100 +
    toNumber(row.promoPct) / 100;

  const overhead = procurement * (toNumber(globals.overheadPct) / 100);
  const finance = procurement * (toNumber(globals.financePct) / 100);
  const gstAmount = gstPct > 0 ? sellingPrice * (gstPct / (100 + gstPct)) : 0;
  const variableFees = sellingPrice * variablePct;
  const totalCost = procurement + flatFees + variableFees + overhead + finance + gstAmount;
  const revenueExGst = sellingPrice - gstAmount;
  const grossMargin = revenueExGst - procurement;
  const netMargin = sellingPrice - totalCost;
  const grossMarginPct = sellingPrice > 0 ? (grossMargin / sellingPrice) * 100 : 0;
  const netMarginPct = sellingPrice > 0 ? (netMargin / sellingPrice) * 100 : 0;

  const contributionFactor = 1 - variablePct - (gstPct > 0 ? gstPct / (100 + gstPct) : 0);
  const fixedBase = procurement + flatFees + overhead + finance;
  const breakEvenSp = contributionFactor > 0 ? fixedBase / contributionFactor : 0;
  const targetFactor = contributionFactor - toNumber(globals.targetMarginPct) / 100;
  const suggestedSp = targetFactor > 0 ? fixedBase / targetFactor : 0;

  return {
    sellingPrice,
    gstPct,
    procurement,
    flatFees,
    variableFees,
    overhead,
    finance,
    gstAmount,
    revenueExGst,
    grossMargin,
    grossMarginPct,
    netMargin,
    netMarginPct,
    breakEvenSp,
    suggestedSp,
    costRows: [
      ["Procurement", "Purchase + inward + packaging + warehousing + other", procurement],
      ["Marketplace Fees", "Commission + collection + shipping + pick-pack + fixed", flatFees + sellingPrice * ((toNumber(row.commissionPct) + toNumber(row.collectionPct || globals.collectionPct)) / 100)],
      ["Growth Spend", "Ads + returns + promos", sellingPrice * ((toNumber(row.adsPct || globals.adsDefaultPct) + toNumber(row.returnsPct || globals.returnsDefaultPct) + toNumber(row.promoPct)) / 100) + toNumber(row.promoFlat)],
      ["Taxes", "GST on selling price", gstAmount],
      ["Overhead & Finance", "Business overhead + finance cost", overhead + finance],
      ["Grand Total", "All above", totalCost],
    ],
  };
}

function makeSessionPayload(sessionId, parsed, fileName) {
  return {
    sessionId,
    fileName,
    sheetName: parsed.sheetName,
    rowCount: parsed.rows.length,
    columns: parsed.columns,
    rows: parsed.rows,
  };
}

function createSessionFromFile(filePath, originalName = path.basename(filePath)) {
  const parsed = parseWorkbook(filePath);
  const sessionId = crypto.randomUUID();
  const payload = makeSessionPayload(sessionId, parsed, originalName);
  sessions.set(sessionId, payload);
  return payload;
}

function getGlobalsFromRequest(query) {
  return {
    gstDefaultPct: toNumber(query.gst_default_pct || 18),
    adsDefaultPct: toNumber(query.ads_default_pct || 3),
    returnsDefaultPct: toNumber(query.returns_default_pct || 1),
    overheadPct: toNumber(query.overhead_pct || 2),
    financePct: toNumber(query.finance_pct || 1.5),
    inwardDefault: toNumber(query.inward_default || 0),
    packagingDefault: toNumber(query.packaging_default || 0),
    collectionPct: toNumber(query.collection_pct || 0),
    targetMarginPct: toNumber(query.target_margin_pct || 20),
  };
}

function getToolRulesFromRequest(query) {
  return {
    adPct: toNumber(query.rule_ad_pct || 10),
    returnPct: toNumber(query.rule_return_pct || 4),
    ohPct: toNumber(query.rule_oh_pct || 5),
    financePct: toNumber(query.rule_finance_pct || 3),
    gstPct: toNumber(query.rule_gst_pct || 18),
    shippingRate: toNumber(query.rule_shipping_rate || 15),
    targetMarginPct: toNumber(query.rule_target_margin_pct || 10),
    fulfillmentFixed: toNumber(query.rule_fulfillment_default || 50),
    inwardingFixed: toNumber(query.rule_inwarding_default || 5),
  };
}

function calculateToolExportRow(row, rules) {
  const finalBlinkitPrice = toNumber(row.blinkitSp);
  const productWeightKg = toNumber(row.weightKg || row.raw?.["Product Weight (Kg)"]);
  const dp = toNumber(row.dpValue);
  const nlc = toNumber(row.nlcValue);
  const storage2 = toNumber(row.storage2Value);
  const commissionValue = toNumber(row.commissionValue);
  const commissionPct = toNumber(row.commissionPctValue) <= 1
    ? toNumber(row.commissionPctValue) * 100
    : toNumber(row.commissionPctValue);
  const fulfillment = toNumber(row.fulfillmentValue || rules.fulfillmentFixed);
  const inwarding = toNumber(row.inwardingValue || rules.inwardingFixed);

  const adCost = finalBlinkitPrice * rules.adPct / 100;
  const returnCost = finalBlinkitPrice * rules.returnPct / 100;
  const shipping = productWeightKg * rules.shippingRate;
  const oh = finalBlinkitPrice * rules.ohPct / 100;
  const finance = finalBlinkitPrice * rules.financePct / 100;
  const gst = rules.gstPct > 0 ? finalBlinkitPrice - (finalBlinkitPrice / (1 + rules.gstPct / 100)) : 0;
  const totalCost = dp + adCost + commissionValue + fulfillment + returnCost + storage2 + shipping + inwarding + oh + finance + gst;
  const margin = finalBlinkitPrice - totalCost;
  const marginPct = finalBlinkitPrice > 0 ? (margin / finalBlinkitPrice) * 100 : 0;
  const variablePct = (rules.adPct + rules.returnPct + rules.ohPct + rules.financePct) / 100;
  const gstFactor = rules.gstPct > 0 ? (rules.gstPct / (100 + rules.gstPct)) : 0;
  const contribution = 1 - variablePct - gstFactor;
  const fixedBase = dp + commissionValue + fulfillment + storage2 + shipping + inwarding;
  const suggestedBlinkitSellingPrice = contribution > 0 ? fixedBase / contribution : 0;
  const targetContribution = contribution - (rules.targetMarginPct / 100);
  const targetMarginSuggestedPrice = targetContribution > 0 ? fixedBase / targetContribution : 0;

  return {
    Brand: row.brand,
    Model: row.model,
    SKU: row.sku,
    "Expansion Level": row.expansionLevel,
    "Item ID": row.itemId,
    ASIN: row.asin,
    "Category L1": row.category,
    "Final Blinkit Price": finalBlinkitPrice,
    "Product Weight (Kg)": productWeightKg,
    DP: dp,
    NLC: nlc,
    "Storage (2 M)": storage2,
    "Commission Value": commissionValue,
    "Commission %": commissionPct,
    Fulfillment: fulfillment,
    Inwarding: inwarding,
    "Ad Cost % Rule": rules.adPct,
    "Return % Rule": rules.returnPct,
    "OH % Rule": rules.ohPct,
    "Finance % Rule": rules.financePct,
    "GST % Rule": rules.gstPct,
    "Shipping Rate / KG": rules.shippingRate,
    "Ad Cost": adCost,
    Return: returnCost,
    "Shipping (15/KG)": shipping,
    OH: oh,
    Finance: finance,
    GST: gst,
    "Total Cost": totalCost,
    Margin: margin,
    "Margin %": marginPct,
    "Suggested Blinkit Selling Price": margin < 0 ? suggestedBlinkitSellingPrice : 0,
    "Suggested Price At Target Margin": targetMarginSuggestedPrice,
  };
}

function ensureLatestSession() {
  if (!fs.existsSync(LATEST_FILE)) {
    return null;
  }
  const existing = [...sessions.values()].find((session) => session.fileName === path.basename(LATEST_FILE));
  if (existing) {
    return existing;
  }
  return createSessionFromFile(LATEST_FILE, path.basename(LATEST_FILE));
}

app.get("/", (_req, res) => {
  res.redirect("/blinkit");
});

app.get("/blinkit", (_req, res) => {
  res.sendFile(path.join(FRONTEND_DIR, "blinkit.html"));
});

app.get("/api/status", (_req, res) => {
  const session = ensureLatestSession();
  res.json({
    hasLatest: Boolean(session),
    session: session || null,
  });
});

app.post("/api/upload", upload.single("file"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No file uploaded" });
  }

  try {
    fs.copyFileSync(req.file.path, LATEST_FILE);
    try {
      fs.unlinkSync(req.file.path);
    } catch (_error) {
      // OneDrive/Windows can hold the temp file briefly; keeping the temp file is safe here.
    }
    const session = createSessionFromFile(LATEST_FILE, req.file.originalname || "blinkit-master.xlsx");
    return res.json({ ok: true, session });
  } catch (error) {
    return res.status(500).json({ error: error.message || "Failed to parse workbook" });
  }
});

app.get("/api/session/:sessionId", (req, res) => {
  const session = sessions.get(req.params.sessionId);
  if (!session) {
    return res.status(404).json({ error: "Session not found" });
  }
  return res.json(session);
});

app.get("/api/export", (req, res) => {
  const session = sessions.get(req.query.session_id);
  if (!session) {
    return res.status(404).json({ error: "Session not found" });
  }

  const rules = getToolRulesFromRequest(req.query);
  const exportRows = session.rows.map((row) => calculateToolExportRow(row, rules));

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(exportRows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Blinkit Margin");

  const exportPath = path.join(UPLOAD_DIR, "blinkit-margin-export.xlsx");
  XLSX.writeFile(workbook, exportPath);
  return res.download(exportPath, "blinkit-margin-export.xlsx");
});

app.get("/api/health", (_req, res) => {
  res.json({ ok: true, timestamp: new Date().toISOString() });
});

ensureLatestSession();

app.listen(PORT, () => {
  console.log(`Blinkit margin calculator running at http://localhost:${PORT}`);
});
