function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GW2')
    .addItem('Update TP Prices', 'updateTradingPostPrices')
    .addItem('Update Profit Costs', 'refreshCosts')
    .addItem('Fetch Missing Item IDs', 'fetchMissingItemIDs')
    .addToUi();
}

/**
 * Convert a copper value (integer) to a human-readable "Xg Ys Zc" string.
 * Examples:
 *  12345 -> "1g 23s 45c"
 *  250   -> "2s 50c"
 *  7     -> "7c"
 *  0     -> "0c"
 */
function formatCopperToGSC(value) {
  if (value === '' || value === null || value === undefined) return '';

  var copper = Number(value);
  if (isNaN(copper)) return '';

  // HACK: for now just return 2.50 for 250 as the g/s/c is not as readable
  return copper/100;

  var gold = Math.floor(copper / 10000);
  var silver = Math.floor((copper % 10000) / 100);
  var copperRemainder = copper % 100;

  var parts = [];
  if (gold) parts.push(gold + 'g');
  if (silver) parts.push(silver + 's');

  // Show copper if non-zero OR everything else is zero
  if (copperRemainder || (!gold && !silver)) {
    parts.push(copperRemainder + 'c');
  }

  return parts.join(' ');
}

/**
 * Update TP prices for multiple sheets.
 * All target sheets must have columns: A=Name, B=Buy, C=Sell.
 * IDs are looked up from the "Items" sheet (Name + ID).
 */
function updateTradingPostPrices() {
  const ss = SpreadsheetApp.getActive();
  const items = ss.getSheetByName('Items');
  const pricesSheet = ss.getSheetByName('Prices');
  if (!items || !pricesSheet) throw new Error('Missing "Items" or "Prices" sheet');

  const norm = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim().toLowerCase();

  // --- Load Items meta: A name, B id, C mode, D manual cost ---
  const itemsLast = items.getLastRow();
  const itemsData = itemsLast >= 2 ? items.getRange(2, 1, itemsLast - 1, 4).getValues() : [];
  const nameToId = {};
  const vendorCost = {}; // key -> number|null

  itemsData.forEach(r => {
    const name = r[0];
    if (!name) return;
    const key = norm(name);
    const id = r[1];
    const mode = String(r[2] || "").toUpperCase().trim();
    const manual = (r[3] !== "" && r[3] != null) ? Number(r[3]) : null;

    if (id) nameToId[key] = String(id);
    if (mode === "VENDOR") vendorCost[key] = manual; // may be null
  });

  // --- Load Prices names + existing buy/sell ---
  const firstDataRow = 2;
  const lastRow = pricesSheet.getLastRow();
  if (lastRow < firstDataRow) return;

  const numRows = lastRow - firstDataRow + 1;
  const names = pricesSheet.getRange(firstDataRow, 1, numRows, 1).getValues();    // col A
  const existing = pricesSheet.getRange(firstDataRow, 2, numRows, 2).getValues(); // cols B:C

  // Prepare output = start with existing values, then overwrite where we have new info
  const out = existing.map(r => [r[0], r[1]]);

  // Build ID -> rowIndexes (0-based within out)
  const idToIdxs = {};
  const ids = [];
  const idSet = new Set();

  names.forEach((row, i) => {
    const nameCell = row[0];
    if (!nameCell) return;
    const key = norm(nameCell);

    // VENDOR override
    if (Object.prototype.hasOwnProperty.call(vendorCost, key)) {
      const manual = vendorCost[key];
      if (manual != null && !Number.isNaN(manual)) {
        out[i][0] = manual; // buy
        out[i][1] = manual; // sell
      } else {
        // If vendor item has no manual cost, you can blank it or leave existing:
        // out[i][0] = ""; out[i][1] = "";
      }
      return;
    }

    const id = nameToId[key];
    if (!id) return;

    if (!idToIdxs[id]) idToIdxs[id] = [];
    idToIdxs[id].push(i);

    if (!idSet.has(id)) {
      idSet.add(id);
      ids.push(id);
    }
  });

  if (ids.length === 0) {
    pricesSheet.getRange(firstDataRow, 2, numRows, 2).setValues(out);
    return;
  }

  // --- Fetch TP prices in chunks ---
  const chunkSize = 200; // bump to 200
  const urlBase = "https://api.guildwars2.com/v2/commerce/prices?ids=";
  const returned = new Set();

  for (let start = 0; start < ids.length; start += chunkSize) {
    const chunk = ids.slice(start, start + chunkSize);
    const url = urlBase + encodeURIComponent(chunk.join(","));

    let resp, code, data;
    try {
      resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      code = resp.getResponseCode();
      if (code !== 200 && code !== 206) continue;
      data = JSON.parse(resp.getContentText());
    } catch (e) {
      continue;
    }

    data.forEach(entry => {
      const id = String(entry.id);
      returned.add(id);
      const idxs = idToIdxs[id];
      if (!idxs) return;

      const buyCopper  = entry.buys  && entry.buys.unit_price  ? entry.buys.unit_price  : null; // buy order
      const sellCopper = entry.sells && entry.sells.unit_price ? entry.sells.unit_price : null; // sell listing

      const buyVal  = buyCopper  != null ? formatCopperToGSC(buyCopper)  : "";
      const sellVal = sellCopper != null ? formatCopperToGSC(sellCopper) : "";

      idxs.forEach(i => {
        out[i][0] = buyVal;
        out[i][1] = sellVal;
      });
    });

    // Optional: smaller sleep or none; API is usually fine without it
    // Utilities.sleep(100);
  }

  // Optional: if an ID wasn’t returned (non-tradable), blank it instead of leaving old values:
  // ids.forEach(id => {
  //   if (returned.has(id)) return;
  //   (idToIdxs[id] || []).forEach(i => { out[i][0] = ""; out[i][1] = ""; });
  // });

  // --- One single write back ---
  pricesSheet.getRange(firstDataRow, 2, numRows, 2).setValues(out);
}



/**
 * Fetch IDs for any item names in the "Items" sheet that do not have IDs yet.
 * Uses GW2TP's bulk items-names JSON as a name -> ID dictionary.
 *
 * Source: http://api.gw2tp.com/1/bulk/items-names.json
 * Format: { "items": [ [id, "Name"], ... ] }
 */
function fetchMissingItemIDs() {
  const sheetName = "Items";  // your sheet with Name + ID
  const nameCol = 1;          // Column A (Name)
  const idCol = 2;            // Column B (ID)

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet "' + sheetName + '" not found.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // nothing but header

  // --- 1) Get all item names from your sheet ---
  const namesRange = sheet.getRange(2, nameCol, lastRow - 1, 1);
  const idsRange   = sheet.getRange(2, idCol,   lastRow - 1, 1);
  const names = namesRange.getValues(); // [[Name], [Name], ...]
  const ids   = idsRange.getValues();   // [[ID],   [ID],   ...]

  // --- 2) Download GW2TP name -> ID data ---
  const url = "http://api.gw2tp.com/1/bulk/items-names.json";
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error("Failed to download items-names.json from GW2TP (HTTP " +
                    response.getResponseCode() + ")");
  }

  const payload = JSON.parse(response.getContentText());
  const items = payload.items; // [ [id, "Name"], ... ]

  // --- 3) Build a lookup: normalized name -> [id, id, ...] ---
  const nameToIds = {}; // { "iron ore": [19697, ...], ... }

  items.forEach(function (entry) {
    const id   = entry[0];
    const name = entry[1];
    if (!name || !id) return;

    const key = String(name).trim().toLowerCase();
    if (!nameToIds[key]) {
      nameToIds[key] = [];
    }
    nameToIds[key].push(id);
  });

  // --- 4) For each row in your Items sheet, fill missing IDs ---
  const updates = [];

  for (let i = 0; i < names.length; i++) {
    const nameCell = names[i][0];
    const existingId = ids[i][0];

    if (!nameCell || existingId) {
      // skip blank names and already-filled IDs
      updates.push([existingId]); // keep whatever was there
      continue;
    }

    const key = String(nameCell).trim().toLowerCase();
    const candidateIds = nameToIds[key];

    if (candidateIds && candidateIds.length > 0) {
      // Take the first matching ID.
      // (There can be duplicates like multiple "Country Coat" variants.)
      updates.push([candidateIds[0]]);
    } else {
      // Couldn’t find a match; mark it so you know to fix manually.
      updates.push(["NOT FOUND"]);
    }
  }

  // --- 5) Write all updated IDs back in one go ---
  idsRange.setValues(updates);
}

function refreshCosts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Refreshing...");
  const items = ss.getSheetByName("Items");
  if (!items) throw new Error('Sheet "Items" not found');
  items.getRange("Z1").setValue(new Date()); // bump refresh token
}
/***************
 * GW2 Crafting Cost Engine (optimized)
 *
 * Sheets expected:
 *  Prices:  A ItemName | B BuyOrder (number) | C SellOrder (optional)
 *  Recipes: A OutputItem | B OutputQty | C IngredientItem | D IngredientQty
 *  Items:   A ItemName | B ItemID (optional) | C Mode | D ManualCost
 *
 * Modes (Items!C):
 *  ""      = normal
 *  VENDOR  = must use ManualCost (Items!D) for that item
 *  ALT     = allow chain to continue even if no TP price (uses craft if exists else 0)
 *  BLOCK   = break chain (unpriceable)
 *  CRAFT   = must craft this item if used as an ingredient (ignore TP even if it exists)
 *
 * Custom functions:
 *  =CRAFTCOST(A3:A500)       -> STRICT craft-only cost for the item (top-level must have recipe) using EFFECTIVE cost for ingredients
 *  =CRAFTCOST_EFF(A3:A500)   -> EFFECTIVE cost for any item (min(TP buy, craft) with your mode rules)
 *
 * Notes:
 * - CRAFTCOST returns "N/A" if the item has no recipe (prevents fake profit from buy->sell spread)
 * - Both functions are vectorized: pass a range, get a column back.
 ***************/

const _GW2_COST_ENGINE_CACHE_KEY = "GW2_COST_ENGINE_V1";
let _engine = null;

function CRAFTCOST(items) {
  const eng = getEngine_();
  return applyToInput_(items, (name) => eng.craftStrict(name));
}

function CRAFTCOST_EFF(items) {
  const eng = getEngine_();
  return applyToInput_(items, (name) => eng.effective(name));
}

/** Optional: manual cache reset if you changed Recipes/Items/Prices */
function CLEAR_CRAFT_CACHE() {
  CacheService.getScriptCache().remove(_GW2_COST_ENGINE_CACHE_KEY);
  _engine = null;
}

/* ------------------------- internals ------------------------- */

function getEngine_() {
  if (_engine) return _engine;

  const cache = CacheService.getScriptCache();
  const cached = cache.get(_GW2_COST_ENGINE_CACHE_KEY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      _engine = buildEngineFromData_(parsed);
      return _engine;
    } catch (e) {}
  }

  const ss = SpreadsheetApp.getActive();
  const pricesSheet = ss.getSheetByName("Prices");
  const recipesSheet = ss.getSheetByName("Recipes");
  const itemsSheet = ss.getSheetByName("Items");
  if (!pricesSheet || !recipesSheet || !itemsSheet) {
    throw new Error('Missing sheet: need "Prices", "Recipes", and "Items".');
  }

  const pricesData = readUsed_(pricesSheet, 3);   // A:C
  const recipesData = readUsed_(recipesSheet, 4); // A:D
  const itemsData = readUsed_(itemsSheet, 4);     // A:D

  const data = { pricesData, recipesData, itemsData };
  _engine = buildEngineFromData_(data);

  cache.put(_GW2_COST_ENGINE_CACHE_KEY, JSON.stringify(data), 300);
  return _engine;
}

function buildEngineFromData_(data) {
  const norm = (v) =>
    String(v ?? "")
      .replace(/\u00A0/g, " ")
      .trim()
      .toLowerCase();

  // ---- Prices map: name -> buyOrder ----
  const priceMap = new Map();
  for (let i = 1; i < data.pricesData.length; i++) {
    const row = data.pricesData[i];
    const k = norm(row[0]);
    if (!k) continue;
    const buy = Number(row[1]);
    priceMap.set(k, isFinite(buy) ? buy : 0);
  }

  // ---- Items meta: name -> {mode, cost} ----
  const metaMap = new Map();
  for (let i = 1; i < data.itemsData.length; i++) {
    const row = data.itemsData[i];
    const k = norm(row[0]);
    if (!k) continue;
    const mode = String(row[2] || "").toUpperCase().trim();
    const cost =
      row[3] !== "" && row[3] != null && isFinite(Number(row[3]))
        ? Number(row[3])
        : null;
    metaMap.set(k, { mode, cost });
  }

  // ---- Recipes: output -> {outQty, ingredients:[{key, qty}...]} ----
  const recipesByOutput = new Map();
  for (let i = 1; i < data.recipesData.length; i++) {
    const row = data.recipesData[i];
    const outKey = norm(row[0]);
    if (!outKey) continue;
    const outQty = Number(row[1]) || 1;
    const ingKey = norm(row[2]);
    const ingQty = Number(row[3]) || 0;
    if (!ingKey) continue;

    if (!recipesByOutput.has(outKey)) {
      recipesByOutput.set(outKey, { outQty: outQty, ingredients: [] });
    }
    recipesByOutput.get(outKey).ingredients.push({ key: ingKey, qty: ingQty });
  }

  const memoEff = new Map();
  const memoStrict = new Map();

  function effectiveCost_(key, stack) {
    if (!key) return "";
    if (memoEff.has(key)) return memoEff.get(key);

    stack = stack || [];
    if (stack.includes(key)) return memoEff.set(key, "").get(key);

    const meta = metaMap.get(key) || null;
    const mode = meta ? meta.mode : "";
    const tpBuy = priceMap.has(key) ? priceMap.get(key) : 0;
    const hasTP = tpBuy > 0;
    const recipe = recipesByOutput.get(key) || null;

    // BLOCK always breaks
    if (mode === "BLOCK") {
      memoEff.set(key, "");
      return "";
    }

    // VENDOR requires manual cost
    if (mode === "VENDOR") {
      const v = meta ? meta.cost : null;
      memoEff.set(key, v == null ? "" : v);
      return memoEff.get(key);
    }

    // Manual override (any mode except BLOCK/VENDOR already handled)
    if (meta && meta.cost != null) {
      memoEff.set(key, meta.cost);
      return meta.cost;
    }

    // Compute craft cost if recipe exists (ingredients use EFFECTIVE)
    let craft = null;
    if (recipe) {
      let total = 0;
      for (const ing of recipe.ingredients) {
        const c = effectiveCost_(ing.key, stack.concat(key));
        if (c === "" || c == null) {
          total = null;
          break;
        }
        total += c * ing.qty;
      }
      if (total != null) craft = total / (recipe.outQty || 1);
    }

    // ALT: keep chain alive even if missing TP and missing recipe
    if (mode === "ALT") {
      const v = craft != null ? craft : 0;
      memoEff.set(key, v);
      return v;
    }

    // CRAFT: MUST craft when used as ingredient (ignore TP)
    if (mode === "CRAFT") {
      // If there's a recipe, use it; if not, break/flag
      if (craft != null) {
        memoEff.set(key, craft);
        return craft;
      }
      // If you prefer "N/A" instead of blank here, change "" to "N/A"
      memoEff.set(key, "");
      return memoEff.get(key);
    }

    // Normal: min(TP buy, craft) when both exist; else whichever exists
    let final = "";
    if (craft != null && hasTP) final = Math.min(tpBuy, craft);
    else if (craft != null) final = craft;
    else if (hasTP) final = tpBuy;
    else final = "";

    memoEff.set(key, final);
    return final;
  }

  // STRICT craft-only for the *output item* (must have recipe), but ingredients use EFFECTIVE
  function craftStrict_(key, stack) {
    if (!key) return "";
    if (memoStrict.has(key)) return memoStrict.get(key);

    stack = stack || [];
    if (stack.includes(key)) return memoStrict.set(key, "").get(key);

    const meta = metaMap.get(key) || null;
    const mode = meta ? meta.mode : "";
    const recipe = recipesByOutput.get(key) || null;

    if (mode === "BLOCK") {
      memoStrict.set(key, "");
      return "";
    }

    if (mode === "VENDOR") {
      const v = meta ? meta.cost : null;
      memoStrict.set(key, v == null ? "" : v);
      return memoStrict.get(key);
    }

    // If no recipe for the output: return N/A
    if (!recipe) {
      memoStrict.set(key, "N/A");
      return "N/A";
    }

    let total = 0;
    for (const ing of recipe.ingredients) {
      const c = effectiveCost_(ing.key, stack.concat(key));
      if (c === "" || c == null) {
        memoStrict.set(key, "");
        return "";
      }
      total += c * ing.qty;
    }

    const perUnit = total / (recipe.outQty || 1);
    memoStrict.set(key, perUnit);
    return perUnit;
  }

  return {
    normName: norm,
    effective: (name) => effectiveCost_(norm(name), []),
    craftStrict: (name) => craftStrict_(norm(name), []),
  };
}

function readUsed_(sheet, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [[""]];
  const width = Math.max(1, numCols);
  return sheet.getRange(1, 1, lastRow, width).getValues();
}

function applyToInput_(items, fn) {
  if (items == null) return [[""]];
  if (!Array.isArray(items)) return [[fn(items)]];
  return items.map((row) => {
    const name = row && row.length ? row[0] : "";
    if (!name) return [""];
    return [fn(name)];
  });
}
