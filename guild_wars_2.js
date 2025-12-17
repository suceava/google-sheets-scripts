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
  const mainSheetNames = ['Prices'];
  const rawSheetName = 'Items';

  const nameCol = 1; // A
  const buyCol  = 2; // B
  const sellCol = 3; // C

  const ss = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName(rawSheetName);
  if (!raw) throw new Error('Lookup sheet "' + rawSheetName + '" not found.');

  // ---- Build Name -> ID map + vendor metadata from Items sheet ----
  const rawLastRow = raw.getLastRow();
  if (rawLastRow < 2) return;

  // A: name, B: id, C: mode, D: manual cost
  const rawData = raw.getRange(2, 1, rawLastRow - 1, 4).getValues();
  const nameToId = {};
  const vendorNames = new Set();
  const vendorCostMap = {};

  rawData.forEach(row => {
    const rawName = row[0];
    const id = row[1];
    const mode = (row[2] || "").toString().toUpperCase(); // ALT / VENDOR / BLOCK / ""
    const manual = row[3];

    if (!rawName) return;
    const key = String(rawName).trim().toLowerCase();

    if (mode === 'VENDOR') {
      vendorNames.add(key);
      vendorCostMap[key] = (manual !== "" && manual != null) ? Number(manual) : null;
    }

    if (id) {
      nameToId[key] = id;
    }
  });

  mainSheetNames.forEach(sheetName => {
    const main = ss.getSheetByName(sheetName);
    if (!main) return;

    const firstDataRow = 2;
    const lastRow = main.getLastRow();
    if (lastRow < firstDataRow) return;

    const numRows = lastRow - firstDataRow + 1;
    const names = main.getRange(firstDataRow, nameCol, numRows, 1).getValues();

    const idToRows = {};
    const idSet = new Set();
    let missingIdCount = 0;

    names.forEach((row, i) => {
      const nameCell = row[0];
      if (!nameCell) return;

      const key = String(nameCell).trim().toLowerCase();
      const rowNum = firstDataRow + i;

      // 🔹 Handle VENDOR items: no API, just write manual cost
      if (vendorNames.has(key)) {
        const manual = vendorCostMap[key];
        if (manual != null) {
          main.getRange(rowNum, buyCol).setValue(manual);
          main.getRange(rowNum, sellCol).setValue(manual);
        } else {
          Logger.log('VENDOR item "' + nameCell + '" has no ManualCost in Items sheet.');
        }
        return; // do NOT add to idSet / API list
      }

      const id = nameToId[key];
      if (!id) {
        missingIdCount++;
        return; // leave Buy/Sell unchanged
      }

      const idStr = String(id);
      if (!idToRows[idStr]) idToRows[idStr] = [];
      idToRows[idStr].push(rowNum);
      idSet.add(idStr);
    });

    const apiIds = Array.from(idSet);
    if (apiIds.length === 0) {
      Logger.log('No IDs to update for sheet "' + sheetName + '". Missing-ID rows: ' + missingIdCount);
      return;
    }

    const chunkSize = 150;
    const urlBase = "https://api.guildwars2.com/v2/commerce/prices?ids=";
    const returnedIds = new Set();

    for (let start = 0; start < apiIds.length; start += chunkSize) {
      const chunk = apiIds.slice(start, start + chunkSize);
      const url = urlBase + encodeURIComponent(chunk.join(","));

      let response, code, data;
      try {
        response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        code = response.getResponseCode();
        if (code !== 200 && code !== 206) {
          Logger.log('GW2 API error ' + code + ' for URL: ' + url);
          continue;
        }
        data = JSON.parse(response.getContentText());
      } catch (e) {
        Logger.log('Fetch/parse failed for chunk starting at ' + start + ': ' + e);
        continue;
      }

      data.forEach(entry => {
        const idStr = String(entry.id);
        returnedIds.add(idStr);

        const rows = idToRows[idStr];
        if (!rows || rows.length === 0) return;

        const buyCopper  = entry.buys  && entry.buys.unit_price  ? entry.buys.unit_price  : '';
        const sellCopper = entry.sells && entry.sells.unit_price ? entry.sells.unit_price : '';

        const buyStr  = buyCopper  !== '' ? formatCopperToGSC(buyCopper)  : '';
        const sellStr = sellCopper !== '' ? formatCopperToGSC(sellCopper) : '';

        rows.forEach(r => {
          main.getRange(r, buyCol).setValue(buyStr);
          main.getRange(r, sellCol).setValue(sellStr);
        });
      });

      Utilities.sleep(300);
    }

    const notReturned = apiIds.filter(id => !returnedIds.has(id));
    if (notReturned.length > 0) {
      Logger.log(
        'Sheet "' + sheetName + '": ' +
        notReturned.length + ' IDs not returned by /commerce/prices (likely non-tradable). Example: ' +
        notReturned.slice(0, 10).join(', ')
      );
    }

    Logger.log(
      'Finished "' + sheetName + '". Updated unique IDs: ' + apiIds.length +
      '. Rows with names missing IDs: ' + missingIdCount +
      '.'
    );
  });
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

  // Try short-lived cache to avoid rereading sheets repeatedly during recalcs
  const cache = CacheService.getScriptCache();
  const cached = cache.get(_GW2_COST_ENGINE_CACHE_KEY);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      _engine = buildEngineFromData_(parsed);
      return _engine;
    } catch (e) {
      // fall through to rebuild
    }
  }

  const ss = SpreadsheetApp.getActive();
  const pricesSheet = ss.getSheetByName("Prices");
  const recipesSheet = ss.getSheetByName("Recipes");
  const itemsSheet = ss.getSheetByName("Items");
  if (!pricesSheet || !recipesSheet || !itemsSheet) {
    throw new Error('Missing sheet: need "Prices", "Recipes", and "Items".');
  }

  // Read only the used ranges (fast)
  const pricesData = readUsed_(pricesSheet, 3);  // A:C
  const recipesData = readUsed_(recipesSheet, 4); // A:D
  const itemsData = readUsed_(itemsSheet, 4);     // A:D

  const data = { pricesData, recipesData, itemsData };
  _engine = buildEngineFromData_(data);

  // Cache for ~5 minutes (tune if desired)
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
  // Expect headers row 1
  const priceMap = new Map();
  for (let i = 1; i < data.pricesData.length; i++) {
    const row = data.pricesData[i];
    const k = norm(row[0]);
    if (!k) continue;
    const buy = Number(row[1]);
    // Store 0 if missing/blank; caller decides what "hasTP" means
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
    // If multiple rows have inconsistent outQty, first one wins (fine for your “single recipe” rule)
    recipesByOutput.get(outKey).ingredients.push({ key: ingKey, qty: ingQty });
  }

  // ---- Memoized recursive evaluators ----
  const memoEff = new Map();
  const memoStrict = new Map();

  function effectiveCost_(key, stack) {
    if (!key) return "";
    if (memoEff.has(key)) return memoEff.get(key);

    stack = stack || [];
    if (stack.includes(key)) return memoEff.set(key, "").get(key); // loop guard

    const meta = metaMap.get(key) || null;
    const tpBuy = priceMap.has(key) ? priceMap.get(key) : 0;
    const hasTP = tpBuy > 0;
    const recipe = recipesByOutput.get(key) || null;

    // BLOCK always breaks
    if (meta && meta.mode === "BLOCK") {
      memoEff.set(key, "");
      return "";
    }

    // VENDOR requires manual cost
    if (meta && meta.mode === "VENDOR") {
      const v = meta.cost;
      memoEff.set(key, v == null ? "" : v);
      return memoEff.get(key);
    }

    // Manual override (any mode except BLOCK/VENDOR already handled)
    if (meta && meta.cost != null) {
      memoEff.set(key, meta.cost);
      return meta.cost;
    }

    // Compute strict craft cost if recipe exists (but ingredients use EFFECTIVE)
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
    if (meta && meta.mode === "ALT") {
      const v = craft != null ? craft : 0;
      memoEff.set(key, v);
      return v;
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
    if (stack.includes(key)) return memoStrict.set(key, "").get(key); // loop guard

    const meta = metaMap.get(key) || null;
    const recipe = recipesByOutput.get(key) || null;

    // BLOCK: break
    if (meta && meta.mode === "BLOCK") {
      memoStrict.set(key, "");
      return "";
    }

    // If output is VENDOR/manual-only and you still call strict craft: treat as that cost (useful for non-TP components)
    if (meta && meta.mode === "VENDOR") {
      const v = meta.cost;
      memoStrict.set(key, v == null ? "" : v);
      return memoStrict.get(key);
    }

    // If no recipe for the output: return N/A (prevents “profit” from spread)
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
  // Accept scalar or range
  if (items == null) return [[""]];
  if (!Array.isArray(items)) {
    const v = fn(items);
    return [[v]];
  }
  // 2D range: return 2D
  return items.map((row) => {
    const name = row && row.length ? row[0] : "";
    if (!name) return [""];
    return [fn(name)];
  });
}


