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
  SpreadsheetApp.getActiveSpreadsheet().toast("Refreshing...");
  const sheet = SpreadsheetApp.getActive().getSheetByName("Profit");
  const range = sheet.getRange("A3:A");
  const values = range.getValues();
  range.setValues(values);  // re-writes same values -> triggers recalculation
}

function CRAFTCOST(items) {
  const ss = SpreadsheetApp.getActive();
  const pricesSheet = ss.getSheetByName('Prices');
  const recipesSheet = ss.getSheetByName('Recipes');
  const itemsSheet = ss.getSheetByName('Items');
  if (!pricesSheet || !recipesSheet || !itemsSheet) throw new Error('Missing necessary sheet: Prices, Recipes, or Items');

  const norm = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim().toLowerCase();

  // Prices: A name, B buy (Buy Order / Sell Insta), C sell (Sell Order / Buy Insta)
  const prices = pricesSheet.getRange(2, 1, Math.max(pricesSheet.getLastRow() - 1, 0), 3).getValues();
  const buyMap = {}, sellMap = {};
  prices.forEach(r => {
    const k = norm(r[0]);
    if (!k) return;
    buyMap[k] = Number(r[1]) || 0;
    sellMap[k] = Number(r[2]) || 0;
  });

  // Recipes: A output, B outQty, C ing, D ingQty
  const recipesData = recipesSheet.getRange(2, 1, Math.max(recipesSheet.getLastRow() - 1, 0), 4).getValues();
  const recipesByOutput = {};
  recipesData.forEach(r => {
    const outKey = norm(r[0]);
    const outQty = Number(r[1]) || 1;
    const ingKey = norm(r[2]);
    const qty = Number(r[3]) || 0;
    if (!outKey || !ingKey) return;
    if (!recipesByOutput[outKey]) recipesByOutput[outKey] = { outQty, ingredients: [] };
    recipesByOutput[outKey].ingredients.push({ key: ingKey, qty });
  });

  // Items meta: A name, C mode, D manual cost
  const itemsData = itemsSheet.getRange(2, 1, Math.max(itemsSheet.getLastRow() - 1, 0), 4).getValues();
  const itemMeta = {};
  itemsData.forEach(r => {
    const k = norm(r[0]);
    if (!k) return;
    itemMeta[k] = {
      mode: String(r[2] || "").toUpperCase().trim(), // ALT / VENDOR / BLOCK / ""
      cost: (r[3] !== "" && r[3] != null) ? Number(r[3]) : null
    };
  });

  // cache needs to respect topLevel vs nested behavior, so key it with a suffix
  const cache = {};

  function getTpFallback(key) {
    const buy = buyMap[key] || 0;
    const sell = sellMap[key] || 0;
    if (buy > 0) return buy;     // prefer buy order
    if (sell > 0) return sell;   // else instant-buy cost
    return null;
  }

  function calcCost(key, stack, topLevel) {
    if (!key) return "";
    const cacheKey = key + (topLevel ? "|T" : "|N");
    if (cache.hasOwnProperty(cacheKey)) return cache[cacheKey];

    stack = stack || [];
    if (stack.includes(key)) return cache[cacheKey] = ""; // cycle protection

    const meta = itemMeta[key] || null;
    const recipe = recipesByOutput[key] || null;

    // BLOCK: always break
    if (meta && meta.mode === "BLOCK") return cache[cacheKey] = "";

    // VENDOR: must have ManualCost
    if (meta && meta.mode === "VENDOR") {
      if (meta.cost === null) return cache[cacheKey] = "";
      return cache[cacheKey] = meta.cost;
    }

    // Manual override (any mode except BLOCK/VENDOR which already returned)
    if (meta && meta.cost !== null) return cache[cacheKey] = meta.cost;

    // Compute craft cost if recipe exists
    let craftCost = null;
    if (recipe) {
      let total = 0;
      for (const ing of recipe.ingredients) {
        const c = calcCost(ing.key, stack.concat(key), false);
        if (c === "" || c == null || c === "N/A") { total = null; break; }
        total += c * ing.qty;
      }
      if (total != null) craftCost = total / (recipe.outQty || 1);
    }

    // ALT: recipe cost if possible, else 0 (keeps chain alive)
    if (meta && meta.mode === "ALT") {
      const v = (craftCost != null) ? craftCost : 0;
      return cache[cacheKey] = v;
    }

    // If recipe exists: profit-mode => use craft cost (or blank if missing ingredient cost)
    if (recipe) {
      return cache[cacheKey] = (craftCost != null ? craftCost : "");
    }

    // No recipe:
    // - Top-level items you’re evaluating: show "N/A" so it doesn't look like craft profit
    // - Nested ingredients: use TP fallback so chains can still compute
    if (topLevel) return cache[cacheKey] = "N/A";

    const tp = getTpFallback(key);
    return cache[cacheKey] = (tp != null ? tp : "");
  }

  return items.map(row => {
    const k = norm(row[0]);
    if (!k) return [""];
    return [calcCost(k, [], true)];
  });
}

function CRAFTCOST_EFF(items) {
  const ss = SpreadsheetApp.getActive();
  const pricesSheet = ss.getSheetByName('Prices');
  const recipesSheet = ss.getSheetByName('Recipes');
  const itemsSheet = ss.getSheetByName('Items');
  if (!pricesSheet || !recipesSheet || !itemsSheet) throw new Error('Missing necessary sheet: Prices, Recipes, or Items');

  const norm = (v) => String(v ?? "").replace(/\u00A0/g, " ").trim().toLowerCase();

  // Prices: A name, B buy (Buy Order / Sell Insta), C sell (Sell Order / Buy Insta)
  const prices = pricesSheet.getRange(2, 1, Math.max(pricesSheet.getLastRow() - 1, 0), 3).getValues();
  const buyMap = {}, sellMap = {};
  prices.forEach(r => {
    const k = norm(r[0]);
    if (!k) return;
    buyMap[k] = Number(r[1]) || 0;
    sellMap[k] = Number(r[2]) || 0;
  });

  // Recipes: A output, B outQty, C ing, D ingQty
  const recipesData = recipesSheet.getRange(2, 1, Math.max(recipesSheet.getLastRow() - 1, 0), 4).getValues();
  const recipesByOutput = {};
  recipesData.forEach(r => {
    const outKey = norm(r[0]);
    const outQty = Number(r[1]) || 1;
    const ingKey = norm(r[2]);
    const qty = Number(r[3]) || 0;
    if (!outKey || !ingKey) return;
    if (!recipesByOutput[outKey]) recipesByOutput[outKey] = { outQty, ingredients: [] };
    recipesByOutput[outKey].ingredients.push({ key: ingKey, qty });
  });

  // Items meta: A name, C mode, D manual cost
  const itemsData = itemsSheet.getRange(2, 1, Math.max(itemsSheet.getLastRow() - 1, 0), 4).getValues();
  const itemMeta = {};
  itemsData.forEach(r => {
    const k = norm(r[0]);
    if (!k) return;
    itemMeta[k] = {
      mode: String(r[2] || "").toUpperCase().trim(), // ALT / VENDOR / BLOCK / ""
      cost: (r[3] !== "" && r[3] != null) ? Number(r[3]) : null
    };
  });

  const cache = {};

  function getTpAcquire(key) {
    const buy = buyMap[key] || 0;
    const sell = sellMap[key] || 0;
    if (buy > 0) return buy;
    if (sell > 0) return sell;
    return null;
  }

  // Returns "effective cost" for ingredients: min(TP acquire, craft) when craftable
  function effCost(key, stack) {
    if (!key) return "";
    if (cache.hasOwnProperty(key)) return cache[key];

    stack = stack || [];
    if (stack.includes(key)) return cache[key] = "";

    const meta = itemMeta[key] || null;
    const recipe = recipesByOutput[key] || null;

    if (meta && meta.mode === "BLOCK") return cache[key] = "";

    if (meta && meta.mode === "VENDOR") {
      if (meta.cost === null) return cache[key] = "";
      return cache[key] = meta.cost;
    }

    if (meta && meta.cost !== null) return cache[key] = meta.cost;

    let craftCost = null;
    if (recipe) {
      let total = 0;
      for (const ing of recipe.ingredients) {
        const c = effCost(ing.key, stack.concat(key));
        if (c === "" || c == null || c === "N/A") { total = null; break; }
        total += c * ing.qty;
      }
      if (total != null) craftCost = total / (recipe.outQty || 1);
    }

    if (meta && meta.mode === "ALT") {
      const v = (craftCost != null) ? craftCost : 0;
      return cache[key] = v;
    }

    const tp = getTpAcquire(key);
    let final = "";

    if (craftCost != null && tp != null) final = Math.min(craftCost, tp);
    else if (craftCost != null) final = craftCost;
    else if (tp != null) final = tp;
    else final = "";

    return cache[key] = final;
  }

  // Top-level: must have a recipe; otherwise N/A
  function topCraftCost(key) {
    const recipe = recipesByOutput[key] || null;
    if (!recipe) return "N/A";

    let total = 0;
    for (const ing of recipe.ingredients) {
      const c = effCost(ing.key, [key]);
      if (c === "" || c == null || c === "N/A") return "";
      total += c * ing.qty;
    }
    return total / (recipe.outQty || 1);
  }

  return items.map(row => {
    const k = norm(row[0]);
    if (!k) return [""];
    return [topCraftCost(k)];
  });
}

