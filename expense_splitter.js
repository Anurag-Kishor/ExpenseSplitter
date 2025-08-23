/**
 * Calculates and processes expenses with subgroups in a Google Sheets spreadsheet.
 * 
 * This function performs the following steps:
 * 1. Identifies required column indexes based on headers.
 * 2. Builds subgroup clusters from the data.
 * 3. Assigns members to subgroups or self-groups.
 * 4. Processes expenses to calculate balances and per-item splits.
 * 5. Writes the results (per-item splits, balances, subgroup balances, and transactions) back to the sheet.
 * 
 * Prerequisites:
 * - The active sheet must have a specific structure with headers for members, subgroups, expenses, paid-by, amount, and split-between columns.
 * - Google Apps Script environment is required to execute this function.
 * 
 * @throws {Error} Logs an error and exits if any required columns are missing.
 */
function calculateExpensesWithSubgroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ===== Step 0: Read from separate sheets =====
  const membersSheet = ss.getSheetByName("Members");
  const subgroupsSheet = ss.getSheetByName("Subgroups");
  const expensesSheet = ss.getSheetByName("Expenses");

  if (!membersSheet || !subgroupsSheet || !expensesSheet) {
    throw new Error("Missing one of the required sheets: Members, Subgroups, or Expenses");
  }

  // Members list
  const headerRow = membersSheet.getRange(1, 1, 1, membersSheet.getLastColumn()).getValues()[0];
  // Find the column index for "Members" (case-insensitive)
  const membersColIndex = headerRow.findIndex(h => h.toString().trim().toLowerCase() === "members") + 1;
  if (membersColIndex === 0) {
    throw new Error('Could not find "Members" column in Members sheet');
  }
  const membersData = membersSheet.getRange(2, membersColIndex, membersSheet.getLastRow() - 1, 1).getValues();
  const allMembers = new Set(
    membersData.flat()
      .filter(Boolean)
      .map(name => String(name).trim().toLowerCase())
  );

  // Subgroups
  const subgroupsData = subgroupsSheet.getRange(2, 2, subgroupsSheet.getLastRow() - 1, 1).getValues();
  const subgroupMap = new Map();
  subgroupsData.flat().filter(Boolean).forEach(raw => {
    const group = raw.split(",").map(n => n.trim().toLowerCase()).filter(Boolean);
    if (group.length > 0) {
      subgroupMap.set(group.sort().join(","), new Set(group));
    }
  });

  // Expenses (header row included)
  const expensesData = expensesSheet.getDataRange().getValues();
  const headers = expensesData[0].map(h => h.toString().toLowerCase().trim());
  const {
    expensesStartCol,
    paidByColIndex,
    amountColIndex,
    splitBetweenColIndex
  } = getColumnIndexesExpenses(headers);

  // ===== Step 1: Assign members to subgroups =====
  const nameToGroupKey = new Map();
  const nameToGroup = new Map();
  assignMembersToSubgroups(allMembers, subgroupMap, nameToGroupKey, nameToGroup);

  // ===== Step 2: Process expenses =====
  const balances = {};
  const perItemSplits = {};
  allMembers.forEach(name => balances[name] = 0);

  processExpenses(expensesData, expensesStartCol, paidByColIndex, amountColIndex, splitBetweenColIndex, allMembers, balances, perItemSplits);

  // ===== Step 3: Create or clear output sheet =====
  const outputSheetName = "Expense Output";
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clear();
  }

  // ===== Step 4: Write output =====
  writePerItemSplitTable(outputSheet, allMembers, perItemSplits, balances);
  writeBalanceTable(outputSheet, balances);
  writeSubgroupBalanceTable(outputSheet, subgroupMap, nameToGroupKey, balances);
  writeTransactionTable(outputSheet, subgroupMap, nameToGroupKey, balances);
  writeMemberTransactionTable(outputSheet, balances);

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Expenses calculated successfully!", "Done", 5);
}

/**
 * Finds column indexes for the Expenses sheet
 */
function getColumnIndexesExpenses(headers) {
  return {
    expensesStartCol: headers.indexOf("item"),
    paidByColIndex: headers.indexOf("paid by"),
    amountColIndex: headers.indexOf("amount"),
    splitBetweenColIndex: headers.indexOf("split between")
  };
}

/**
 * Builds subgroup clusters from the data.
 * @param {Array} data - Spreadsheet data.
 * @param {number} membersColIndex - Index of the members column.
 * @param {number} subgroupsColIndex - Index of the subgroups column.
 * @param {Set} allMembers - Set to store all unique members.
 * @param {Map} subgroupMap - Map to store subgroup clusters.
 */
function buildSubgroupClusters(data, membersColIndex, subgroupsColIndex, allMembers, subgroupMap) {
  for (let i = 1; i < data.length; i++) {
    const memberRaw = data[i][membersColIndex];
    const subgroupRaw = data[i][subgroupsColIndex];

    if (memberRaw) {
      allMembers.add(memberRaw.trim().toLowerCase());
    }

    if (subgroupRaw) {
      const group = subgroupRaw
        .split(",")
        .map(name => name.trim().toLowerCase())
        .filter(Boolean);
      const groupKey = group.sort().join(",");
      if (!subgroupMap.has(groupKey)) {
        subgroupMap.set(groupKey, new Set(group));
      }
    }
  }
}

/**
 * Assigns members to subgroups or creates self-groups if no subgroup is found.
 * @param {Set} allMembers - Set of all unique members.
 * @param {Map} subgroupMap - Map of subgroup clusters.
 * @param {Map} nameToGroupKey - Map to store member-to-group key mapping.
 * @param {Map} nameToGroup - Map to store member-to-group mapping.
 */
function assignMembersToSubgroups(allMembers, subgroupMap, nameToGroupKey, nameToGroup) {
  for (const member of allMembers) {
    let found = false;
    for (const [key, group] of subgroupMap.entries()) {
      if (group.has(member)) {
        nameToGroupKey.set(member, key);
        nameToGroup.set(member, Array.from(group));
        found = true;
        break;
      }
    }
    if (!found) {
      const key = member;
      subgroupMap.set(key, new Set([member]));
      nameToGroupKey.set(member, key);
      nameToGroup.set(member, [member]);
    }
  }
}

/**
 * Processes expenses to calculate balances and per-item splits.
 * @param {Array} data - Spreadsheet data.
 * @param {number} expensesStartCol - Index of the item column.
 * @param {number} paidByColIndex - Index of the paid-by column.
 * @param {number} amountColIndex - Index of the amount column.
 * @param {number} splitBetweenColIndex - Index of the split-between column.
 * @param {Set} allMembers - Set of all unique members.
 * @param {Object} balances - Object to store balances for each member.
 * @param {Object} perItemSplits - Object to store per-item splits.
 */
function processExpenses(data, expensesStartCol, paidByColIndex, amountColIndex, splitBetweenColIndex, allMembers, balances, perItemSplits) {
  for (let i = 1; i < data.length; i++) {
    const item = data[i][expensesStartCol];
    const paidByRaw = data[i][paidByColIndex];
    const amount = parseFloat(data[i][amountColIndex]);
    const splitBetweenRaw = data[i][splitBetweenColIndex];

    if (!item || !paidByRaw || isNaN(amount)) continue;

    const paidBy = paidByRaw.trim().toLowerCase();
    let participants = getParticipants(splitBetweenRaw, allMembers);
    if (participants.length === 0) continue;

    // Weighted split
    const totalWeight = participants.reduce((sum, p) => sum + p.weight, 0);

    if (!perItemSplits[item]) perItemSplits[item] = {};

    // Subtract share for each participant
    participants.forEach(({ name, weight }) => {
      const share = amount * (weight / totalWeight);
      if (!balances.hasOwnProperty(name)) balances[name] = 0;
      balances[name] -= share;
      perItemSplits[item][name] = (perItemSplits[item][name] || 0) - share;
    });

    // Add full amount to payer
    if (!balances.hasOwnProperty(paidBy)) balances[paidBy] = 0;
    balances[paidBy] += amount;
    perItemSplits[item][paidBy] = (perItemSplits[item][paidBy] || 0) + amount;
  }
}


/**
 * Retrieves participants for an expense based on the split-between column.
 * @param {string} splitBetweenRaw - Raw value from the split-between column.
 * @param {Set} allMembers - Set of all unique members.
 * @returns {Array} - Array of participant names.
 */
function getParticipants(splitBetweenRaw, allMembers) {
  const participants = [];

  if (splitBetweenRaw && splitBetweenRaw.trim() === "*") {
    // All members with equal weight
    allMembers.forEach(m => participants.push({ name: m, weight: 1 }));
    return participants;
  }

  splitBetweenRaw
    .split(",")
    .map(n => n.trim().toLowerCase())
    .filter(Boolean)
    .forEach(entry => {
      let [name, weight] = entry.split(":").map(s => s.trim());
      weight = parseFloat(weight);
      if (isNaN(weight) || weight <= 0) weight = 1; // Default weight = 1
      participants.push({ name, weight });
    });

  // Remove duplicates by name (keep first weight found)
  const unique = [];
  const seen = new Set();
  for (const p of participants) {
    if (!seen.has(p.name)) {
      unique.push(p);
      seen.add(p.name);
    }
  }

  return unique;
}


function writePerItemSplitTable(sheet, allMembers, perItemSplits, balances) {
  const itemHeader = ["Item", ...Array.from(allMembers).map(name => capitalize(name))];
  const perItemRows = [];

  // Create rows for each item with the corresponding per member net share
  for (const item in perItemSplits) {
    const row = [item];
    for (const member of allMembers) {
      const memberShare = perItemSplits[item][member] || 0;
      row.push(parseFloat(memberShare.toFixed(2)));
    }
    perItemRows.push(row);
  }

  const perItemTotalRow = ["Person's Net Total"];
  for (const member of allMembers) {
    const totalShare = Object.values(perItemSplits).reduce((sum, itemShares) => {
      return sum + (parseFloat(itemShares[member] || 0));
    }, 0);
    perItemTotalRow.push(parseFloat(totalShare.toFixed(2)));
  }

  const fullPerItemTable = [itemHeader, ...perItemRows, perItemTotalRow];

  // Write per item split table
  const startRow = 1;
  const itemStartCol = 11;
  const perItemRange = sheet.getRange(startRow, itemStartCol, fullPerItemTable.length, fullPerItemTable[0].length);
  perItemRange.setValues(fullPerItemTable);
  styleTable(perItemRange);

  // Apply color formatting for positive/negative values (skip header col)
  const numRows = fullPerItemTable.length - 1; // excluding header
  const numCols = fullPerItemTable[0].length - 1; // excluding "Item" col
  const valueRange = sheet.getRange(startRow + 1, itemStartCol + 1, numRows, numCols);

  valueRange.setFontWeight("bold"); // make numbers bold
  const values = valueRange.getValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = valueRange.getCell(r + 1, c + 1);
      if (values[r][c] > 0) {
        cell.setFontColor("green");
      } else if (values[r][c] < 0) {
        cell.setFontColor("red");
      } else {
        cell.setFontColor("black");
      }
    }
  }
}


/**
 * Writes the balance table to the sheet.
 * @param {Object} sheet - Google Sheets object.
 * @param {Object} balances - Object containing balances for each member.
 */
function writeBalanceTable(sheet, balances) {
  const balanceStartRow = 1 + sheet.getLastRow() + 2;
  const balanceOutput = [["Name", "Balance (₹)"]];

  for (const [name, balance] of Object.entries(balances)) {
    balanceOutput.push([capitalize(name), balance.toFixed(2)]);
  }

  const balanceRange = sheet.getRange(balanceStartRow, 11, balanceOutput.length, 2);
  balanceRange.setValues(balanceOutput);
  styleTable(balanceRange);
}

/**
 * Writes the subgroup balance table to the sheet.
 * @param {Object} sheet - Google Sheets object.
 * @param {Map} subgroupMap - Map of subgroup clusters.
 * @param {Map} nameToGroupKey - Map of member-to-group key mapping.
 * @param {Object} balances - Object containing balances for each member.
 */
function writeSubgroupBalanceTable(sheet, subgroupMap, nameToGroupKey, balances) {
  const groupBalances = calculateGroupBalances(subgroupMap, nameToGroupKey, balances);

  const subgroupStartRow = sheet.getLastRow() + 2;
  const subgroupOutput = [["Subgroup", "Total Balance (₹)"]];

  for (const [groupKey, total] of groupBalances.entries()) {
    const members = Array.from(subgroupMap.get(groupKey)).map(capitalize).join(", ");
    subgroupOutput.push([members, total.toFixed(2)]);
  }

  const subgroupRange = sheet.getRange(subgroupStartRow, 11, subgroupOutput.length, 2);
  subgroupRange.setValues(subgroupOutput);
  styleTable(subgroupRange);
}

/**
 * Calculates balances for each subgroup.
 * @param {Map} subgroupMap - Map of subgroup clusters.
 * @param {Map} nameToGroupKey - Map of member-to-group key mapping.
 * @param {Object} balances - Object containing balances for each member.
 * @returns {Map} - Map of subgroup balances.
 */
function calculateGroupBalances(subgroupMap, nameToGroupKey, balances) {
  const groupBalances = new Map();
  for (const [name, balance] of Object.entries(balances)) {
    const groupKey = nameToGroupKey.get(name);
    if (!groupBalances.has(groupKey)) {
      groupBalances.set(groupKey, 0);
    }
    groupBalances.set(groupKey, groupBalances.get(groupKey) + balance);
  }
  return groupBalances;
}

/**
 * Writes the transaction table to the sheet.
 * @param {Object} sheet - Google Sheets object.
 * @param {Map} subgroupMap - Map of subgroup clusters.
 * @param {Map} nameToGroupKey - Map of member-to-group key mapping.
 * @param {Object} balances - Object containing balances for each member.
 */
function writeTransactionTable(sheet, subgroupMap, nameToGroupKey, balances) {
  const groupBalances = calculateGroupBalances(subgroupMap, nameToGroupKey, balances);
  const transactions = generateMinimalSubgroupTransactions(groupBalances);

  const transactionStartRow = sheet.getLastRow() + 2;
  const transactionStartCol = 11;
  const transactionOutput = [["From Subgroup", "To Subgroup", "Amount (₹)"], ...transactions];

  const transactionRange = sheet.getRange(transactionStartRow, transactionStartCol, transactionOutput.length, 3);
  transactionRange.setValues(transactionOutput);
  styleTable(transactionRange);
}
/**
 * Generate minimal member-level transactions to settle balances.
 * @param {Object} balances - Map of member -> balance (positive = gets money, negative = owes money).
 * @returns {Array} - Array of transactions [from, to, amount].
 */
function generateMinimalMemberTransactions(balances) {
  const creditors = [];
  const debtors = [];

  // Separate creditors and debtors
  for (const [name, balance] of Object.entries(balances)) {
    const val = parseFloat(balance.toFixed(2));
    if (val > 0.01) creditors.push({ name, amount: val });
    else if (val < -0.01) debtors.push({ name, amount: -val });
  }

  const transactions = [];

  // Minimize transactions
  while (debtors.length && creditors.length) {
    // Find max creditor
    creditors.sort((a, b) => b.amount - a.amount);
    debtors.sort((a, b) => b.amount - a.amount);

    const debtor = debtors[0];
    const creditor = creditors[0];

    const amount = Math.min(debtor.amount, creditor.amount);
    transactions.push([capitalize(debtor.name), capitalize(creditor.name), amount.toFixed(2)]);

    debtor.amount -= amount;
    creditor.amount -= amount;

    if (debtor.amount < 0.01) debtors.shift();
    if (creditor.amount < 0.01) creditors.shift();
  }

  return transactions;
}

function writeMemberTransactionTable(sheet, balances) {
  const transactions = generateMinimalMemberTransactions(balances);
  const startRow = sheet.getLastRow() + 2;
  const output = [["From", "To", "Amount (₹)"], ...transactions];

  const range = sheet.getRange(startRow, 11, output.length, 3);
  range.setValues(output);
  styleTable(range);
}

function generateMinimalSubgroupTransactions(groupBalances) {
  const creditors = [];
  const debtors = [];

  for (const [groupKey, balance] of groupBalances.entries()) {
    const val = parseFloat(balance.toFixed(2));
    if (val > 0.01) creditors.push({ groupKey, amount: val });
    else if (val < -0.01) debtors.push({ groupKey, amount: -val });
  }

  const transactions = [];

  while (creditors.length && debtors.length) {
    creditors.sort((a, b) => b.amount - a.amount);
    debtors.sort((a, b) => b.amount - a.amount);

    const creditor = creditors[0];
    const debtor = debtors[0];
    const amount = Math.min(creditor.amount, debtor.amount);

    transactions.push([
      formatGroup(debtor.groupKey),
      formatGroup(creditor.groupKey),
      amount.toFixed(2)
    ]);

    creditor.amount -= amount;
    debtor.amount -= amount;

    if (creditor.amount < 0.01) creditors.shift();
    if (debtor.amount < 0.01) debtors.shift();
  }

  return transactions;
}

/**
 * Formats a group key into a readable string of member names.
 * @param {string} groupKey - Group key representing subgroup members.
 * @returns {string} - Formatted string of member names.
 */
function formatGroup(groupKey) {
  const members = groupKey.split(",").map(name => capitalize(name)).join(", ");
  return members;
}

/**
 * Capitalizes the first letter of a string.
 * @param {string} str - Input string.
 * @returns {string} - Capitalized string.
 */
function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

/**
 * Constants for styling
 */
const HEADER_BACKGROUND_COLOR = "#4CAF50";
const HEADER_FONT_COLOR = "white";
const BODY_BACKGROUND_COLOR = "#ffffff";
const ALTERNATE_ROW_COLOR = "#f9f9f9";

/**
 * Styles a table range in the sheet.
 * @param {Object} range - Google Sheets range object.
 */
function styleTable(range) {
  const sheet = range.getSheet();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  const header = range.offset(0, 0, 1, numCols);
  const body = range.offset(1, 0, numRows - 1, numCols);

  // Set borders for the whole table
  range.setBorder(true, true, true, true, true, true);

  // Header styling: Bold, background color, centered text
  header.setFontWeight("bold")
    .setBackground(HEADER_BACKGROUND_COLOR)
    .setFontSize(12)
    .setHorizontalAlignment("center")
    .setFontColor(HEADER_FONT_COLOR);

  // Body styling: alternating row colors, smaller font size, centered text
  body.setFontSize(10);
  body.setBackground(BODY_BACKGROUND_COLOR);
  body.setHorizontalAlignment("center");

  for (let i = 1; i < numRows; i++) {
    if (i % 2 === 1) {
      body.getCell(i, 1).setBackground(ALTERNATE_ROW_COLOR);  // Alternate row colors
    }
  }

  // Row height adjustment
  sheet.setRowHeights(range.getRow(), numRows, 30);
  

  // Resize all columns in the range at once
  sheet.autoResizeColumns(range.getColumn(), numCols);
}