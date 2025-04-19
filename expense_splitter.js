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
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
  
    // Step 1: Find column indexes based on headers
    const headers = data[0];
    const normalizedHeaders = headers.map(header => header.toLowerCase());
    const {
      membersColIndex,
      subgroupsColIndex,
      expensesStartCol,
      paidByColIndex,
      amountColIndex,
      splitBetweenColIndex
    } = getColumnIndexes(normalizedHeaders);
  
    // If any required columns are missing, log an error and exit
    if (membersColIndex === -1 || subgroupsColIndex === -1 || expensesStartCol === -1 || paidByColIndex === -1 || amountColIndex === -1 || splitBetweenColIndex === -1) {
      Logger.log("One or more required columns are missing.");
      return;
    }
  
    const allMembers = new Set();
    const subgroupMap = new Map();
    const nameToGroupKey = new Map();
    const nameToGroup = new Map();
  
    // Step 2: Build subgroup clusters
    buildSubgroupClusters(data, membersColIndex, subgroupsColIndex, allMembers, subgroupMap);
  
    // Step 3: Assign members to subgroups (or self)
    assignMembersToSubgroups(allMembers, subgroupMap, nameToGroupKey, nameToGroup);
  
    // Step 4: Process Expenses
    const balances = {};
    const perItemSplits = {};
    allMembers.forEach(name => balances[name] = 0);
  
    processExpenses(data, expensesStartCol, paidByColIndex, amountColIndex, splitBetweenColIndex, allMembers, balances, perItemSplits);
  
    // Step 5: Write tables to sheet
    writePerItemSplitTable(sheet, allMembers, perItemSplits, balances);
    writeBalanceTable(sheet, balances);
    writeSubgroupBalanceTable(sheet, subgroupMap, nameToGroupKey, balances);
    writeTransactionTable(sheet, subgroupMap, nameToGroupKey, balances);
  }
  
  /**
   * Finds the column indexes for required headers in the spreadsheet.
   * @param {Array} headers - Array of normalized header names.
   * @returns {Object} - Object containing column indexes for required headers.
   */
  function getColumnIndexes(headers) {
    return {
      membersColIndex: headers.indexOf('members'),
      subgroupsColIndex: headers.indexOf('subgroups'),
      expensesStartCol: headers.indexOf('item'),
      paidByColIndex: headers.indexOf('paid by'),
      amountColIndex: headers.indexOf('amount'),
      splitBetweenColIndex: headers.indexOf('split between')
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
  
      const share = amount / participants.length;
  
      participants.forEach(name => {
        if (!balances.hasOwnProperty(name)) balances[name] = 0;
        balances[name] -= share;
        if (!perItemSplits[item]) perItemSplits[item] = {};
        perItemSplits[item][name] = share.toFixed(2);
      });
  
      if (!balances.hasOwnProperty(paidBy)) balances[paidBy] = 0;
      balances[paidBy] += amount;
    }
  }
  
  /**
   * Retrieves participants for an expense based on the split-between column.
   * @param {string} splitBetweenRaw - Raw value from the split-between column.
   * @param {Set} allMembers - Set of all unique members.
   * @returns {Array} - Array of participant names.
   */
  function getParticipants(splitBetweenRaw, allMembers) {
    let participants = [];
    if (splitBetweenRaw && splitBetweenRaw.trim() === "*") {
      participants = Array.from(allMembers);
    } else {
      const names = splitBetweenRaw
        .split(",")
        .map(n => n.trim().toLowerCase())
        .filter(Boolean);
      participants = names;
    }
    return Array.from(new Set(participants));
  }
  
  /**
   * Writes the per-item split table to the sheet.
   * @param {Object} sheet - Google Sheets object.
   * @param {Set} allMembers - Set of all unique members.
   * @param {Object} perItemSplits - Object containing per-item splits.
   * @param {Object} balances - Object containing balances for each member.
   */
  function writePerItemSplitTable(sheet, allMembers, perItemSplits, balances) {
    const itemHeader = ["Item", ...Array.from(allMembers).map(name => capitalize(name))];
    const perItemRows = [];
  
    // Create rows for each item with the corresponding per member share
    for (const item in perItemSplits) {
      const row = [item];
      for (const member of allMembers) {
        const memberShare = perItemSplits[item][member] || 0;
        row.push(memberShare);
      }
      perItemRows.push(row);
    }
  
    const perItemTotalRow = ["Person's Share"];
    for (const member of allMembers) {
      const totalShare = Object.values(perItemSplits).reduce((sum, itemShares) => {
        return sum + (parseFloat(itemShares[member] || 0));
      }, 0);
      perItemTotalRow.push(totalShare.toFixed(2));
    }
  
    const fullPerItemTable = [itemHeader, ...perItemRows, perItemTotalRow];
  
    // Write per item split table
    const startRow = 1;
    const itemStartCol = 11;
    const perItemRange = sheet.getRange(startRow, itemStartCol, fullPerItemTable.length, fullPerItemTable[0].length);
    perItemRange.setValues(fullPerItemTable);
    styleTable(perItemRange);
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
    const transactions = generateTransactions(groupBalances);
  
    const transactionStartRow = sheet.getLastRow() + 2;
    const transactionStartCol = 11;
    const transactionOutput = [["From Subgroup", "To Subgroup", "Amount (₹)"], ...transactions];
  
    const transactionRange = sheet.getRange(transactionStartRow, transactionStartCol, transactionOutput.length, 3);
    transactionRange.setValues(transactionOutput);
    styleTable(transactionRange);
  }
  
  /**
   * Generates transactions between subgroups to settle balances.
   * @param {Map} groupBalances - Map of subgroup balances.
   * @returns {Array} - Array of transactions.
   */
  function generateTransactions(groupBalances) {
    const creditors = [];
    const debtors = [];
    const transactions = [];
  
    for (const [groupKey, balance] of groupBalances.entries()) {
      const val = parseFloat(balance.toFixed(2));
      if (val > 0.01) creditors.push({ groupKey, amount: val });
      else if (val < -0.01) debtors.push({ groupKey, amount: -val });
    }
  
    let i = 0, j = 0;
    while (i < debtors.length && j < creditors.length) {
      const debtor = debtors[i];
      const creditor = creditors[j];
  
      const amount = Math.min(debtor.amount, creditor.amount);
  
      transactions.push([
        formatGroup(debtor.groupKey),
        formatGroup(creditor.groupKey),
        amount.toFixed(2)
      ]);
  
      debtor.amount -= amount;
      creditor.amount -= amount;
  
      if (debtor.amount < 0.01) i++;
      if (creditor.amount < 0.01) j++;
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
      .setBackground("#4CAF50")
      .setFontSize(12)
      .setHorizontalAlignment("center")
      .setFontColor("white");
  
    // Body styling: alternating row colors, smaller font size, centered text
    body.setFontWeight("normal").setFontSize(10);
    body.setBackground("#ffffff");
    body.setHorizontalAlignment("center");
  
    for (let i = 1; i < numRows; i++) {
      if (i % 2 === 1) {
        body.getCell(i, 1).setBackground("#f9f9f9");  // Alternate row colors
      }
    }
  
    // Row height adjustment
    for (let i = 0; i < numRows; i++) {
      sheet.setRowHeight(range.getRow() + i, 30);
    }
  
    // Auto-resize columns to fit content
    for (let i = 0; i < numCols; i++) {
      sheet.autoResizeColumn(range.getColumn() + i);
    }
  }