/*******************************************************
 * PERSONAL FINANCE TRACKER
 * One-time setup script  —  v2 (fixed)
 *
 * Main function:
 *     setupFinanceTracker()
 *
 * After setup:
 *     Google Form → Form Responses → Monthly Sheet
 *
 * Monthly sheet:
 *     Sep-2026
 *     Oct-2026
 *     etc.
 *
 * Invoice detail:
 *     Sep-2026_Items
 *     Oct-2026_Items
 *
 * FIXES IN THIS VERSION:
 *   1. Transaction Date is now OPTIONAL. Leave it blank on
 *      the form to use today's date automatically.
 *   2. Date parser now correctly reads the M/D/YYYY format
 *      that this form actually sends (e.g. "8/31/2026" =
 *      31 August 2026), instead of misreading it as
 *      day-first and instead of drifting a day due to UTC
 *      conversion. Whatever month the entered date falls
 *      in, the transaction is written into THAT month's
 *      sheet (creating it if needed) — not always the
 *      current month.
 *   3. getFormValue_() is now whitespace-tolerant when
 *      matching question titles, so a stray space in a
 *      question title can no longer cause a silent empty
 *      read (which was crashing the whole trigger with
 *      "Invalid amount").
 *   4. Amount parsing now trims/validates more defensively
 *      and reports a clearer error including the raw
 *      value received, if it's ever still invalid.
 *   5. All dropdown (List) questions and the initial
 *      Multiple Choice question now have helpful
 *      descriptions (.setHelpText()).
 *   6. Added updateExistingFormDescriptions() — a one-off
 *      utility to patch descriptions and the "optional"
 *      setting onto a form that was already created by an
 *      earlier run, without recreating the form.
 *   7. Diagnostic Logger.log lines added in
 *      processFinanceTransaction so Executions → Cloud
 *      logs show exactly what was received from the form
 *      on every submission, making future issues easy to
 *      diagnose.
 *******************************************************/


/***********************
 * GLOBAL CONFIGURATION
 ***********************/

const CONFIG = {
  APP_NAME: "My Finance Tracker",

  FORM_TITLE: "💰 My Finance Tracker",

  FORM_DESCRIPTION:
    "Record your income, expenses, transfers, investments, lending and borrowing in a few seconds.",

  DASHBOARD: "Dashboard",
  SETTINGS: "Settings",
  CATEGORIES: "Categories",
  ACCOUNTS: "Accounts",

  TRANSACTION_SHEET_SUFFIX: "",
  ITEM_SHEET_SUFFIX: "_Items",

  CURRENCY: "₹",

  DATE_FORMAT: "dd-MMM-yyyy",
  DATETIME_FORMAT: "dd-MMM-yyyy hh:mm AM/PM",

  MONTH_FORMAT: "MMM-yyyy"
};


/***********************
 * MAIN SETUP FUNCTION
 ***********************/

function setupFinanceTracker() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("Please create or open a Google Sheet first.");
  }

  ss.setName(CONFIG.APP_NAME);

  // Create base sheets
  createSettingsSheet_(ss);
  createCategoriesSheet_(ss);
  createAccountsSheet_(ss);
  createDashboardSheet_(ss);

  // Create Google Form
  const form = createFinanceForm_(ss);

  // Install trigger
  installFormSubmitTrigger_(ss);

  // Create current month's sheet
  const today = new Date();
  const monthName = getMonthSheetName_(today);

  createMonthlyTransactionSheet_(ss, monthName);
  createMonthlyItemSheet_(ss, monthName);

  // Formatting
  formatBaseSheets_(ss);

  // Dashboard update
  updateDashboard_(ss);

  // Dashboard filter dropdowns + filtered views (Year / Category)
  setupDashboardFilters_(ss);
  refreshFilteredDashboard_(ss);

  // Save important IDs
  saveConfiguration_(ss, form);

  SpreadsheetApp.flush();

  Logger.log("SETUP COMPLETED");
  Logger.log("Google Form URL: " + form.getPublishedUrl());
  Logger.log("Google Form Edit URL: " + form.getEditUrl());

  SpreadsheetApp.getUi().alert(
    "Finance Tracker Setup Completed!\n\n" +
    "Your Google Form has been created.\n\n" +
    "Open Apps Script → Executions/Logs to see the Form URLs."
  );
}


/*******************************************************
 * GOOGLE FORM CREATION
 *******************************************************/

function createFinanceForm_(ss) {

  // Check whether form already exists
  const properties = PropertiesService.getScriptProperties();

  const existingFormId =
    properties.getProperty("FINANCE_FORM_ID");

  if (existingFormId) {

    try {
      const existingForm =
        FormApp.openById(existingFormId);

      return existingForm;

    } catch (e) {
      // Continue and create a new form
    }
  }

  const form =
    FormApp.create(CONFIG.FORM_TITLE);

  form.setDescription(CONFIG.FORM_DESCRIPTION);

  form.setConfirmationMessage(
    "✅ Transaction saved successfully!\n\n" +
    "You can close this form."
  );

  /***********************
   * 1. TRANSACTION TYPE
   ***********************/

  form.addMultipleChoiceItem()
    .setTitle("What do you want to record?")
    .setHelpText(
      "Choose the type of transaction you are logging today."
    )
    .setChoiceValues([
      "Expense",
      "Money Received",
      "Transfer",
      "Money Lent",
      "Money Borrowed",
      "Investment"
    ])
    .setRequired(true);


  /***********************
   * 2. TRANSACTION DATE
   ***********************/

  form.addDateItem()
    .setTitle("Transaction Date")
    .setHelpText(
      "Leave blank to use today's date. Enter a date only if this transaction happened on a different day (e.g. last month) — it will be recorded in that month's sheet."
    )
    .setRequired(false);


  /***********************
   * 3. AMOUNT
   ***********************/

  form.addTextItem()
    .setTitle("Amount (₹)")
    .setHelpText(
      "Enter the transaction amount as a number, e.g. 250 or 1500.50"
    )
    .setRequired(true);


  /***********************
   * 4. CATEGORY
   ***********************/

  form.addListItem()
    .setTitle("Category")
    .setHelpText(
      "Select the category this transaction belongs to (e.g. Food, Bills, Salary)."
    )
    .setChoiceValues(getDefaultCategories_())
    .setRequired(true);


  /***********************
   * 5. DESCRIPTION
   ***********************/

  form.addTextItem()
    .setTitle("Description")
    .setHelpText(
      "Example: Monthly grocery, petrol, electricity bill."
    )
    .setRequired(false);


  /***********************
   * 6. FROM / TO
   ***********************/

  form.addTextItem()
    .setTitle("From / To")
    .setHelpText(
      "Example: ABC Store, Rahul, Company."
    )
    .setRequired(false);


  /***********************
   * 7. ACCOUNT / WALLET
   ***********************/

  form.addListItem()
    .setTitle("Account / Wallet")
    .setHelpText(
      "Select the account or wallet used for this transaction."
    )
    .setChoiceValues(getDefaultAccounts_())
    .setRequired(true);


  /***********************
   * 8. PAYMENT METHOD
   ***********************/

  form.addListItem()
    .setTitle("Payment Method")
    .setHelpText(
      "Select how this transaction was paid or received."
    )
    .setChoiceValues([
      "UPI",
      "Cash",
      "Debit Card",
      "Credit Card",
      "Bank Transfer",
      "Auto Debit",
      "Other"
    ])
    .setRequired(true);


  /***********************
   * 9. EXPENSE NATURE
   ***********************/

  form.addListItem()
    .setTitle("Expense Nature")
    .setHelpText(
      "Classify this expense as a Need, a Want, or an Investment. Choose 'Not Applicable' if this isn't an expense."
    )
    .setChoiceValues([
      "Need",
      "Want",
      "Investment",
      "Not Applicable"
    ])
    .setRequired(false);


  /***********************
   * 10. FIXED / VARIABLE
   ***********************/

  form.addListItem()
    .setTitle("Fixed or Variable")
    .setHelpText(
      "Fixed = same amount every time (e.g. rent, EMI). Variable = amount changes each time (e.g. groceries)."
    )
    .setChoiceValues([
      "Fixed",
      "Variable",
      "Not Applicable"
    ])
    .setRequired(false);


  /***********************
   * 11. RECURRING
   ***********************/

  form.addListItem()
    .setTitle("Recurring?")
    .setHelpText(
      "Does this transaction repeat on a schedule, or is it a one-time entry?"
    )
    .setChoiceValues([
      "No",
      "Yes - Monthly",
      "Yes - Yearly",
      "Other"
    ])
    .setRequired(false);


 /***********************
 * 12. INVOICE PHOTO
 ***********************/

// File upload question must be added manually in Google Forms.
// Apps Script FormApp does not support creating File Upload items.
  /***********************
   * 13. NOTES
   ***********************/

  form.addParagraphTextItem()
    .setTitle("Notes")
    .setHelpText(
      "Optional: add any extra detail about this transaction."
    )
    .setRequired(false);


  /***********************
   * LINK FORM TO SHEET
   ***********************/

  form.setDestination(
    FormApp.DestinationType.SPREADSHEET,
    ss.getId()
  );

  return form;
}


/*******************************************************
 * ONE-OFF UTILITY:
 * PATCH DESCRIPTIONS + "OPTIONAL DATE" ONTO AN
 * ALREADY-CREATED FORM
 *
 * Run this function once (select it in the Apps Script
 * function dropdown → Run) if your form already exists
 * and you just want to apply the fixes WITHOUT recreating
 * the form or losing the existing shareable link /
 * responses.
 *******************************************************/

function updateExistingFormDescriptions() {

  const properties =
    PropertiesService.getScriptProperties();

  const formId =
    properties.getProperty("FINANCE_FORM_ID");

  if (!formId) {
    SpreadsheetApp.getUi().alert(
      "No existing form found (FINANCE_FORM_ID not set).\n\n" +
      "Run setupFinanceTracker() instead to create a new form."
    );
    return;
  }

  const form = FormApp.openById(formId);

  // Map of exact question Title -> help text to apply
  const helpTextByTitle = {
    "What do you want to record?":
      "Choose the type of transaction you are logging today.",

    "Transaction Date":
      "Leave blank to use today's date. Enter a date only if this transaction happened on a different day (e.g. last month) — it will be recorded in that month's sheet.",

    "Amount (₹)":
      "Enter the transaction amount as a number, e.g. 250 or 1500.50",

    "Category":
      "Select the category this transaction belongs to (e.g. Food, Bills, Salary).",

    "Account / Wallet":
      "Select the account or wallet used for this transaction.",

    "Payment Method":
      "Select how this transaction was paid or received.",

    "Expense Nature":
      "Classify this expense as a Need, a Want, or an Investment. Choose 'Not Applicable' if this isn't an expense.",

    "Fixed or Variable":
      "Fixed = same amount every time (e.g. rent, EMI). Variable = amount changes each time (e.g. groceries).",

    "Recurring?":
      "Does this transaction repeat on a schedule, or is it a one-time entry?",

    "Notes":
      "Optional: add any extra detail about this transaction."
  };

  const items = form.getItems();
  let updatedCount = 0;

  items.forEach(function(item) {

    const title = item.getTitle();

    if (helpTextByTitle.hasOwnProperty(title)) {

      const type = item.getType();

      if (type === FormApp.ItemType.LIST) {
        item.asListItem().setHelpText(helpTextByTitle[title]);
        updatedCount++;
      } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        item.asMultipleChoiceItem().setHelpText(helpTextByTitle[title]);
        updatedCount++;
      } else if (type === FormApp.ItemType.PARAGRAPH_TEXT) {
        item.asParagraphTextItem().setHelpText(helpTextByTitle[title]);
        updatedCount++;
      } else if (type === FormApp.ItemType.TEXT) {
        item.asTextItem().setHelpText(helpTextByTitle[title]);
        updatedCount++;
      } else if (type === FormApp.ItemType.DATE) {
        item.asDateItem()
          .setHelpText(helpTextByTitle[title])
          .setRequired(false); // make Transaction Date optional
        updatedCount++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    "Updated " + updatedCount + " question(s).\n\n" +
    "Transaction Date is now optional. Open the form to confirm."
  );

  Logger.log("Updated " + updatedCount + " question(s).");
}


/*******************************************************
 * FORM SUBMIT TRIGGER
 *******************************************************/

function installFormSubmitTrigger_(ss) {

  const triggers =
    ScriptApp.getProjectTriggers();

  // Remove duplicate finance triggers
  triggers.forEach(function(trigger) {

    if (
      trigger.getHandlerFunction() ===
      "processFinanceTransaction"
    ) {

      ScriptApp.deleteTrigger(trigger);
    }

  });


  ScriptApp.newTrigger(
    "processFinanceTransaction"
  )
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}


/*******************************************************
 * FORM SUBMISSION PROCESSING
 *******************************************************/

function processFinanceTransaction(e) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!e || !e.namedValues) {
    return;
  }

  const data = e.namedValues;

  // Diagnostic: shows exactly what Forms sent, visible in
  // Apps Script -> Executions -> Cloud logs for this run.
  Logger.log("ALL NAMED VALUES: " + JSON.stringify(data));


  /***********************
   * READ FORM VALUES
   ***********************/

  const transactionType =
    getFormValue_(data, "What do you want to record?");

  const transactionDateRaw =
    getFormValue_(data, "Transaction Date");

  const amountRaw =
    getFormValue_(data, "Amount (₹)");

  const category =
    getFormValue_(data, "Category");

  const description =
    getFormValue_(data, "Description");

  const fromTo =
    getFormValue_(data, "From / To");

  const account =
    getFormValue_(data, "Account / Wallet");

  const paymentMethod =
    getFormValue_(data, "Payment Method");

  const expenseNature =
    getFormValue_(data, "Expense Nature");

  const fixedVariable =
    getFormValue_(data, "Fixed or Variable");

  const recurring =
    getFormValue_(data, "Recurring?");

  const invoice =
    getFormValue_(data, "Bill / Invoice Photo");

  const notes =
    getFormValue_(data, "Notes");

  Logger.log("RAW DATE VALUE: [" + transactionDateRaw + "]");
  Logger.log("RAW AMOUNT VALUE: [" + amountRaw + "]");


  /***********************
   * RESOLVE TRANSACTION DATE
   * - Blank  -> use today's date (submission day)
   * - Filled -> use that date, and file the row into
   *             THAT month's sheet
   ***********************/

  let transactionDate;

  if (!transactionDateRaw || String(transactionDateRaw).trim() === "") {

    transactionDate = e.timestamp || new Date();

  } else {

    transactionDate = parseTransactionDate_(transactionDateRaw);

    if (!transactionDate) {
      throw new Error(
        "Invalid transaction date: [" + transactionDateRaw + "]"
      );
    }
  }


  /***********************
   * VALIDATE AMOUNT
   ***********************/

  const cleanedAmount =
    String(amountRaw)
      .replace(/,/g, "")
      .replace(/[₹]/g, "")
      .trim();

  const amount = parseFloat(cleanedAmount);

  if (cleanedAmount === "" || isNaN(amount)) {
    throw new Error(
      "Invalid amount: [" + amountRaw + "] " +
      "(cleaned value was: [" + cleanedAmount + "])"
    );
  }


  /***********************
   * DETERMINE MONTH
   * (based on transactionDate, so a past-dated entry
   *  always lands in that month's own sheet)
   ***********************/

  const monthSheet =
    getMonthSheetName_(transactionDate);

  Logger.log("RESOLVED MONTH SHEET: " + monthSheet);


  /***********************
   * CREATE MONTHLY SHEET
   ***********************/

  const transactionSheet =
    createMonthlyTransactionSheet_(
      ss,
      monthSheet
    );


  /***********************
   * CREATE MONTHLY ITEM SHEET
   ***********************/

  createMonthlyItemSheet_(
    ss,
    monthSheet
  );


  /***********************
   * CREATE TRANSACTION ID
   ***********************/

  const transactionId =
    generateTransactionId_(
      transactionDate,
      transactionSheet
    );


  /***********************
   * ENTRY TIMESTAMP
   ***********************/

  const entryTimestamp =
    new Date();


  /***********************
   * ADD TRANSACTION
   ***********************/

  transactionSheet.appendRow([

    transactionId,

    transactionDate,

    entryTimestamp,

    transactionType,

    amount,

    category,

    description,

    fromTo,

    account,

    paymentMethod,

    expenseNature,

    fixedVariable,

    recurring,

    invoice,

    notes,

    "Recorded"

  ]);


  /***********************
   * FORMAT NEW ROW
   ***********************/

  const lastRow =
    transactionSheet.getLastRow();

  transactionSheet
    .getRange(lastRow, 2)
    .setNumberFormat(CONFIG.DATE_FORMAT);

  transactionSheet
    .getRange(lastRow, 3)
    .setNumberFormat(CONFIG.DATETIME_FORMAT);

  transactionSheet
    .getRange(lastRow, 5)
    .setNumberFormat(
      '₹#,##0.00'
    );


  /***********************
   * UPDATE DASHBOARD
   ***********************/

  updateDashboard_(ss);
  refreshFilteredDashboard_(ss);

}


/*******************************************************
 * MONTHLY TRANSACTION SHEET
 *******************************************************/

function createMonthlyTransactionSheet_(
  ss,
  sheetName
) {

  let sheet =
    ss.getSheetByName(sheetName);

  if (sheet) {
    return sheet;
  }


  sheet =
    ss.insertSheet(sheetName);


  const headers = [

    "Transaction ID",

    "Transaction Date",

    "Entry Timestamp",

    "Transaction Type",

    "Amount (₹)",

    "Category",

    "Description",

    "From / To",

    "Account / Wallet",

    "Payment Method",

    "Expense Nature",

    "Fixed / Variable",

    "Recurring",

    "Invoice / Bill",

    "Notes",

    "Status"

  ];


  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);


  /***********************
   * HEADER FORMAT
   ***********************/

  formatHeader_(
    sheet.getRange(
      1,
      1,
      1,
      headers.length
    )
  );


  sheet.setFrozenRows(1);


  /***********************
   * COLUMN WIDTHS
   ***********************/

  const widths = [
    145,
    110,
    150,
    130,
    110,
    150,
    200,
    160,
    150,
    130,
    130,
    130,
    130,
    220,
    220,
    100
  ];


  widths.forEach(
    function(width, index) {

      sheet.setColumnWidth(
        index + 1,
        width
      );

    }
  );


  /***********************
   * FILTER
   ***********************/

  sheet
    .getRange(
      1,
      1,
      Math.max(sheet.getMaxRows(), 2),
      headers.length
    )
    .createFilter();


  return sheet;
}


/*******************************************************
 * MONTHLY INVOICE ITEM SHEET
 *******************************************************/

function createMonthlyItemSheet_(
  ss,
  monthName
) {

  const sheetName =
    monthName +
    CONFIG.ITEM_SHEET_SUFFIX;

  let sheet =
    ss.getSheetByName(sheetName);


  if (sheet) {
    return sheet;
  }


  sheet =
    ss.insertSheet(sheetName);


  const headers = [

    "Transaction ID",

    "Invoice ID",

    "Vendor",

    "Product / Item",

    "Quantity",

    "Unit",

    "Unit Price (₹)",

    "Discount (₹)",

    "Tax (₹)",

    "Item Amount (₹)",

    "Invoice Date",

    "Invoice Number",

    "OCR Status",

    "Verified",

    "Notes"

  ];


  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([headers]);


  formatHeader_(
    sheet.getRange(
      1,
      1,
      1,
      headers.length
    )
  );


  sheet.setFrozenRows(1);


  return sheet;
}


/*******************************************************
 * SETTINGS SHEET
 *******************************************************/

function createSettingsSheet_(ss) {

  let sheet =
    ss.getSheetByName(
      CONFIG.SETTINGS
    );


  if (sheet) {
    return sheet;
  }


  sheet =
    ss.insertSheet(
      CONFIG.SETTINGS
    );


  const data = [

    ["Setting", "Value"],

    ["Application Name",
      CONFIG.APP_NAME],

    ["Currency",
      CONFIG.CURRENCY],

    ["Date Format",
      CONFIG.DATE_FORMAT],

    ["Monthly Sheet Format",
      CONFIG.MONTH_FORMAT],

    ["Invoice OCR",
      "Future Module"],

    ["Google Pay Reconciliation",
      "Future Module"],

    ["Bank Statement Import",
      "Future Module"],

    ["Budget Management",
      "Future Module"],

    ["AI Spending Analysis",
      "Future Module"]

  ];


  sheet
    .getRange(
      1,
      1,
      data.length,
      2
    )
    .setValues(data);


  formatHeader_(
    sheet.getRange(
      1,
      1,
      1,
      2
    )
  );


  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 300);
}


/*******************************************************
 * CATEGORIES SHEET
 *******************************************************/

function createCategoriesSheet_(ss) {

  let sheet =
    ss.getSheetByName(
      CONFIG.CATEGORIES
    );


  if (sheet) {
    return sheet;
  }


  sheet =
    ss.insertSheet(
      CONFIG.CATEGORIES
    );


  const categories =
    getDefaultCategories_();


  sheet
    .getRange(1, 1)
    .setValue("Category");


  sheet
    .getRange(
      2,
      1,
      categories.length,
      1
    )
    .setValues(
      categories.map(
        function(item) {
          return [item];
        }
      )
    );


  formatHeader_(
    sheet.getRange(1, 1)
  );


  sheet.setColumnWidth(
    1,
    250
  );
}


/*******************************************************
 * ACCOUNTS SHEET
 *******************************************************/

function createAccountsSheet_(ss) {

  let sheet =
    ss.getSheetByName(
      CONFIG.ACCOUNTS
    );


  if (sheet) {
    return sheet;
  }


  sheet =
    ss.insertSheet(
      CONFIG.ACCOUNTS
    );


  const accounts =
    getDefaultAccounts_();


  sheet
    .getRange(1, 1)
    .setValue("Account / Wallet");


  sheet
    .getRange(
      2,
      1,
      accounts.length,
      1
    )
    .setValues(
      accounts.map(
        function(item) {
          return [item];
        }
      )
    );


  formatHeader_(
    sheet.getRange(1, 1)
  );


  sheet.setColumnWidth(
    1,
    250
  );
}


/*******************************************************
 * DASHBOARD
 *******************************************************/

function createDashboardSheet_(ss) {

  let sheet =
    ss.getSheetByName(
      CONFIG.DASHBOARD
    );


  if (sheet) {
    return sheet;
  }


  sheet =
    ss.insertSheet(
      CONFIG.DASHBOARD,
      0
    );


  sheet
    .getRange("A1:H1")
    .merge();


  sheet
    .getRange("A1")
    .setValue(
      "💰 MY FINANCE TRACKER"
    );


  sheet
    .getRange("A1")
    .setFontSize(20)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");


  sheet
    .getRange("A3:B3")
    .setValues([
      [
        "Financial Summary",
        "Amount (₹)"
      ]
    ]);


  formatHeader_(
    sheet.getRange("A3:B3")
  );


  const summary = [

    ["Total Money Received", ""],

    ["Total Expenses", ""],

    ["Savings", ""],

    ["Savings %", ""],

    ["Total Investment", ""],

    ["Money Lent", ""],

    ["Money Borrowed", ""]

  ];


  sheet
    .getRange(
      4,
      1,
      summary.length,
      2
    )
    .setValues(summary);


  sheet.setColumnWidth(
    1,
    250
  );

  sheet.setColumnWidth(
    2,
    160
  );


  sheet
    .getRange("A13:B13")
    .setValues([
      [
        "Category",
        "Expense (₹)"
      ]
    ]);


  formatHeader_(
    sheet.getRange("A13:B13")
  );


  sheet
    .getRange("D3:E3")
    .setValues([
      [
        "Account / Wallet",
        "Amount (₹)"
      ]
    ]);


  formatHeader_(
    sheet.getRange("D3:E3")
  );


  sheet
    .getRange("D13:E13")
    .setValues([
      [
        "Expense Nature",
        "Amount (₹)"
      ]
    ]);


  formatHeader_(
    sheet.getRange("D13:E13")
  );


  sheet
    .getRange("G3:H3")
    .setValues([
      [
        "Payment Method",
        "Amount (₹)"
      ]
    ]);


  formatHeader_(
    sheet.getRange("G3:H3")
  );


  return sheet;
}


/*******************************************************
 * DASHBOARD UPDATE
 *******************************************************/

function updateDashboard_(ss) {

  const dashboard =
    ss.getSheetByName(
      CONFIG.DASHBOARD
    );


  if (!dashboard) {
    return;
  }


  const sheets =
    ss.getSheets();


  let totalReceived = 0;
  let totalExpenses = 0;
  let totalInvestment = 0;
  let totalLent = 0;
  let totalBorrowed = 0;


  const categoryTotals = {};
  const accountTotals = {};
  const natureTotals = {};
  const paymentTotals = {};


  sheets.forEach(
    function(sheet) {

      const name =
        sheet.getName();


      if (
        name === CONFIG.DASHBOARD ||
        name === CONFIG.SETTINGS ||
        name === CONFIG.CATEGORIES ||
        name === CONFIG.ACCOUNTS ||
        name.endsWith(
          CONFIG.ITEM_SHEET_SUFFIX
        )
      ) {
        return;
      }


      if (
        !/^[A-Z][a-z]{2}-\d{4}$/.test(name)
      ) {
        return;
      }


      const lastRow =
        sheet.getLastRow();


      if (lastRow < 2) {
        return;
      }


      const values =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            16
          )
          .getValues();


      values.forEach(
        function(row) {

          const type =
            row[3];

          const amount =
            Number(row[4]) || 0;

          const category =
            row[5] || "Other";

          const account =
            row[8] || "Other";

          const payment =
            row[9] || "Other";

          const nature =
            row[10] || "Not Applicable";


          if (type === "Money Received") {
            totalReceived += amount;
          }


          if (type === "Expense") {

            totalExpenses += amount;

            categoryTotals[category] =
              (categoryTotals[category] || 0)
              + amount;
          }


          if (type === "Investment") {
            totalInvestment += amount;
          }


          if (type === "Money Lent") {
            totalLent += amount;
          }


          if (type === "Money Borrowed") {
            totalBorrowed += amount;
          }


          accountTotals[account] =
            (accountTotals[account] || 0)
            + amount;


          paymentTotals[payment] =
            (paymentTotals[payment] || 0)
            + amount;


          if (
            nature !== "Not Applicable" &&
            nature !== ""
          ) {

            natureTotals[nature] =
              (natureTotals[nature] || 0)
              + amount;
          }

        }
      );

    }
  );


  const savings =
    totalReceived -
    totalExpenses -
    totalInvestment;


  const savingsPercent =
    totalReceived === 0
      ? 0
      : savings / totalReceived;


  /***********************
   * SUMMARY
   ***********************/

  dashboard
    .getRange("B4:B10")
    .setValues([

      [totalReceived],

      [totalExpenses],

      [savings],

      [savingsPercent],

      [totalInvestment],

      [totalLent],

      [totalBorrowed]

    ]);


  dashboard
    .getRange("B4:B10")
    .setNumberFormat(
      '₹#,##0.00'
    );


  dashboard
    .getRange("B7")
    .setNumberFormat(
      "0.00%"
    );


  /***********************
   * CATEGORY TABLE
   ***********************/

  clearDashboardArea_(
    dashboard,
    "A14:B100"
  );


  const categoryRows =
    Object.keys(categoryTotals)
      .map(
        function(key) {
          return [
            key,
            categoryTotals[key]
          ];
        }
      );


  if (categoryRows.length > 0) {

    dashboard
      .getRange(
        14,
        1,
        categoryRows.length,
        2
      )
      .setValues(
        categoryRows
      );


    dashboard
      .getRange(
        14,
        2,
        categoryRows.length,
        1
      )
      .setNumberFormat(
        '₹#,##0.00'
      );
  }


  /***********************
   * ACCOUNT TABLE
   ***********************/

  clearDashboardArea_(
    dashboard,
    "D4:E100"
  );


  dashboard
    .getRange("D3:E3")
    .setValues([
      [
        "Account / Wallet",
        "Amount (₹)"
      ]
    ]);


  formatHeader_(
    dashboard.getRange("D3:E3")
  );


  const accountRows =
    Object.keys(accountTotals)
      .map(
        function(key) {
          return [
            key,
            accountTotals[key]
          ];
        }
      );


  if (accountRows.length > 0) {

    dashboard
      .getRange(
        4,
        4,
        accountRows.length,
        2
      )
      .setValues(
        accountRows
      );


    dashboard
      .getRange(
        4,
        5,
        accountRows.length,
        1
      )
      .setNumberFormat(
        '₹#,##0.00'
      );
  }


  /***********************
   * NATURE TABLE
   ***********************/

  clearDashboardArea_(
    dashboard,
    "D14:E100"
  );


  const natureRows =
    Object.keys(natureTotals)
      .map(
        function(key) {
          return [
            key,
            natureTotals[key]
          ];
        }
      );


  if (natureRows.length > 0) {

    dashboard
      .getRange(
        14,
        4,
        natureRows.length,
        2
      )
      .setValues(
        natureRows
      );


    dashboard
      .getRange(
        14,
        5,
        natureRows.length,
        1
      )
      .setNumberFormat(
        '₹#,##0.00'
      );
  }


  /***********************
   * PAYMENT METHOD
   ***********************/

  clearDashboardArea_(
    dashboard,
    "G4:H100"
  );


  const paymentRows =
    Object.keys(paymentTotals)
      .map(
        function(key) {
          return [
            key,
            paymentTotals[key]
          ];
        }
      );


  if (paymentRows.length > 0) {

    dashboard
      .getRange(
        4,
        7,
        paymentRows.length,
        2
      )
      .setValues(
        paymentRows
      );


    dashboard
      .getRange(
        4,
        8,
        paymentRows.length,
        1
      )
      .setNumberFormat(
        '₹#,##0.00'
      );
  }


  dashboard.autoResizeColumns(
    1,
    8
  );
}


/*******************************************************
 * DASHBOARD FILTERS
 * (Year + Category dropdowns, in columns J:K, kept
 *  clear of the existing summary tables in columns A:H)
 *******************************************************/

function getTransactionSheets_(ss) {

  return ss.getSheets().filter(function(sheet) {
    return /^[A-Za-z]{3}-\d{4}$/.test(sheet.getName());
  });
}


function getYearFromSheetName_(name) {

  const match = name.match(/-(\d{4})$/);
  return match ? match[1] : null;
}


function setupDashboardFilters_(ss) {

  const dashboard = ss.getSheetByName(CONFIG.DASHBOARD);

  if (!dashboard) {
    return;
  }

  // Header
  dashboard.getRange("J1:K1").merge();
  dashboard.getRange("J1")
    .setValue("🔍 FILTERS")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#4a86e8")
    .setFontColor("#ffffff");

  dashboard.getRange("J2").setValue("Select Year:");
  dashboard.getRange("J3").setValue("Select Category:");
  dashboard.getRange("J2:J3").setFontWeight("bold");

  // ----- Year dropdown -----
  const yearsFound = {};

  getTransactionSheets_(ss).forEach(function(sheet) {
    const y = getYearFromSheetName_(sheet.getName());
    if (y) {
      yearsFound[y] = true;
    }
  });

  let yearList = Object.keys(yearsFound).sort();

  if (yearList.length === 0) {
    yearList = [String(new Date().getFullYear())];
  }

  const yearOptions = ["All Years"].concat(yearList);

  const yearCell = dashboard.getRange("K2");

  const yearRule =
    SpreadsheetApp.newDataValidation()
      .requireValueInList(yearOptions, true)
      .setAllowInvalid(false)
      .build();

  yearCell.setDataValidation(yearRule);

  if (!yearCell.getValue()) {
    yearCell.setValue("All Years");
  }

  // ----- Category dropdown -----
  const categoryOptions =
    ["All Categories"].concat(getDefaultCategories_());

  const categoryCell = dashboard.getRange("K3");

  const categoryRule =
    SpreadsheetApp.newDataValidation()
      .requireValueInList(categoryOptions, true)
      .setAllowInvalid(false)
      .build();

  categoryCell.setDataValidation(categoryRule);

  if (!categoryCell.getValue()) {
    categoryCell.setValue("All Categories");
  }

  dashboard.setColumnWidth(10, 150); // column J
  dashboard.setColumnWidth(11, 150); // column K
}


/*******************************************************
 * onEdit - SIMPLE TRIGGER
 * Fires automatically whenever a cell is edited by hand.
 * When the Year (K2) or Category (K3) filter on the
 * Dashboard changes, the filtered views refresh instantly.
 *******************************************************/

function onEdit(e) {

  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.DASHBOARD) {
    return;
  }

  const editedCell = e.range.getA1Notation();

  if (editedCell === "K2" || editedCell === "K3") {
    refreshFilteredDashboard_(sheet.getParent());
  }
}


/*******************************************************
 * REFRESH FILTERED DASHBOARD VIEWS
 * Reads the Year / Category selection and rebuilds:
 *   - Filtered Summary (Total Expense / Income / Net)
 *   - Monthly Trend (Jan-Dec) for the selected Year
 *     and Category
 *   - Yearly Trend (across all years) for the selected
 *     Category, so long-term trend is visible regardless
 *     of the Year filter
 *******************************************************/

function refreshFilteredDashboard_(ss) {

  const dashboard = ss.getSheetByName(CONFIG.DASHBOARD);

  if (!dashboard) {
    return;
  }

  const selectedYear =
    dashboard.getRange("K2").getValue() || "All Years";

  const selectedCategory =
    dashboard.getRange("K3").getValue() || "All Categories";

  const monthNames = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
  ];

  const monthlyTotals = {};   // monthIndex -> expense total (Year + Category filtered)
  const yearlyTotals = {};    // year -> expense total (Category filtered only)

  let filteredExpense = 0;
  let filteredIncome = 0;

  getTransactionSheets_(ss).forEach(function(sheet) {

    const name = sheet.getName();
    const year = getYearFromSheetName_(name);
    const monthAbbr = name.split("-")[0];
    const monthIndex = monthNames.indexOf(monthAbbr);

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return;
    }

    const values =
      sheet.getRange(2, 1, lastRow - 1, 16).getValues();

    values.forEach(function(row) {

      const type = row[3];
      const amount = Number(row[4]) || 0;
      const category = row[5] || "Other";

      const categoryMatches =
        (selectedCategory === "All Categories" || category === selectedCategory);

      const yearMatches =
        (selectedYear === "All Years" || year === selectedYear);

      // Yearly trend spans ALL years, respecting only the Category filter
      if (type === "Expense" && categoryMatches) {
        yearlyTotals[year] = (yearlyTotals[year] || 0) + amount;
      }

      // Filtered Summary + Monthly Trend respect BOTH filters
      if (yearMatches && categoryMatches) {

        if (type === "Expense") {

          filteredExpense += amount;

          if (monthIndex >= 0) {
            monthlyTotals[monthIndex] =
              (monthlyTotals[monthIndex] || 0) + amount;
          }
        }

        if (type === "Money Received") {
          filteredIncome += amount;
        }
      }
    });
  });

  /***********************
   * FILTERED SUMMARY
   ***********************/

  dashboard.getRange("J5:K5").merge();
  dashboard.getRange("J5").setValue("Filtered Summary");
  formatHeader_(dashboard.getRange("J5:K5"));

  dashboard.getRange("J6").setValue("Total Expenses");
  dashboard.getRange("K6")
    .setValue(filteredExpense)
    .setNumberFormat('₹#,##0.00');

  dashboard.getRange("J7").setValue("Total Income");
  dashboard.getRange("K7")
    .setValue(filteredIncome)
    .setNumberFormat('₹#,##0.00');

  dashboard.getRange("J8").setValue("Net (Income - Expenses)");
  dashboard.getRange("K8")
    .setValue(filteredIncome - filteredExpense)
    .setNumberFormat('₹#,##0.00');

  /***********************
   * MONTHLY TREND TABLE
   ***********************/

  dashboard.getRange("J10:K10").merge();
  dashboard.getRange("J10").setValue("Monthly Trend (Expenses)");
  formatHeader_(dashboard.getRange("J10:K10"));

  dashboard.getRange("J11:K11")
    .setValues([["Month", "Expense (₹)"]]);
  formatHeader_(dashboard.getRange("J11:K11"));

  const monthRows = monthNames.map(function(m, idx) {
    return [m, monthlyTotals[idx] || 0];
  });

  dashboard.getRange(12, 10, monthRows.length, 2)
    .setValues(monthRows);

  dashboard.getRange(12, 11, monthRows.length, 1)
    .setNumberFormat('₹#,##0.00');

  /***********************
   * YEARLY TREND TABLE
   ***********************/

  const yearlyRow = 12 + monthRows.length + 2;

  dashboard.getRange(yearlyRow, 10, 1, 2).merge();
  dashboard.getRange(yearlyRow, 10)
    .setValue("Yearly Trend (Expenses)");
  formatHeader_(dashboard.getRange(yearlyRow, 10, 1, 2));

  dashboard.getRange(yearlyRow + 1, 10, 1, 2)
    .setValues([["Year", "Expense (₹)"]]);
  formatHeader_(dashboard.getRange(yearlyRow + 1, 10, 1, 2));

  const yearKeys = Object.keys(yearlyTotals).sort();

  const yearRows = yearKeys.map(function(y) {
    return [y, yearlyTotals[y]];
  });

  if (yearRows.length > 0) {

    dashboard.getRange(yearlyRow + 2, 10, yearRows.length, 2)
      .setValues(yearRows);

    dashboard.getRange(yearlyRow + 2, 11, yearRows.length, 1)
      .setNumberFormat('₹#,##0.00');
  }

  dashboard.autoResizeColumns(10, 11);
}


/*******************************************************
 * MANUAL FILTER REFRESH (menu-triggered fallback,
 * in case onEdit doesn't fire for any reason)
 *******************************************************/

function manualFilteredDashboardRefresh() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  setupDashboardFilters_(ss);
  refreshFilteredDashboard_(ss);

  SpreadsheetApp.getUi().alert(
    "Filtered dashboard refreshed."
  );
}


/*******************************************************
 * DEFAULT CATEGORIES
 *******************************************************/

function getDefaultCategories_() {

  return [

    "Food & Grocery",

    "Household",

    "Transport / Fuel",

    "Medical",

    "Shopping",

    "Bills & Utilities",

    "Education",

    "Entertainment",

    "EMI / Loan",

    "Investment",

    "Salary / Income",

    "Other"

  ];
}


/*******************************************************
 * DEFAULT ACCOUNTS
 *******************************************************/

function getDefaultAccounts_() {

  return [

    "Bank Account 1",

    "Bank Account 2",

    "Cash",

    "Credit Card",

    "Other"

  ];
}


/*******************************************************
 * GET FORM VALUE
 * (whitespace-tolerant title matching)
 *******************************************************/

function getFormValue_(
  namedValues,
  fieldName
) {

  if (!namedValues) {
    return "";
  }

  if (namedValues[fieldName] !== undefined) {

    const value = namedValues[fieldName];

    return Array.isArray(value)
      ? value.join(", ")
      : value;
  }

  // Fallback: match ignoring leading/trailing whitespace,
  // in case the question title has stray spaces.
  const keys = Object.keys(namedValues);

  for (let i = 0; i < keys.length; i++) {

    if (keys[i].trim() === fieldName.trim()) {

      const value = namedValues[keys[i]];

      return Array.isArray(value)
        ? value.join(", ")
        : value;
    }
  }

  return "";
}


/*******************************************************
 * DATE PARSER
 * Handles the M/D/YYYY format Google Forms sends for
 * this form (e.g. "8/31/2026" = 31 August 2026), plus
 * fallbacks. Builds the Date using local year/month/day
 * components so there is no UTC day-shift.
 *******************************************************/

function parseTransactionDate_(
  value
) {

  if (!value) {
    return null;
  }

  if (
    Object.prototype.toString.call(value)
    === "[object Date]"
  ) {
    return value;
  }

  const str = String(value).trim();

  // Primary format sent by this form: M/D/YYYY
  let match = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (match) {

    const month = parseInt(match[1], 10) - 1;
    const day = parseInt(match[2], 10);
    const year = parseInt(match[3], 10);

    return new Date(year, month, day);
  }

  // Fallback: yyyy-MM-dd
  match = str.match(/^(\d{4})-(\d{2})-(\d{2})/);

  if (match) {

    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1;
    const day = parseInt(match[3], 10);

    return new Date(year, month, day);
  }

  // Fallback: dd-mm-yyyy or dd-Mon-yyyy etc, let JS try
  const parsed = new Date(str);

  if (isNaN(parsed.getTime())) {
    return null;
  }

  return parsed;
}


/*******************************************************
 * MONTH SHEET NAME
 *******************************************************/

function getMonthSheetName_(
  date
) {

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    CONFIG.MONTH_FORMAT
  );

}


/*******************************************************
 * TRANSACTION ID
 *******************************************************/

function generateTransactionId_(
  date,
  sheet
) {

  const prefix =
    Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );


  const existingRows =
    Math.max(
      sheet.getLastRow() - 1,
      0
    );


  const sequence =
    String(existingRows + 1)
      .padStart(4, "0");


  return (
    "TXN-" +
    prefix +
    "-" +
    sequence
  );
}


/*******************************************************
 * SAVE CONFIGURATION
 *******************************************************/

function saveConfiguration_(
  ss,
  form
) {

  const properties =
    PropertiesService
      .getScriptProperties();


  properties.setProperty(
    "FINANCE_FORM_ID",
    form.getId()
  );


  properties.setProperty(
    "FINANCE_SHEET_ID",
    ss.getId()
  );


  properties.setProperty(
    "FINANCE_FORM_URL",
    form.getPublishedUrl()
  );


  properties.setProperty(
    "FINANCE_FORM_EDIT_URL",
    form.getEditUrl()
  );

}


/*******************************************************
 * BASE FORMATTING
 *******************************************************/

function formatBaseSheets_(ss) {

  const dashboard =
    ss.getSheetByName(
      CONFIG.DASHBOARD
    );


  if (dashboard) {

    dashboard.setFrozenRows(3);

    dashboard
      .getRange("A1:H100")
      .setVerticalAlignment(
        "middle"
      );

  }


  const settings =
    ss.getSheetByName(
      CONFIG.SETTINGS
    );


  if (settings) {
    settings.setFrozenRows(1);
  }


  const categories =
    ss.getSheetByName(
      CONFIG.CATEGORIES
    );


  if (categories) {
    categories.setFrozenRows(1);
  }


  const accounts =
    ss.getSheetByName(
      CONFIG.ACCOUNTS
    );


  if (accounts) {
    accounts.setFrozenRows(1);
  }
}


/*******************************************************
 * HEADER FORMAT
 *******************************************************/

function formatHeader_(
  range
) {

  range
    .setFontWeight("bold")
    .setHorizontalAlignment(
      "center"
    )
    .setVerticalAlignment(
      "middle"
    );


  range.setWrap(true);

}


/*******************************************************
 * CLEAR DASHBOARD AREA
 *******************************************************/

function clearDashboardArea_(
  sheet,
  range
) {

  sheet
    .getRange(range)
    .clearContent();

}


/*******************************************************
 * OPEN FORM
 *******************************************************/

function openFinanceForm() {

  const properties =
    PropertiesService
      .getScriptProperties();


  const url =
    properties.getProperty(
      "FINANCE_FORM_URL"
    );


  if (!url) {

    SpreadsheetApp.getUi().alert(
      "Finance Form has not been created yet.\n\n" +
      "Run setupFinanceTracker() first."
    );

    return;
  }


  const html =
    HtmlService.createHtmlOutput(
      '<script>' +
      'window.open("' +
      url +
      '","_blank");' +
      'google.script.host.close();' +
      '</script>'
    );


  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "Open Finance Form"
    );
}


/*******************************************************
 * CREATE MENU
 *******************************************************/

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu("💰 Finance Tracker")
    .addItem(
      "Open Finance Form",
      "openFinanceForm"
    )
    .addItem(
      "Refresh Dashboard",
      "manualDashboardRefresh"
    )
    .addItem(
      "Update Form Descriptions",
      "updateExistingFormDescriptions"
    )
    .addItem(
      "Refresh Filtered Dashboard",
      "manualFilteredDashboardRefresh"
    )
    .addToUi();

}


/*******************************************************
 * MANUAL DASHBOARD REFRESH
 *******************************************************/

function manualDashboardRefresh() {

  updateDashboard_(
    SpreadsheetApp.getActiveSpreadsheet()
  );


  SpreadsheetApp
    .getUi()
    .alert(
      "Dashboard refreshed successfully."
    );

}
