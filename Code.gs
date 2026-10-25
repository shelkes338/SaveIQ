
const SPREADSHEET_ID = "1lQVUAveh5bTNs-A-r4XYi7Og1JKw5GrryFFGHSyCaCA";
const SHEET_NAME = "Goals";

// ============================================
// DATABASE CONNECTION
// ============================================

function getDatabase() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet "' + SHEET_NAME + '" was not found. Check the sheet tab name.'
    );
  }

  return sheet;
}

function testConnection() {
  const sheet = getDatabase();
  const headers = sheet.getRange(1, 1, 1, 11).getDisplayValues()[0];

  Logger.log("SaveIQ connected successfully!");
  Logger.log("Sheet name: " + sheet.getName());
  Logger.log("Column headings: " + headers.join(", "));

  return "Connection successful!";
}

function formatDateValue(value, format) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      format
    );
  }

  return String(value || "");
}


// ============================================
// RETRIEVE SAVINGS GOALS
// ============================================

function getGoals() {
  const sheet = getDatabase();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return [];
  }

  return data.slice(1)
    .filter(row => row[0] !== "")
    .map(row => ({
      id: String(row[0]),
      name: String(row[1] || ""),
      targetAmount: Number(row[2]) || 0,
      initialSavings: Number(row[3]) || 0,
      currentSavings: Number(row[4]) || 0,
      deadline: formatDateValue(row[5], "yyyy-MM-dd"),
      email: String(row[6] || ""),
      purpose: String(row[7] || ""),
      createdAt: formatDateValue(row[8], "yyyy-MM-dd HH:mm:ss"),
      status: String(row[9] || ""),
      lastAlertSent: formatDateValue(
        row[10],
        "yyyy-MM-dd HH:mm:ss"
      )
    }));
}


// ============================================
// CREATE A NEW SAVINGS GOAL
// ============================================

function addGoal(goal) {
  if (!goal || !goal.name || !String(goal.name).trim() || !goal.deadline) {
    throw new Error("Please enter a goal name and deadline.");
  }

  const target = Number(goal.targetAmount);
  const initial = Number(goal.initialSavings || 0);

  if (!Number.isFinite(target) || target <= 0) {
    throw new Error("Target amount must be greater than zero.");
  }

  if (!Number.isFinite(initial) || initial < 0 || initial > target) {
    throw new Error(
      "Initial savings must be between zero and the target."
    );
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(goal.deadline))) {
    throw new Error("Please enter a valid deadline.");
  }

  const deadline = new Date(goal.deadline + "T12:00:00");

  if (isNaN(deadline.getTime())) {
    throw new Error("Please enter a valid deadline.");
  }

  const sheet = getDatabase();
  const id = Utilities.getUuid();
  const now = new Date();
  const status = initial >= target ? "Completed" : "Active";

  sheet.appendRow([
    id,
    String(goal.name).trim(),
    target,
    initial,
    initial,
    deadline,
    String(goal.email || "").trim(),
    String(goal.purpose || "").trim(),
    now,
    status,
    ""
  ]);

  return {
    success: true,
    message: "Savings goal created successfully!",
    id: id
  };
}


// ============================================
// TRANSACTIONS SHEET
// ============================================

function getTransactionsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Transactions");

  if (!sheet) {
    sheet = ss.insertSheet("Transactions");
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "TransactionID",
      "GoalID",
      "GoalName",
      "Amount",
      "TransactionDate"
    ]);
  }

  return sheet;
}


// ============================================
// ADD SAVINGS AND RECORD TRANSACTION
// ============================================

function addSavings(goalId, amount) {
  const id = String(goalId || "").trim();
  const savingsAmount = Number(amount);

  if (!id) {
    throw new Error("Invalid savings goal.");
  }

  if (!Number.isFinite(savingsAmount) || savingsAmount <= 0) {
    throw new Error("Enter a savings amount greater than zero.");
  }

  const goalsSheet = getDatabase();
  const lastRow = goalsSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("No savings goals found.");
  }

  const ids = goalsSheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues();

  let goalRow = -1;

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      goalRow = i + 2;
      break;
    }
  }

  if (goalRow === -1) {
    throw new Error("Savings goal not found.");
  }

  const goalData = goalsSheet
    .getRange(goalRow, 1, 1, 11)
    .getValues()[0];

  const goalName = String(goalData[1] || "");
  const targetAmount = Number(goalData[2]) || 0;
  const currentSavings = Number(goalData[4]) || 0;

  if (targetAmount <= 0) {
    throw new Error("This goal has an invalid target amount.");
  }

  if (currentSavings + savingsAmount > targetAmount) {
    throw new Error(
      "This amount exceeds your remaining target of ₹" +
      (targetAmount - currentSavings).toFixed(2)
    );
  }

  const updatedSavings = currentSavings + savingsAmount;
  const status = updatedSavings >= targetAmount
    ? "Completed"
    : "Active";

  const transactionsSheet = getTransactionsSheet();

  // Record the transaction before updating the goal.
  // If writing the transaction fails, the savings total
  // will not be updated by this function.
  transactionsSheet.appendRow([
    Utilities.getUuid(),
    id,
    goalName,
    savingsAmount,
    new Date()
  ]);

  try {
    goalsSheet.getRange(goalRow, 5).setValue(updatedSavings);
    goalsSheet.getRange(goalRow, 10).setValue(status);
    SpreadsheetApp.flush();
  } catch (error) {
    // Remove the transaction if updating the goal fails.
    const transactionRow = transactionsSheet.getLastRow();

    if (transactionRow > 1) {
      transactionsSheet.deleteRow(transactionRow);
    }

    throw new Error(
      "Could not update the savings goal. Please try again."
    );
  }

  return {
    success: true,
    message: "Savings added and transaction recorded successfully!",
    currentSavings: updatedSavings,
    status: status
  };
}


// ============================================
// OPEN THE SAVEIQ DASHBOARD
// ============================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("SaveIQ - AI Savings Goal Tracker");
}


// ============================================
// DEBUG AND TEST FUNCTIONS
// ============================================

function debugGoals() {
  const goals = getGoals();

  Logger.log("Total goals found: " + goals.length);
  Logger.log(JSON.stringify(goals));

  return goals.length;
}

function testAddGoal() {
  const result = addGoal({
    name: "Laptop Purchase Test",
    targetAmount: 50000,
    initialSavings: 5000,
    deadline: "2027-06-30",
    email: "",
    purpose: "Test savings goal"
  });

  Logger.log(JSON.stringify(result));
}


// ============================================
// DELETE A SAVINGS GOAL
// ============================================

function deleteGoal(goalId) {
  const id = String(goalId || "").trim();

  if (!id) {
    throw new Error("Invalid savings goal ID.");
  }

  const sheet = getDatabase();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("No savings goals found.");
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.deleteRow(i + 2);

      return {
        success: true,
        message: "Savings goal deleted successfully."
      };
    }
  }

  throw new Error("Savings goal not found.");
}


// ============================================
// EDIT A SAVINGS GOAL
// ============================================

function editGoal(goalId, updatedGoal) {
  const id = String(goalId || "").trim();

  if (!id || !updatedGoal) {
    throw new Error("Invalid goal information.");
  }

  const name = String(updatedGoal.name || "").trim();
  const target = Number(updatedGoal.targetAmount);
  const deadlineText = String(updatedGoal.deadline || "");

  if (!name || name.length > 100) {
    throw new Error(
      "Goal name is required and must be under 100 characters."
    );
  }

  if (!Number.isFinite(target) || target <= 0) {
    throw new Error("Target amount must be greater than zero.");
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(deadlineText)) {
    throw new Error("Please enter a valid deadline.");
  }

  const deadline = new Date(deadlineText + "T12:00:00");

  if (isNaN(deadline.getTime())) {
    throw new Error("Please enter a valid deadline.");
  }

  const sheet = getDatabase();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("No savings goals found.");
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getDisplayValues();

  let goalRow = -1;

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      goalRow = i + 2;
      break;
    }
  }

  if (goalRow === -1) {
    throw new Error("Savings goal not found.");
  }

  const currentSavings =
    Number(sheet.getRange(goalRow, 5).getValue()) || 0;

  if (target < currentSavings) {
    throw new Error(
      "Target cannot be lower than your current savings of ₹" +
      currentSavings.toFixed(2)
    );
  }

  // Update columns B, C, F, G and H.
  sheet.getRange(goalRow, 2).setValue(name);
  sheet.getRange(goalRow, 3).setValue(target);
  sheet.getRange(goalRow, 6).setValue(deadline);
  sheet.getRange(goalRow, 7).setValue(
    String(updatedGoal.email || "").trim()
  );
  sheet.getRange(goalRow, 8).setValue(
    String(updatedGoal.purpose || "").trim()
  );

  const status = currentSavings >= target
    ? "Completed"
    : "Active";

  sheet.getRange(goalRow, 10).setValue(status);

  return {
    success: true,
    message: "Savings goal updated successfully."
  };
}


// ============================================
// AI SAVINGS RECOMMENDATIONS — GROQ
// ============================================

function getAIRecommendations(goal) {
  if (!goal || typeof goal !== "object") {
    throw new Error("Invalid goal information.");
  }

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GROQ_API_KEY");

  if (!apiKey) {
    throw new Error(
      "Missing GROQ_API_KEY in Project Settings > Script Properties."
    );
  }

  const name = String(goal.name || "").trim().slice(0, 100);
  const target = Number(goal.targetAmount);
  const saved = Number(goal.currentSavings);
  const deadlineText = String(goal.deadline || "");
  const purpose = String(goal.purpose || "")
    .trim()
    .slice(0, 300);

  if (
    !name ||
    !Number.isFinite(target) ||
    !Number.isFinite(saved) ||
    target <= 0 ||
    saved < 0 ||
    saved > target
  ) {
    throw new Error("Invalid goal name or savings amount.");
  }

  if (!/^\d{4}-\d{2}-\d{2}$/.test(deadlineText)) {
    throw new Error("Please provide a valid deadline.");
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const deadline = new Date(deadlineText + "T12:00:00");

  if (isNaN(deadline.getTime())) {
    throw new Error("Invalid deadline.");
  }

  const daysLeft = Math.ceil(
    (deadline.getTime() - today.getTime()) /
    (24 * 60 * 60 * 1000)
  );

  const remaining = Math.max(0, target - saved);
  const daily = daysLeft > 0 ? remaining / daysLeft : remaining;

  const prompt = [
    "You are SaveIQ, a practical personal savings coach.",
    "Use Indian rupees and simple language.",
    "Give general budgeting advice, not investment advice.",
    "Never recommend gambling, loans, or risky investments.",
    "Do not invent the user's income or expenses.",
    "",
    "Goal: " + name,
    "Purpose: " + (purpose || "Not specified"),
    "Target: INR " + target.toFixed(2),
    "Saved: INR " + saved.toFixed(2),
    "Remaining: INR " + remaining.toFixed(2),
    "Deadline: " + deadlineText,
    "Days remaining: " + daysLeft,
    "Daily savings required: INR " + daily.toFixed(2),
    "",
    "Provide:",
    "1. A short goal assessment.",
    "2. Daily and weekly savings suggestions.",
    "3. Three realistic money-saving ideas.",
    "4. A fallback plan if the deadline is difficult.",
    "Keep the response concise and encouraging."
  ].join("\n");

  const response = UrlFetchApp.fetch(
    "https://api.groq.com/openai/v1/chat/completions",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: "Bearer " + apiKey
      },
      payload: JSON.stringify({
        model: "openai/gpt-oss-20b",
        messages: [
          {
            role: "system",
            content: "You are the AI savings advisor for SaveIQ."
          },
          {
            role: "user",
            content: prompt
          }
        ],
        temperature: 0.5,
        max_completion_tokens: 700
      }),
      muteHttpExceptions: true
    }
  );

  const responseStatus = response.getResponseCode();
  const body = response.getContentText();

  if (responseStatus < 200 || responseStatus >= 300) {
    console.error("Groq error " + responseStatus + ": " + body);

    throw new Error(
      responseStatus === 429
        ? "AI usage limit reached. Try again later."
        : "AI request failed. Check the API key and Groq model availability."
    );
  }

  const result = JSON.parse(body);
  const advice = result.choices &&
    result.choices[0] &&
    result.choices[0].message &&
    result.choices[0].message.content;

  if (!advice) {
    throw new Error("The AI returned no recommendation.");
  }

  return {
    success: true,
    advice: advice,
    dailySavings: Math.round(daily * 100) / 100,
    remaining: Math.round(remaining * 100) / 100,
    daysLeft: daysLeft
  };
}


// ============================================
// DAILY EMAIL DEADLINE REMINDERS
// ============================================

function sendDailyDeadlineReminders() {
  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("Goals");

  if (!sheet) {
    throw new Error('The "Goals" sheet was not found.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log("No savings goals found.");
    return;
  }

  const timezone = Session.getScriptTimeZone();
  const todayString = Utilities.formatDate(
    new Date(),
    timezone,
    "yyyy-MM-dd"
  );

  const today = dateFromYmd_(todayString);

  const rows = sheet
    .getRange(2, 1, lastRow - 1, 11)
    .getValues();

  let sentCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    const goalId = row[0];
    const goalName = row[1];
    const target = Number(row[2]) || 0;
    const currentSavings = Number(row[4]) || 0;
    const deadlineValue = row[5];
    const email = String(row[6] || "").trim();
    const status = String(row[9] || "").trim().toLowerCase();
    const lastAlertSent = row[10];

    if (!email) {
      continue;
    }

    if (
      status === "completed" ||
      (target > 0 && currentSavings >= target)
    ) {
      continue;
    }

    if (!deadlineValue) {
      continue;
    }

    const deadlineString = getGoalDateString_(
      deadlineValue,
      timezone
    );

    if (!deadlineString) {
      Logger.log("Invalid deadline for goal: " + goalName);
      continue;
    }

    const deadline = dateFromYmd_(deadlineString);

    const daysRemaining = Math.round(
      (deadline.getTime() - today.getTime()) /
      (24 * 60 * 60 * 1000)
    );

    if (daysRemaining !== 3) {
      continue;
    }

    const lastSentString = lastAlertSent
      ? getGoalDateString_(lastAlertSent, timezone)
      : "";

    if (lastSentString === todayString) {
      continue;
    }

    const remaining = Math.max(0, target - currentSavings);

    const subject =
      "SaveIQ Reminder: " + goalName + " is due in 3 days";

    const body =
      "Hello!\n\n" +
      "This is your SaveIQ savings goal reminder.\n\n" +
      "Your goal deadline is approaching!\n\n" +
      "Goal: " + goalName + "\n" +
      "Target amount: " + formatRupees_(target) + "\n" +
      "Current savings: " + formatRupees_(currentSavings) + "\n" +
      "Amount remaining: " + formatRupees_(remaining) + "\n" +
      "Deadline: " + deadlineString + "\n\n" +
      (remaining > 0
        ? "To reach your target in the next 3 days, " +
          "you would need to save approximately " +
          formatRupees_(remaining / 3) +
          " per day.\n\n"
        : "Your savings target has been reached. " +
          "Please check your goal status.\n\n") +
      "Keep going — every saving brings you closer to your goal!\n\n" +
      "Regards,\nSaveIQ";

    if (MailApp.getRemainingDailyQuota() < 1) {
      Logger.log("Daily email quota exhausted. Stopping.");
      return;
    }

    MailApp.sendEmail(email, subject, body);

    sheet.getRange(i + 2, 11).setValue(new Date());

    sentCount++;

    Logger.log(
      "Reminder sent for goal " + goalId + " to " + email
    );
  }

  Logger.log(
    "Daily reminder check completed. Emails sent: " + sentCount
  );
}


// ============================================
// EMAIL REMINDER DATE HELPERS
// ============================================

function getGoalDateString_(value, timezone) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      timezone,
      "yyyy-MM-dd"
    );
  }

  const text = String(value || "").trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (match) {
    return text;
  }

  const parsed = new Date(text);

  if (isNaN(parsed.getTime())) {
    return "";
  }

  return Utilities.formatDate(
    parsed,
    timezone,
    "yyyy-MM-dd"
  );
}

function dateFromYmd_(ymd) {
  const parts = ymd.split("-").map(Number);

  return new Date(
    parts[0],
    parts[1] - 1,
    parts[2]
  );
}

function formatRupees_(amount) {
  return "Rs. " + Number(amount || 0).toLocaleString(
    "en-IN",
    {
      maximumFractionDigits: 2
    }
  );
}


// ============================================
// SET UP DAILY EMAIL REMINDER TRIGGER
// Run this function once manually.
// ============================================

function setupDailyReminderTrigger() {
  const functionName = "sendDailyDeadlineReminders";
  const existingTriggers = ScriptApp.getProjectTriggers();

  existingTriggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log(
    "Daily SaveIQ reminder trigger created for the 9 AM hour."
  );
}


function getTransactions() {
  const sheet = getTransactionsSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  return data
    .filter(row => row[0] !== "")
    .reverse()
    .map(row => ({
      transactionId: String(row[0] || ""),
      goalId: String(row[1] || ""),
      goalName: String(row[2] || ""),
      amount: Number(row[3]) || 0,
      transactionDate: formatDateValue(
        row[4],
        "dd MMM yyyy, hh:mm a"
      )
    }));
}
