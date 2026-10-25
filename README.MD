# SaveIQ – AI Savings Goal Tracker 💰

**Save smarter. Track better. Achieve your financial goals.**

SaveIQ is an AI-powered savings goal tracking web application that helps users set financial goals, monitor their savings progress, record deposits, and receive personalized savings recommendations.

The application uses Google Apps Script as its backend, Google Sheets as its database, and the Groq API for AI-powered recommendations.

## ✨ Features

- 🎯 **Savings Goal Management** — Create, edit, and delete financial goals.
- 💰 **Savings Tracking** — Record deposits and track current savings against target amounts.
- 📜 **Transaction History** — View recorded deposits, amounts, dates, and transaction IDs.
- 📊 **Progress Dashboard** — Monitor total goals, target amounts, total savings, and overall progress.
- 🤖 **AI Savings Recommendations** — Get personalized recommendations through the Groq API.
- 📧 **Deadline Reminders** — Support for scheduled email reminders when configured.
- ☁️ **Cloud-Based Storage** — Store goal and transaction records in Google Sheets.
- 🌐 **Web Application** — Access the application through a deployed Google Apps Script web app.

## 🛠️ Technology Stack

| Component | Technology |
|---|---|
| Frontend | HTML5, CSS3, JavaScript |
| Backend | Google Apps Script |
| Database | Google Sheets |
| AI Integration | Groq API |
| Deployment | Google Apps Script Web App |
| Version Control | Git and GitHub |

## 🏗️ Project Architecture

```text
User
 |
 v
SaveIQ Web Interface
(HTML, CSS, JavaScript)
 |
 v
Google Apps Script
(Backend Logic)
 |
 +----------------------+
 |                      |
 v                      v
Google Sheets         Groq API
Goals and              AI Savings
Transactions           Recommendations
 |
 v
Dashboard Updates
```

### How It Works

1. Users create savings goals through the web interface.
2. The frontend sends requests to Google Apps Script using `google.script.run`.
3. The backend stores and retrieves goal information from Google Sheets.
4. Savings deposits update the corresponding goal and are recorded in the Transactions sheet.
5. The dashboard displays updated savings information and progress.
6. The AI recommendation feature sends relevant goal information to the Groq API through the backend.
7. Scheduled email reminders can notify users about approaching deadlines when the reminder trigger is configured.

### AI Integration — Groq API

SaveIQ integrates the Groq API to generate AI-powered savings recommendations based on a user's savings goal, current progress, target amount, and deadline.

**How it works:**
1. The user selects a savings goal.
2. The Google Apps Script backend prepares the relevant goal information.
3. The backend sends a request to the Groq API.
4. The AI-generated recommendations are returned to the application.

**Security:** The API key is stored in Google Apps Script Script Properties using the `GROQ_API_KEY` property. It is not included in the frontend code or committed to this repository.

**Requirements:** A valid Groq API key and access to the Groq API are required for AI recommendations.


## 📁 Project Structure

```text
SaveIQ/
│
├── Code.gs
│   └── Backend functions and Google Sheets integration
│
├── Index.html
│   └── Frontend interface, styling, and JavaScript
│
├── README.md
│   └── Project documentation
│
└── screenshots/
    └── Application screenshots
```

## 🗄️ Database Structure

SaveIQ uses a Google Spreadsheet with two sheets.

### 1. Goals

| Column | Description |
|---|---|
| GoalID | Unique identifier for a goal |
| GoalName | Name of the savings goal |
| TargetAmount | Target savings amount |
| InitialSavings | Amount already saved when the goal was created |
| CurrentSavings | Current accumulated savings |
| Deadline | Target completion date |
| Email | Email address for reminders |
| Purpose | Reason for the savings goal |
| CreatedAt | Goal creation timestamp |
| Status | Current goal status |
| LastAlertSent | Tracks the last reminder sent |

### 2. Transactions

| Column | Description |
|---|---|
| TransactionID | Unique transaction identifier |
| GoalID | ID of the associated savings goal |
| GoalName | Name of the associated goal |
| Amount | Deposit amount |
| TransactionDate | Date and time of the deposit |

## 🚀 Setup and Installation

### Prerequisites

- A Google account
- Google