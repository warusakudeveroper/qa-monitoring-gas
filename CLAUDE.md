# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script (GAS) project for automatically monitoring a Q&A spreadsheet and sending notifications via Discord webhooks and email when new questions or answers are detected.

## Key Architecture Components

### Core Script Structure
- **Container-bound GAS script** (`qa_monitoring.gs`) that runs within a Google Spreadsheet
- **Time-based triggers** execute `monitorQASheet()` function hourly to check for changes
- **Script Properties** store configuration (webhooks, email addresses, processed items)
- **Dual notification system** supporting both Discord and Email channels

### Data Flow
1. Monitor function checks spreadsheet data rows
2. Compares against processed items stored in Script Properties
3. Detects new questions (when columns A-H are filled) and answers (when columns I-J are added)
4. Sends formatted notifications and updates processed lists
5. Creates backups in separate spreadsheet

## Common Development Tasks

### Testing the Monitoring System
```javascript
// Run manual test
testRun()
```

### Resetting Processed Items
```javascript
// Clear all processed lists to re-trigger notifications
resetProcessedLists()
```

### Setting Up Initial Configuration
```javascript
// Configure script properties
setupScriptProperties()

// Set up hourly trigger
setupTimeTrigger()
```

## Spreadsheet Structure

The monitored spreadsheet expects these columns:
- **A**: 番号 (Number) - Unique identifier
- **B**: 起案日 (Issue Date) - Date when the question was raised
- **C**: 発信者 (Sender)
- **D**: フェーズ (Phase)
- **E**: カテゴリ (Category)
- **F**: 参照資料 (Reference Materials)
- **G**: 質疑内容 (Question Content)
- **H**: 宛先 (Recipient)
- **I**: 回答日 (Answer Date)
- **J**: 回答 (Answer Content)

## Script Properties Configuration

Required properties in Google Apps Script:
- `DISCORD_WEBHOOKS`: Comma-separated Discord webhook URLs
- `EMAIL_ADDRESSES`: Comma-separated email addresses for notifications
- `BACKUP_SHEET_ID`: Google Sheets ID for backup spreadsheet
- `PROCESSED_QUESTIONS`: JSON array of processed question IDs (initialize as `[]`)
- `PROCESSED_ANSWERS`: JSON array of processed answer IDs (initialize as `[]`)

## Key Functions

### Monitoring Functions
- `monitorQASheet()`: Main monitoring function (triggered hourly)
- `checkNewQuestions(data)`: Detects new questions (A-H columns filled)
- `checkNewAnswers(data)`: Detects new answers (I-J columns filled)
- `backupAndCheckChanges(sheet, data)`: Creates backups and detects changes/deletions

### Notification Functions
- `sendNewQuestionsNotification(questions)`: Sends Discord/email for new questions
- `sendNewAnswersNotification(answers)`: Sends Discord/email for new answers
- `formatQuestionsForDiscord(questions)`: Formats questions for Discord
- `formatAnswersForDiscord(answers)`: Formats answers for Discord

### Setup Functions
- `setupScriptProperties()`: Initialize script properties
- `setupTimeTrigger()`: Configure time-based trigger
- `testRun()`: Manual test execution
- `resetProcessedLists()`: Clear processed items

## Error Handling

The system includes comprehensive error handling:
- Try-catch blocks in main monitoring function
- Error notifications sent to Discord/email when failures occur
- Validation for required data fields before processing
- Safe parsing of JSON properties with fallbacks

## Development Notes

- This is a **container-bound script** - must be developed within the Google Sheets script editor
- Changes to column structure require updating the array indices in `checkNewQuestions()`, `checkNewAnswers()`, and `getColumnName()`
- Column B contains the issue date (起案日), Column C contains the sender (発信者)
- Discord messages have 2000 character limit - the script handles automatic message splitting
- All datetime handling uses JST (Japan Standard Time) timezone