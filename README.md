# Telegram Scheduled Messages Bot

A Google Apps Script that automatically sends scheduled Telegram messages based on a Google Spreadsheet configuration. Supports various repeat patterns including one-time, daily, weekly, and custom intervals.

## Features

- 📅 **Flexible Scheduling**: One-time, daily, weekly, monthly, yearly, and custom interval messages
- 🔄 **Smart Repeat Patterns**: Support for specific weekdays, every X days, and complex schedules
- 🛡️ **Duplicate Prevention**: Prevents sending the same message multiple times within a minute
- 🔐 **Secure Configuration**: Uses Google Apps Script Properties Service for secure credential storage
- 🧪 **Debug Mode**: Comprehensive testing function to analyze scheduling logic
- 📊 **Error Handling**: Robust error handling with detailed logging
- ⏰ **Forgiving Timing**: Works even if triggers are delayed by a few seconds

## Setup

### 1. Google Spreadsheet Setup

**Quick Start:** Use this pre-configured template: [Telegram Scheduled Messages Template](https://docs.google.com/spreadsheets/d/1KC2QzXQDOU7whuxbIYAxDnHK0Lx0ORNLz9cJPcA6X7k/edit?usp=sharing)

*Or create your own Google Spreadsheet with the following columns:*

| Column | Name | Description | Example |
|--------|------|-------------|---------|
| A | DateTime | When to send the message | `2024-08-21 19:00:00` |
| B | Message | The message content | `🎉 Daily reminder!` |
| C | Repeat | How often to repeat | `daily`, `weekly`, `monday`, `every 3 days` |
| D | Enabled | Whether message is active | `✅`, `true`, `yes` |
| E | Last Sent | Auto-updated timestamp | *Auto-filled* |

### 2. Telegram Bot Setup

1. Create a new bot with [@BotFather](https://t.me/BotFather)
2. Get your bot token (format: `123456789:ABCdefGHIjklMNOpqrsTUVwxyz`)
3. Add your bot to the target chat/channel
4. Get the chat ID (use [@userinfobot](https://t.me/userinfobot) or check bot logs)

### 3. Google Apps Script Setup

1. Open [Google Apps Script](https://script.google.com/)
2. Create a new project
3. Replace the default code with `montagtontag_appscript.js`
4. Go to **Project Settings** → **Script Properties**
5. Add the following properties:
   - `TELEGRAM_TOKEN`: Your bot token
   - `CHAT_ID`: Your target chat ID
   - `DEBUG_CHAT_ID`: Your debug chat ID (optional)

### 4. Trigger Setup

1. In Apps Script, go to **Triggers** (clock icon)
2. Add a new trigger:
   - Function: `sendScheduledMessages`
   - Event source: Time-driven
   - Type: Minutes timer
   - Interval: Every minute (or as needed)

## Usage

### Repeat Patterns

The script supports various repeat patterns in the "Repeat" column:

- **One-time**: Leave empty, `don't`, or `don't repeat`
- **Regular intervals**: `daily`, `weekly`, `monthly`, `yearly`
- **Specific weekdays**: `monday`, `tuesday`, `wednesday`, etc.
- **Custom intervals**: `every 2 days`, `every 7 days`, `every 14 days`
- **Testing**: `every minute` (for debugging)

### Enabled Status

The "Enabled" column accepts multiple formats:
- Boolean: `true`/`false`
- Text: `yes`/`no`
- Emoji: `✅`/`❌`
- Mixed: Any string containing `✅` or `true`/`yes` (case-insensitive)

### Message Formatting

Messages support Telegram's Markdown formatting:
- `**bold**` for **bold text**
- `*italic*` for *italic text*
- `[link](URL)` for clickable links
- Line breaks work normally

## Debugging

Use the `debugTest()` function to analyze your scheduling logic:

1. In Apps Script, select `debugTest` from the function dropdown
2. Click "Run"
3. Check your debug chat for detailed analysis of each message
4. Review console logs for technical details

The debug output shows:
- Current scheduling status for each message
- Why messages will or won't send
- Time/date matching analysis
- Duplicate prevention status

## Security Notes

- ✅ **Credentials are stored securely** in Google Apps Script Properties Service
- ✅ **Fallback values** are included for development but should be removed in production
- ✅ **Error messages** don't expose sensitive information
- ⚠️ **Remove hardcoded tokens** from the script after setting up Properties Service

## Troubleshooting

### Messages Not Sending

1. **Check the debug function**: Run `debugTest()` to see detailed analysis
2. **Verify credentials**: Ensure `TELEGRAM_TOKEN` and `CHAT_ID` are set correctly
3. **Check timing**: Messages need exact hour:minute match (seconds are ignored)
4. **Verify enabled status**: Ensure the "Enabled" column contains `✅`, `true`, or `yes`

### Common Issues

- **"400 Bad Request"**: Usually Markdown formatting errors in messages
- **"Unauthorized"**: Invalid bot token or bot not added to chat
- **"Chat not found"**: Incorrect chat ID
- **Time mismatch**: Check your spreadsheet timezone vs script timezone

## File Structure

```
Kamine_TelegramBots/
├── montagtontag_appscript.js    # Main script file
├── README.md                    # This documentation
└── .gitignore                   # Git ignore rules
```

## Contributing

1. Fork the repository
2. Make your changes
3. Test thoroughly using `debugTest()`
4. Submit a pull request

## License

This project is open source. Feel free to use and modify as needed.

---

**Made with ❤️ for automated Telegram messaging**
