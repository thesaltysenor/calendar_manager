# 📅 Google Calendar Manager

A Python-based tool to manage Google Calendar events efficiently. This script provides functionalities for adding, managing, and syncing events across your calendars, with robust options for bulk imports and customizable event settings.

## 🚀 Features

1. **List Calendars**  
   - Display all calendars in your Google Calendar account.

2. **Create Calendar**  
   - Add a new calendar with a custom name.

3. **Add Event**  
   - Add single or multiple events with structured date and time prompts.

4. **Bulk Import Events from CSV**  
   - Add events in bulk using a properly formatted CSV file.

5. **Add Multiple Events Continuously**  
   - Add multiple events interactively without restarting the script.

6. **Sync Event Colors**  
   - Ensure event colors match the custom color assigned to each calendar.

7. **Inspect Calendar Color**  
   - Retrieve and display the custom color set for a specific calendar.

## 🛠️ Requirements

- Python 3.8+
- Google Calendar API credentials (`credentials.json`)
- Required Python packages:
  ```bash
  pip install -r requirements.txt
  ```

## 📂 Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/calendar_manager.git
   cd calendar_manager
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Place your `credentials.json` file in the root directory.

5. Run the script:
   ```bash
   python calendar_manager.py
   ```

## 📊 CSV Format for Bulk Import

Your CSV should follow this format:

```csv
Summary,Start Date,Start Time,End Date,End Time
Event 1,2025-01-01,09:00:00,2025-01-01,10:00:00
Event 2,2025-01-02,13:00:00,2025-01-02,14:00:00
```

- `Summary`: Event title
- `Start Date` & `End Date`: Format `YYYY-MM-DD`
- `Start Time` & `End Time`: Format `HH:MM:SS`

## 📝 Usage

### Run the Script
```bash
python calendar_manager.py
```

### Choose an Option
1. **List Calendars** - Display all your calendars.
2. **Create Calendar** - Add a new calendar.
3. **Add Event** - Add single or multiple events.
4. **Import Bulk Events from CSV** - Import events in bulk.
5. **Add Multiple Events Continuously** - Keep adding events interactively.
6. **Sync Event Colors** - Ensure event colors align with calendar colors.
7. **Inspect Calendar Color** - Check calendar color settings.

## 🎨 Color Configuration

Each calendar is assigned a default custom color based on the `calendars` dictionary:
```python
calendars = {
    "Gopher Hockey": "#ac503c",
    "My Appointments": "#007ba7",
    "Bills": "#ccae00",
    ...
}
```
Colors are mapped using the `color_map` dictionary:
```python
color_map = {
    "#ac503c": "11",
    "#007ba7": "9",
    "#ccae00": "5",
    ...
}
```

## 🛡️ Authentication
- The script uses OAuth2 for Google Calendar API authentication.
- A `token.json` file will be generated after the first successful run.

## 🐞 Troubleshooting

- **Token Issues:** Delete `token.json` and re-authenticate.
- **CSV Import Errors:** Ensure your file follows the specified format.
- **Color Syncing Issues:** Verify custom calendar colors in Google Calendar.

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss your ideas.

## 📄 License

This project is licensed under the MIT License.

## 🌟 Acknowledgments

- Google Calendar API
- Python Community

---

Happy Scheduling! 🎉

