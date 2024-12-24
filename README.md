📅 Google Calendar Manager
🚀 Overview
Google Calendar Manager is a Python-based tool that streamlines calendar and event management via the Google Calendar API. It provides an intuitive command-line interface to create, update, search, and manage calendar events with advanced functionalities such as recurring events, templates, bulk imports, and color synchronization.

🛠️ Features
✅ List Calendars – View all your Google Calendars.
✅ Create Calendar – Add new calendars with custom configurations.
✅ Add Events – Create single, recurring, or multiple events across dates and times.
✅ Import Events – Bulk import events from a CSV file.
✅ Search Events – Search events by keyword or date range.
✅ Update Events – Modify event details like time, summary, or description.
✅ Recurring Events – Easily set up daily, weekly, monthly, or yearly events.
✅ Event Templates – Use predefined templates for quick event creation.
✅ Sync Event Colors – Ensure event colors match their assigned calendar color.
✅ Inspect Calendar Colors – View calendar-specific color configurations.
✅ Comprehensive Logging – Monitor errors, warnings, and API activity via log files.
📦 Setup and Installation
1. Clone the Repository
bash
Copy code
git clone https://github.com/yourusername/google-calendar-manager.git
cd google-calendar-manager
2. Create a Virtual Environment
bash
Copy code
python -m venv venv
source venv/bin/activate  # MacOS/Linux
venv\Scripts\activate     # Windows
3. Install Dependencies
bash
Copy code
pip install -r requirements.txt
4. Configure Environment Variables
Create a .env file in the project root directory with the following structure:

env
Copy code
# Google Calendar API Credentials
GOOGLE_CALENDAR_CLIENT_ID="<your_client_id>"
GOOGLE_CALENDAR_CLIENT_SECRET="<your_client_secret>"
GOOGLE_CALENDAR_TOKEN="<your_token>"
GOOGLE_CALENDAR_REFRESH_TOKEN="<your_refresh_token>"
GOOGLE_CALENDAR_SCOPES='["https://www.googleapis.com/auth/calendar"]'
DEFAULT_TIMEZONE="America/Chicago"

# Calendar Configurations (JSON)
CALENDAR_NAMES={
    "Example Calendar 1": "#color1",
    "Example Calendar 2": "#color2"
}

COLOR_MAP={
    "#color1": "1",
    "#color2": "2"
}

EVENT_TEMPLATES={
    "Meeting": {"summary": "Meeting", "duration": 30},
    "Workout": {"summary": "Workout Session", "duration": 60}
}
Replace the placeholder values (<your_*) with the appropriate API credentials.

🏃‍♀️ Running the Application
Start the script using:

bash
Copy code
python calendar_manager.py
Main Menu Options
1️⃣ List Calendars
2️⃣ Create Calendar
3️⃣ Add Event with Multiple Dates/Times
4️⃣ Add Multiple Unique Events
5️⃣ Import Bulk Events from CSV
6️⃣ Sync Event Colors with Calendar Color
7️⃣ Inspect Calendar Color
8️⃣ Add Event from Template
9️⃣ Add Recurring Event
🔟 Search Events
1️⃣1️⃣ Update Event
🛑 Exit

Follow on-screen instructions for each menu option.

📊 CSV Import Template
To bulk import events, create a CSV file with the following structure:

csv
Copy code
Summary,Start Date,Start Time,End Date,End Time
"Project Kickoff",2024-12-01,09:00:00,2024-12-01,10:00:00
"Client Meeting",2024-12-05,13:00:00,2024-12-05,14:00:00
Save it as events.csv and select the Import Bulk Events from CSV option in the menu.

📖 Logging
Logs are saved in the logs/calendar_manager.log file.
Detailed activity, warnings, and errors are recorded for debugging and monitoring purposes.
🧠 Best Practices
Keep your .env file secure and never share it publicly.
Regularly review logs for API errors or misconfigurations.
Ensure your CSV files follow the correct structure when importing events.
🤝 Contributing
Contributions are welcome! Follow these steps:

Fork the repository.
Create a new branch: git checkout -b feature-branch.
Make changes and commit: git commit -m "Add feature XYZ".
Push changes: git push origin feature-branch.
Submit a Pull Request.
🛡️ License
This project is licensed under the MIT License.

💬 Support
Open an issue on the repository.
Contact via email: your-email@example.com.
🎯 Happy Calendar Managing! 🗓️✨