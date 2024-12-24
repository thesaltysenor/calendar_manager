from datetime import datetime, timedelta
import logging
import os
import csv
import json
from dotenv import load_dotenv
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# Load environment variables
load_dotenv()

# IMPLEMENT ERROR HANDLING
# IMPLEMENT EVENT IMPORT VALIDATION
# IMPLEMENT COMPLETE ENV MANAGEMENT
# Logging Configuration EXPAND ON THIS

LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "calendar_manager.log")

# Ensure the log directory exists
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),  # Log to file
        logging.StreamHandler()         # Log to console
    ]
)

# Google Calendar Credentials
CLIENT_ID = os.getenv("GOOGLE_CALENDAR_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_CALENDAR_CLIENT_SECRET")
TOKEN = os.getenv("GOOGLE_CALENDAR_TOKEN")
REFRESH_TOKEN = os.getenv("GOOGLE_CALENDAR_REFRESH_TOKEN")
SCOPES = json.loads(os.getenv("GOOGLE_CALENDAR_SCOPES", '["https://www.googleapis.com/auth/calendar"]'))
DEFAULT_TIMEZONE = os.getenv("DEFAULT_TIMEZONE", "America/Chicago")

# Calendar Configuration
calendars = json.loads(os.getenv("CALENDAR_NAMES", "{}"))
color_map = json.loads(os.getenv("COLOR_MAP", "{}"))
event_templates = json.loads(os.getenv("EVENT_TEMPLATES", "{}"))

def authenticate_google_calendar():
    """Authenticate and return a Google Calendar API service instance."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("calendar", "v3", credentials=creds)

service = authenticate_google_calendar()

def list_calendars():
    """List all calendars in the user's calendar list."""
    calendars_list = service.calendarList().list().execute()
    for calendar in calendars_list["items"]:
        print(f"Calendar Name: {calendar['summary']} (ID: {calendar['id']})")

def create_calendar(name):
    """Create a new calendar with the given name."""
    calendar = {
        "summary": name,
        "timeZone": "America/Chicago",
    }
    created_calendar = service.calendars().insert(body=calendar).execute()
    print(f"Calendar created: {created_calendar['summary']} (ID: {created_calendar['id']})")
    return created_calendar["id"]

def create_event(calendar_name, summary, start_time, end_time):
    """Add an event to a specific calendar with proper color matching."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    # Validate time range
    start = datetime.strptime(start_time, "%Y-%m-%dT%H:%M:%S")
    end = datetime.strptime(end_time, "%Y-%m-%dT%H:%M:%S")
    if start >= end:
        print("Error: Start time must be earlier than end time.")
        return

    # Get the calendar's assigned color
    color_hex = calendars.get(calendar_name, "#236192")  # Default if not in .env
    color_id = color_map.get(color_hex, "1")  # Default color if not mapped

    event = {
        "summary": summary,
        "start": {"dateTime": start_time, "timeZone": DEFAULT_TIMEZONE},
        "end": {"dateTime": end_time, "timeZone": DEFAULT_TIMEZONE},
        "colorId": color_id,
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "email", "minutes": 30},
                {"method": "popup", "minutes": 10},
            ],
        },
    }

    try:
        created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f"✅ Event created: {created_event['htmlLink']}")
    except Exception as e:
        print(f"❌ Error creating event: {e}")

def prompt_for_datetime(prompt_text):
    """Prompt the user for a date and time in a more structured way."""
    print(f"{prompt_text}:")
    year = input("  Year (e.g., 2024): ")
    month = input("  Month (1-12): ")
    day = input("  Day (1-31): ")
    hour = input("  Hour (0-23): ")
    minute = input("  Minute (0-59): ")
    return f"{year}-{month.zfill(2)}-{day.zfill(2)}T{hour.zfill(2)}:{minute.zfill(2)}:00"

def get_calendar_id(calendar_name):
    """Get the calendar ID for a given calendar name."""
    calendars_list = service.calendarList().list().execute()
    for calendar in calendars_list["items"]:
        if calendar["summary"] == calendar_name:
            return calendar["id"]
    return None

def add_event_with_multiple_dates(calendar_name):
    """Add a single event with multiple dates and times."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    summary = input("Enter event summary (e.g., Twins vs Yankees): ").strip()
    events = []

    while True:
        print("\nEnter details for an event occurrence:")
        start_time = prompt_for_datetime("Enter start date and time")
        end_time = prompt_for_datetime("Enter end date and time")

        try:
            start_dt = datetime.fromisoformat(start_time)
            end_dt = datetime.fromisoformat(end_time)
            if start_dt >= end_dt:
                print("❌ Error: Start time must be earlier than end time.")
                continue
        except ValueError as e:
            print(f"❌ Invalid datetime format: {e}")
            continue

        events.append({"summary": summary, "start_time": start_time, "end_time": end_time})

        more = input("Add another date/time for this event? (y/n): ").strip().lower()
        if more != "y":
            break

    for event in events:
        try:
            create_event(calendar_name, event['summary'], event['start_time'], event['end_time'])
        except Exception as e:
            print(f"❌ Error creating event '{event['summary']}': {e}")

    print(f"✅ {len(events)} occurrence(s) of '{summary}' added successfully to calendar '{calendar_name}'!")

def add_multiple_unique_events(calendar_name):
    """Continuously add unique events to a calendar until the user stops."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    events = []

    while True:
        summary = input("\nEnter event summary (or type 'done' to finish): ").strip()
        if summary.lower() == "done":
            break

        start_time = prompt_for_datetime("Enter start date and time")
        end_time = prompt_for_datetime("Enter end date and time")

        try:
            start_dt = datetime.fromisoformat(start_time)
            end_dt = datetime.fromisoformat(end_time)
            if start_dt >= end_dt:
                print("❌ Error: Start time must be earlier than end time.")
                continue
        except ValueError as e:
            print(f"❌ Invalid datetime format: {e}")
            continue

        events.append({"summary": summary, "start_time": start_time, "end_time": end_time})

    for event in events:
        try:
            create_event(calendar_name, event['summary'], event['start_time'], event['end_time'])
        except Exception as e:
            print(f"❌ Error creating event '{event['summary']}': {e}")

    print(f"✅ {len(events)} unique event(s) added successfully to calendar '{calendar_name}'!")

def import_from_csv(calendar_name, csv_file):
    """Import events from a CSV file."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    # Fetch calendar's default color ID
    color_id = get_calendar_color_id(calendar_name)

    try:
        with open(csv_file, mode="r", newline="", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            print("✅ CSV Headers:", reader.fieldnames)  # Debug to check headers
            
            for row in reader:
                summary = row.get("Summary")
                start_date = row.get("Start Date")
                start_time = row.get("Start Time")
                end_date = row.get("End Date")
                end_time = row.get("End Time")
                
                if not summary or not start_date or not start_time or not end_date or not end_time:
                    print("❌ Skipping invalid row (missing fields):", row)
                    continue

                try:
                    start_datetime = f"{start_date}T{start_time}"
                    end_datetime = f"{end_date}T{end_time}"
                    
                    start_dt = datetime.fromisoformat(start_datetime)
                    end_dt = datetime.fromisoformat(end_datetime)
                    
                    if start_dt >= end_dt:
                        print("❌ Skipping row with invalid time range:", row)
                        continue
                    
                    # Pass the colorId dynamically
                    event = {
                        "summary": summary,
                        "start": {"dateTime": start_datetime, "timeZone": os.getenv("DEFAULT_TIMEZONE", "America/Chicago")},
                        "end": {"dateTime": end_datetime, "timeZone": os.getenv("DEFAULT_TIMEZONE", "America/Chicago")},
                        "colorId": color_id,
                        "reminders": {
                            "useDefault": False,
                            "overrides": [
                                {"method": "email", "minutes": 30},
                                {"method": "popup", "minutes": 10},
                            ],
                        },
                    }

                    created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
                    print(f"✅ Event added: {summary} on {start_date}")
                except ValueError as e:
                    print(f"❌ Skipping row with invalid datetime format: {row}. Error: {e}")
                except KeyError as e:
                    print(f"❌ Missing key in row: {row}. Error: {e}")
    except FileNotFoundError:
        print("❌ CSV file not found. Please provide a valid path.")
    except Exception as e:
        print(f"❌ Error processing CSV: {e}")

def add_event_using_template(calendar_name):
    """Add an event using a predefined template."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    print("Available Event Templates:")
    for idx, template_name in enumerate(event_templates.keys(), start=1):
        print(f"{idx}. {template_name}")

    choice = input("Select a template by number: ").strip()
    try:
        template_index = int(choice) - 1
        template_keys = list(event_templates.keys())
        selected_template = template_keys[template_index]
    except (ValueError, IndexError):
        print("❌ Invalid choice! Please select a valid template.")
        return

    template = event_templates[selected_template]
    summary = template.get("summary", "Untitled Event")
    duration = template.get("duration", 60)  # Default duration if not provided

    # Prompt for date and start time
    start_time = prompt_for_datetime("Enter event start date and time")
    start_dt = datetime.fromisoformat(start_time)
    end_dt = start_dt + timedelta(minutes=duration)

    start_time_str = start_dt.isoformat()
    end_time_str = end_dt.isoformat()

    try:
        create_event(calendar_name, summary, start_time_str, end_time_str)
        print(f"✅ Event '{summary}' added to calendar '{calendar_name}' using template '{selected_template}'.")
    except Exception as e:
        print(f"❌ Failed to create event: {e}")

def add_recurring_event(calendar_name):
    """Add a recurring event to a specific calendar."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    summary = input("Enter event summary (e.g., Weekly Standup Meeting): ").strip()
    start_time = prompt_for_datetime("Enter start date and time")
    end_time = prompt_for_datetime("Enter end date and time")
    
    try:
        start_dt = datetime.fromisoformat(start_time)
        end_dt = datetime.fromisoformat(end_time)
        if start_dt >= end_dt:
            print("❌ Error: Start time must be earlier than end time.")
            return
    except ValueError as e:
        print(f"❌ Invalid datetime format: {e}")
        return

    # Recurrence Pattern
    print("\nChoose a recurrence pattern:")
    print("1️⃣ Daily")
    print("2️⃣ Weekly")
    print("3️⃣ Monthly")
    print("4️⃣ Yearly")
    recurrence_choice = input("👉 Enter your choice: ").strip()

    if recurrence_choice == "1":
        frequency = "DAILY"
    elif recurrence_choice == "2":
        frequency = "WEEKLY"
    elif recurrence_choice == "3":
        frequency = "MONTHLY"
    elif recurrence_choice == "4":
        frequency = "YEARLY"
    else:
        print("❌ Invalid recurrence choice.")
        return

    end_condition = input("Should the recurrence end by date (d) or after a number of occurrences (n)? ").strip().lower()
    if end_condition == "d":
        end_date = input("Enter end date (YYYY-MM-DD): ").strip()
        recurrence_rule = f"RRULE:FREQ={frequency};UNTIL={end_date.replace('-', '')}T000000Z"
    elif end_condition == "n":
        count = input("Enter the number of occurrences: ").strip()
        if not count.isdigit():
            print("❌ Invalid number of occurrences.")
            return
        recurrence_rule = f"RRULE:FREQ={frequency};COUNT={count}"
    else:
        print("❌ Invalid end condition choice.")
        return

    # Prepare the event with recurrence
    event = {
        "summary": summary,
        "start": {"dateTime": start_time, "timeZone": DEFAULT_TIMEZONE},
        "end": {"dateTime": end_time, "timeZone": DEFAULT_TIMEZONE},
        "recurrence": [recurrence_rule],
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "email", "minutes": 30},
                {"method": "popup", "minutes": 10},
            ],
        },
    }

    try:
        created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f"✅ Recurring event created: {created_event['htmlLink']}")
    except Exception as e:
        print(f"❌ Error creating recurring event: {e}")

def search_events(calendar_name):
    """Search events in a calendar by keyword or date range."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    print("\n🔍 Search Options:")
    print("1️⃣ Search by keyword")
    print("2️⃣ Search by date range")
    choice = input("👉 Enter your choice: ").strip()

    search_query = None
    time_min = None
    time_max = None

    if choice == "1":
        search_query = input("Enter keyword to search in event summaries: ").strip()
    elif choice == "2":
        time_min = input("Enter start date (YYYY-MM-DD): ").strip()
        time_max = input("Enter end date (YYYY-MM-DD): ").strip()
    else:
        print("❌ Invalid choice.")
        return

    try:
        events_result = service.events().list(
            calendarId=calendar_id,
            q=search_query,
            timeMin=f"{time_min}T00:00:00Z" if time_min else None,
            timeMax=f"{time_max}T23:59:59Z" if time_max else None,
            singleEvents=True,
            orderBy="startTime"
        ).execute()

        events = events_result.get('items', [])
        if not events:
            print("❌ No matching events found.")
            return

        print("\n🔗 Matching Events:")
        for i, event in enumerate(events, start=1):
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(f"{i}. {event['summary']} | Start: {start} | ID: {event['id']}")

        return events

    except Exception as e:
        print(f"❌ Error searching events: {e}")

def update_event(calendar_name):
    """Update an existing event in a calendar."""
    events = search_events(calendar_name)
    if not events:
        return

    event_index = input("\nEnter the number of the event you want to update: ").strip()
    if not event_index.isdigit() or int(event_index) < 1 or int(event_index) > len(events):
        print("❌ Invalid event selection.")
        return

    selected_event = events[int(event_index) - 1]
    event_id = selected_event['id']

    print("\n🛠️ Update Event Details:")
    summary = input(f"New summary (leave blank to keep '{selected_event['summary']}'): ").strip()
    start_time = prompt_for_datetime("Enter new start date and time (leave blank to keep current)")
    end_time = prompt_for_datetime("Enter new end date and time (leave blank to keep current)")
    description = input("Enter new description (leave blank to keep current): ").strip()

    # Prepare updated event body
    updated_event = {
        "summary": summary or selected_event['summary'],
        "start": selected_event['start'],
        "end": selected_event['end'],
        "description": description or selected_event.get('description', "")
    }

    if start_time:
        updated_event['start'] = {"dateTime": start_time, "timeZone": DEFAULT_TIMEZONE}
    if end_time:
        updated_event['end'] = {"dateTime": end_time, "timeZone": DEFAULT_TIMEZONE}

    # Validate time range
    try:
        if start_time and end_time:
            start_dt = datetime.fromisoformat(start_time)
            end_dt = datetime.fromisoformat(end_time)
            if start_dt >= end_dt:
                print("❌ Error: Start time must be earlier than end time.")
                return
    except ValueError as e:
        print(f"❌ Invalid datetime format: {e}")
        return

    try:
        service.events().update(
            calendarId=get_calendar_id(calendar_name),
            eventId=event_id,
            body=updated_event
        ).execute()
        print("✅ Event updated successfully.")
    except Exception as e:
        print(f"❌ Error updating event: {e}")


def get_calendar_color_id(calendar_name):
    """Retrieve the default color ID of a calendar."""
    try:
        calendars_list = service.calendarList().list().execute()
        for calendar in calendars_list["items"]:
            if calendar["summary"] == calendar_name:
                # Retrieve the calendar's default color ID
                color_id = calendar.get("colorId", "1")  # Default to '1' if not found
                return color_id
    except Exception as e:
        print(f"Failed to retrieve calendar color for '{calendar_name}': {e}")
    return "1"  # Default color ID if retrieval fails

def sync_event_colors(calendar_name):
    """Sync event colors to match the calendar's assigned color."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    # Get the calendar's designated color
    color_hex = calendars.get(calendar_name, "#236192")
    color_id = color_map.get(color_hex, "1")

    try:
        events_result = service.events().list(calendarId=calendar_id).execute()
        events = events_result.get('items', [])

        updated_count = 0
        for event in events:
            event_id = event.get('id')
            event_summary = event.get('summary', 'No Summary')
            event_color_id = event.get('colorId', None)

            if event_color_id != color_id:
                event['colorId'] = color_id
                service.events().update(calendarId=calendar_id, eventId=event_id, body=event).execute()
                print(f"✅ Updated color for event: {event_summary}")
                updated_count += 1

        print(f"🎨 Finished syncing colors. {updated_count} events updated.")
    except Exception as e:
        print(f"❌ Error syncing event colors: {e}")

def inspect_calendar_color(calendar_name):
    """Inspect the colorId of a calendar."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        print(f"Calendar '{calendar_name}' not found!")
        return

    try:
        calendar = service.calendarList().get(calendarId=calendar_id).execute()
        color_id = calendar.get('colorId', 'Default')
        print(f"Calendar '{calendar_name}' has colorId: {color_id}")
    except Exception as e:
        print(f"Failed to retrieve calendar color: {e}")

def main():
    """Main menu for the script."""
    print("\n🗓️  Google Calendar Manager")
    print("1️⃣  List Calendars")
    print("2️⃣  Create Calendar")
    print("3️⃣  Add Event with Multiple Dates/Times")
    print("4️⃣  Add Multiple Unique Events")
    print("5️⃣  Import Bulk Events from CSV")
    print("6️⃣  Sync Event Colors with Calendar Color")
    print("7️⃣  Inspect Calendar Color")
    print("8️⃣  Add Event from Template")
    print("9️⃣ Add Recurring Event")
    print("🔟 Search Events")
    print("1️⃣1️⃣ Update Event")
    print("🛑  Type 'exit' to quit.")
    
    choice = input("👉 Enter your choice: ").strip().lower()

    if choice == "1":
        list_calendars()
    elif choice == "2":
        calendar_name = input("Enter calendar name: ").strip()
        create_calendar(calendar_name)
    elif choice == "3":
        calendar_name = input("Enter calendar name: ").strip()
        add_event_with_multiple_dates(calendar_name)
    elif choice == "4":
        calendar_name = input("Enter calendar name: ").strip()
        add_multiple_unique_events(calendar_name)
    elif choice == "5":
        calendar_name = input("Enter calendar name: ").strip()
        csv_file = input("Enter the path to the CSV file: ").strip()
        import_from_csv(calendar_name, csv_file)
    elif choice == "6":
        calendar_name = input("Enter calendar name: ").strip()
        sync_event_colors(calendar_name)
    elif choice == "7":
        calendar_name = input("Enter calendar name: ").strip()
        inspect_calendar_color(calendar_name)
    elif choice == "8":
        calendar_name = input("Enter calendar name: ").strip()
        add_event_using_template(calendar_name)
    elif choice == "9":
        calendar_name = input("Enter calendar name: ").strip()
        add_recurring_event(calendar_name)
    elif choice == "10":
        calendar_name = input("Enter calendar name: ").strip()
        search_events(calendar_name)
    elif choice == "11":
        calendar_name = input("Enter calendar name: ").strip()
        update_event(calendar_name)
    elif choice == "exit":
        print("👋 Goodbye!")
        exit()
    else:
        print("❌ Invalid choice! Please try again.")

if __name__ == "__main__":
    main()
