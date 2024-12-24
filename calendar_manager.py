from datetime import datetime, timedelta
import logging
import os
import sys
import csv
import json
from dotenv import load_dotenv
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# ✅ Load environment variables
load_dotenv()

# ✅ Logging Configuration with Enhanced Error Handling
LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "calendar_manager.log")

# Ensure the log directory exists
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),  # Explicit UTF-8 encoding
        logging.StreamHandler(stream=sys.stdout)         # Ensure console encoding
    ]
)

logging.debug("Logging system initialized successfully.")

# ✅ Google Calendar Credentials
try:
    CLIENT_ID = os.getenv("GOOGLE_CALENDAR_CLIENT_ID")
    CLIENT_SECRET = os.getenv("GOOGLE_CALENDAR_CLIENT_SECRET")
    TOKEN = os.getenv("GOOGLE_CALENDAR_TOKEN")
    REFRESH_TOKEN = os.getenv("GOOGLE_CALENDAR_REFRESH_TOKEN")
    TOKEN_URI = os.getenv("GOOGLE_CALENDAR_TOKEN_URI", "https://oauth2.googleapis.com/token")
    AUTH_URI = os.getenv("GOOGLE_CALENDAR_AUTH_URI", "https://accounts.google.com/o/oauth2/auth")
    SCOPES = json.loads(os.getenv("GOOGLE_CALENDAR_SCOPES", '["https://www.googleapis.com/auth/calendar"]'))
    DEFAULT_TIMEZONE = os.getenv("DEFAULT_TIMEZONE", "America/Chicago")
    calendars = json.loads(os.getenv("CALENDAR_NAMES", "{}"))
    color_map = json.loads(os.getenv("COLOR_MAP", "{}"))
    event_templates = json.loads(os.getenv("EVENT_TEMPLATES", "{}"))
    
    # Validate Critical Environment Variables
    required_env_vars = [
        "GOOGLE_CALENDAR_CLIENT_ID",
        "GOOGLE_CALENDAR_CLIENT_SECRET",
        "GOOGLE_CALENDAR_TOKEN",
        "GOOGLE_CALENDAR_REFRESH_TOKEN",
        "DEFAULT_TIMEZONE"
    ]
    for var in required_env_vars:
        if not os.getenv(var):
            raise EnvironmentError(f"❌ Missing required environment variable: {var}")
    
    logging.debug("✅ Environment variables loaded and validated successfully.")

except (json.JSONDecodeError, EnvironmentError) as e:
    logging.error(f"❌ Error with environment setup: {e}")
    raise SystemExit("❌ Critical environment variables are missing or misconfigured. Exiting...")

# ✅ Authentication for Google Calendar API
def authenticate_google_calendar():
    """Authenticate and return a Google Calendar API service instance."""
    creds = None
    try:
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        logging.debug("✅ Google Calendar authentication successful.")
        return build("calendar", "v3", credentials=creds)
    except Exception as e:
        logging.error(f"❌ Google Calendar authentication failed: {e}")
        raise SystemExit("❌ Failed to authenticate with Google Calendar API.")

service = authenticate_google_calendar()


# ✅ Validate Environment Variables
def validate_env_variables():
    """Validate all required environment variables are set."""
    required_vars = [
        "GOOGLE_CALENDAR_CLIENT_ID",
        "GOOGLE_CALENDAR_CLIENT_SECRET",
        "GOOGLE_CALENDAR_TOKEN",
        "GOOGLE_CALENDAR_REFRESH_TOKEN",
        "DEFAULT_TIMEZONE"
    ]

    missing_vars = [var for var in required_vars if not os.getenv(var)]
    if missing_vars:
        missing_vars_str = ", ".join(missing_vars)
        logging.error(f"❌ Missing required environment variables: {missing_vars_str}")
        raise EnvironmentError(f"❌ Missing required environment variables: {missing_vars_str}")
    
    logging.debug("✅ Environment variables validated successfully.")


# ✅ List Calendars
def list_calendars():
    """List all calendars in the user's calendar list."""
    try:
        calendars_list = service.calendarList().list().execute()
        if not calendars_list.get("items"):
            print("❌ No calendars found.")
            logging.warning("No calendars found in the user's calendar list.")
            return

        print("\n📅 Available Calendars:")
        for calendar in calendars_list["items"]:
            print(f"- {calendar['summary']} (ID: {calendar['id']})")
        logging.debug("✅ Calendars listed successfully.")
    except Exception as e:
        logging.error(f"❌ Failed to list calendars: {e}")
        print(f"❌ Error listing calendars: {e}")


# ✅ Create Calendar
def create_calendar(name):
    """Create a new calendar with the given name."""
    calendar = {
        "summary": name,
        "timeZone": DEFAULT_TIMEZONE,
    }

    try:
        created_calendar = service.calendars().insert(body=calendar).execute()
        print(f"✅ Calendar created: {created_calendar['summary']} (ID: {created_calendar['id']})")
        logging.debug(f"✅ Calendar '{name}' created successfully with ID: {created_calendar['id']}.")
        return created_calendar["id"]
    except Exception as e:
        logging.error(f"❌ Error creating calendar '{name}': {e}")
        print(f"❌ Error creating calendar '{name}': {e}")
        return None


# ✅ Create Event
def create_event(calendar_name, summary, start_time, end_time):
    """Add an event to a specific calendar with proper color matching."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.warning(f"❌ Calendar '{calendar_name}' not found!")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    try:
        # Validate time range
        start = datetime.fromisoformat(start_time)
        end = datetime.fromisoformat(end_time)
        if start >= end:
            logging.warning("❌ Error: Start time must be earlier than end time.")
            print("❌ Error: Start time must be earlier than end time.")
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

        created_event = service.events().insert(calendarId=calendar_id, body=event).execute()
        print(f"✅ Event created: {created_event['htmlLink']}")
        logging.debug(f"✅ Event '{summary}' created successfully in calendar '{calendar_name}'.")
    except ValueError as ve:
        logging.error(f"❌ Invalid datetime format: {ve}")
        print(f"❌ Invalid datetime format: {ve}")
    except Exception as e:
        logging.error(f"❌ Error creating event: {e}")
        print(f"❌ Error creating event: {e}")


# ✅ Prompt for Datetime
def prompt_for_datetime(prompt_text):
    """Prompt the user for a date and time in a more structured way."""
    print(f"{prompt_text}:")
    try:
        year = int(input("  Year (e.g., 2024): ").strip())
        month = int(input("  Month (1-12): ").strip())
        day = int(input("  Day (1-31): ").strip())
        hour = int(input("  Hour (0-23): ").strip())
        minute = int(input("  Minute (0-59): ").strip())

        # Validate datetime inputs
        datetime_obj = datetime(year, month, day, hour, minute)
        return datetime_obj.isoformat()
    except ValueError as e:
        logging.error(f"❌ Invalid datetime input: {e}")
        print(f"❌ Invalid datetime input: {e}")
        return prompt_for_datetime(prompt_text)


# ✅ Get Calendar ID
def get_calendar_id(calendar_name):
    """Get the calendar ID for a given calendar name."""
    try:
        calendars_list = service.calendarList().list().execute()
        for calendar in calendars_list.get("items", []):
            if calendar["summary"] == calendar_name:
                logging.debug(f"✅ Found calendar '{calendar_name}' with ID: {calendar['id']}")
                return calendar["id"]
        logging.warning(f"❌ Calendar '{calendar_name}' not found.")
        print(f"❌ Calendar '{calendar_name}' not found.")
        return None
    except Exception as e:
        logging.error(f"❌ Error fetching calendar ID: {e}")
        print(f"❌ Error fetching calendar ID: {e}")
        return None


# ✅ Add Event with Multiple Dates
def add_event_with_multiple_dates(calendar_name):
    """Add a single event with multiple dates and times."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
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
                logging.warning("❌ Start time must be earlier than end time.")
                continue
        except ValueError as e:
            print(f"❌ Invalid datetime format: {e}")
            logging.error(f"❌ Invalid datetime format: {e}")
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
            logging.error(f"❌ Error creating event '{event['summary']}': {e}")

    print(f"✅ {len(events)} occurrence(s) of '{summary}' added successfully to calendar '{calendar_name}'!")
    logging.debug(f"✅ {len(events)} occurrences of '{summary}' added to calendar '{calendar_name}'.")


# ✅ Add Multiple Unique Events
def add_multiple_unique_events(calendar_name):
    """Continuously add unique events to a calendar until the user stops."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
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
                logging.warning("❌ Start time must be earlier than end time.")
                continue
        except ValueError as e:
            print(f"❌ Invalid datetime format: {e}")
            logging.error(f"❌ Invalid datetime format: {e}")
            continue

        events.append({"summary": summary, "start_time": start_time, "end_time": end_time})

    for event in events:
        try:
            create_event(calendar_name, event['summary'], event['start_time'], event['end_time'])
        except Exception as e:
            print(f"❌ Error creating event '{event['summary']}': {e}")
            logging.error(f"❌ Error creating event '{event['summary']}': {e}")

    print(f"✅ {len(events)} unique event(s) added successfully to calendar '{calendar_name}'!")
    logging.debug(f"✅ {len(events)} unique events added to calendar '{calendar_name}'.")


# ✅ Import Events from CSV
def import_from_csv(calendar_name, csv_file):
    """Import events from a CSV file."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found.")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    # Fetch calendar's default color ID
    color_id = get_calendar_color_id(calendar_name)

    try:
        with open(csv_file, mode="r", newline="", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            if not reader.fieldnames:
                logging.error("❌ CSV file is empty or headers are missing.")
                print("❌ CSV file is empty or headers are missing.")
                return

            logging.info(f"✅ CSV Headers: {reader.fieldnames}")
            print("✅ CSV Headers:", reader.fieldnames)  # Debug to check headers

            for row in reader:
                summary = row.get("Summary", "").strip()
                start_date = row.get("Start Date", "").strip()
                start_time = row.get("Start Time", "").strip()
                end_date = row.get("End Date", "").strip()
                end_time = row.get("End Time", "").strip()

                # Validate required fields
                if not all([summary, start_date, start_time, end_date, end_time]):
                    logging.warning(f"❌ Skipping invalid row (missing fields): {row}")
                    print(f"❌ Skipping invalid row (missing fields): {row}")
                    continue

                try:
                    start_datetime = f"{start_date}T{start_time}"
                    end_datetime = f"{end_date}T{end_time}"

                    start_dt = datetime.fromisoformat(start_datetime)
                    end_dt = datetime.fromisoformat(end_datetime)

                    if start_dt >= end_dt:
                        logging.warning(f"❌ Skipping row with invalid time range: {row}")
                        print(f"❌ Skipping row with invalid time range: {row}")
                        continue

                    # Create event payload
                    event = {
                        "summary": summary,
                        "start": {"dateTime": start_datetime, "timeZone": DEFAULT_TIMEZONE},
                        "end": {"dateTime": end_datetime, "timeZone": DEFAULT_TIMEZONE},
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
                    logging.info(f"✅ Event added: {summary} on {start_date}")
                    print(f"✅ Event added: {summary} on {start_date}")
                except ValueError as e:
                    logging.error(f"❌ Skipping row with invalid datetime format: {row}. Error: {e}")
                    print(f"❌ Skipping row with invalid datetime format: {row}. Error: {e}")
                except KeyError as e:
                    logging.error(f"❌ Missing key in row: {row}. Error: {e}")
                    print(f"❌ Missing key in row: {row}. Error: {e}")
    except FileNotFoundError:
        logging.error("❌ CSV file not found. Please provide a valid path.")
        print("❌ CSV file not found. Please provide a valid path.")
    except Exception as e:
        logging.error(f"❌ Error processing CSV: {e}")
        print(f"❌ Error processing CSV: {e}")


# ✅ Add Event Using a Template
def add_event_using_template(calendar_name):
    """Add an event using a predefined template."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found.")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    print("\n📝 Available Event Templates:")
    for idx, template_name in enumerate(event_templates.keys(), start=1):
        print(f"{idx}. {template_name}")

    choice = input("Select a template by number: ").strip()
    try:
        template_index = int(choice) - 1
        template_keys = list(event_templates.keys())
        selected_template = template_keys[template_index]
    except (ValueError, IndexError):
        logging.warning("❌ Invalid choice! Please select a valid template.")
        print("❌ Invalid choice! Please select a valid template.")
        return

    template = event_templates[selected_template]
    summary = template.get("summary", "Untitled Event")
    duration = template.get("duration", 60)  # Default duration if not provided

    # Prompt for date and start time
    start_time = prompt_for_datetime("Enter event start date and time")
    try:
        start_dt = datetime.fromisoformat(start_time)
        end_dt = start_dt + timedelta(minutes=duration)

        start_time_str = start_dt.isoformat()
        end_time_str = end_dt.isoformat()
    except ValueError as e:
        logging.error(f"❌ Invalid datetime input: {e}")
        print(f"❌ Invalid datetime input: {e}")
        return

    try:
        create_event(calendar_name, summary, start_time_str, end_time_str)
        logging.info(f"✅ Event '{summary}' added to calendar '{calendar_name}' using template '{selected_template}'.")
        print(f"✅ Event '{summary}' added to calendar '{calendar_name}' using template '{selected_template}'.")
    except Exception as e:
        logging.error(f"❌ Failed to create event: {e}")
        print(f"❌ Failed to create event: {e}")


# ✅ Add Recurring Event
def add_recurring_event(calendar_name):
    """Add a recurring event to a specific calendar."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found.")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    summary = input("Enter event summary (e.g., Weekly Standup Meeting): ").strip()
    start_time = prompt_for_datetime("Enter start date and time")
    end_time = prompt_for_datetime("Enter end date and time")

    try:
        start_dt = datetime.fromisoformat(start_time)
        end_dt = datetime.fromisoformat(end_time)
        if start_dt >= end_dt:
            logging.warning("❌ Start time must be earlier than end time.")
            print("❌ Error: Start time must be earlier than end time.")
            return
    except ValueError as e:
        logging.error(f"❌ Invalid datetime format: {e}")
        print(f"❌ Invalid datetime format: {e}")
        return

    # Recurrence Pattern
    print("\nChoose a recurrence pattern:")
    print("1️⃣ Daily")
    print("2️⃣ Weekly")
    print("3️⃣ Monthly")
    print("4️⃣ Yearly")
    recurrence_choice = input("👉 Enter your choice: ").strip()

    frequency_mapping = {"1": "DAILY", "2": "WEEKLY", "3": "MONTHLY", "4": "YEARLY"}
    frequency = frequency_mapping.get(recurrence_choice)

    if not frequency:
        logging.warning("❌ Invalid recurrence choice.")
        print("❌ Invalid recurrence choice.")
        return

    end_condition = input("Should the recurrence end by date (d) or after a number of occurrences (n)? ").strip().lower()
    if end_condition == "d":
        end_date = input("Enter end date (YYYY-MM-DD): ").strip()
        recurrence_rule = f"RRULE:FREQ={frequency};UNTIL={end_date.replace('-', '')}T000000Z"
    elif end_condition == "n":
        count = input("Enter the number of occurrences: ").strip()
        if not count.isdigit():
            logging.warning("❌ Invalid number of occurrences.")
            print("❌ Invalid number of occurrences.")
            return
        recurrence_rule = f"RRULE:FREQ={frequency};COUNT={count}"
    else:
        logging.warning("❌ Invalid end condition choice.")
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
        logging.info(f"✅ Recurring event created: {created_event['htmlLink']}")
        print(f"✅ Recurring event created: {created_event['htmlLink']}")
    except Exception as e:
        logging.error(f"❌ Error creating recurring event: {e}")
        print(f"❌ Error creating recurring event: {e}")


# ✅ Search Events
def search_events(calendar_name):
    """Search events in a calendar by keyword or date range."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found.")
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
        logging.warning("❌ Invalid choice.")
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
            logging.info("❌ No matching events found.")
            print("❌ No matching events found.")
            return

        print("\n🔗 Matching Events:")
        for i, event in enumerate(events, start=1):
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(f"{i}. {event['summary']} | Start: {start} | ID: {event['id']}")

        return events

    except Exception as e:
        logging.error(f"❌ Error searching events: {e}")
        print(f"❌ Error searching events: {e}")


# ✅ Update Event
def update_event(calendar_name):
    """Update an existing event in a calendar."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found!")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    events = search_events(calendar_name)
    if not events:
        return

    event_index = input("\nEnter the number of the event you want to update: ").strip()
    if not event_index.isdigit() or int(event_index) < 1 or int(event_index) > len(events):
        logging.warning("❌ Invalid event selection.")
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
                logging.error("❌ Error: Start time must be earlier than end time.")
                print("❌ Error: Start time must be earlier than end time.")
                return
    except ValueError as e:
        logging.error(f"❌ Invalid datetime format: {e}")
        print(f"❌ Invalid datetime format: {e}")
        return

    # Update the event
    try:
        service.events().update(
            calendarId=calendar_id,
            eventId=event_id,
            body=updated_event
        ).execute()
        logging.info("✅ Event updated successfully.")
        print("✅ Event updated successfully.")
    except Exception as e:
        logging.error(f"❌ Error updating event: {e}")
        print(f"❌ Error updating event: {e}")

# ✅ Get Calendar Color ID
def get_calendar_color_id(calendar_name):
    """Retrieve the default color ID of a calendar."""
    try:
        calendars_list = service.calendarList().list().execute()
        for calendar in calendars_list["items"]:
            if calendar["summary"] == calendar_name:
                color_id = calendar.get("colorId", "1")  # Default to '1' if not found
                logging.info(f"🎨 Retrieved colorId '{color_id}' for calendar '{calendar_name}'.")
                return color_id
    except Exception as e:
        logging.error(f"❌ Failed to retrieve calendar color for '{calendar_name}': {e}")
        print(f"❌ Failed to retrieve calendar color for '{calendar_name}': {e}")
    return "1"  # Default color ID if retrieval fails


# ✅ Sync Event Colors
def sync_event_colors(calendar_name):
    """Sync event colors to match the calendar's assigned color."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found!")
        print(f"❌ Calendar '{calendar_name}' not found!")
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
                logging.info(f"✅ Updated color for event: {event_summary}")
                print(f"✅ Updated color for event: {event_summary}")
                updated_count += 1

        logging.info(f"🎨 Finished syncing colors. {updated_count} events updated.")
        print(f"🎨 Finished syncing colors. {updated_count} events updated.")
    except Exception as e:
        logging.error(f"❌ Error syncing event colors: {e}")
        print(f"❌ Error syncing event colors: {e}")


# ✅ Inspect Calendar Color
def inspect_calendar_color(calendar_name):
    """Inspect the colorId of a calendar."""
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found!")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    try:
        calendar = service.calendarList().get(calendarId=calendar_id).execute()
        color_id = calendar.get('colorId', 'Default')
        logging.info(f"🎨 Calendar '{calendar_name}' has colorId: {color_id}")
        print(f"🎨 Calendar '{calendar_name}' has colorId: {color_id}")
    except Exception as e:
        logging.error(f"❌ Failed to retrieve calendar color: {e}")
        print(f"❌ Failed to retrieve calendar color: {e}")


# ✅ Main Menu
def main():
    """Main menu for the script."""
    menu_options = {
        "1": list_calendars,
        "2": lambda: create_calendar(input("Enter calendar name: ").strip()),
        "3": lambda: add_event_with_multiple_dates(input("Enter calendar name: ").strip()),
        "4": lambda: add_multiple_unique_events(input("Enter calendar name: ").strip()),
        "5": lambda: import_from_csv(
            input("Enter calendar name: ").strip(),
            input("Enter the path to the CSV file: ").strip()
        ),
        "6": lambda: sync_event_colors(input("Enter calendar name: ").strip()),
        "7": lambda: inspect_calendar_color(input("Enter calendar name: ").strip()),
        "8": lambda: add_event_using_template(input("Enter calendar name: ").strip()),
        "9": lambda: add_recurring_event(input("Enter calendar name: ").strip()),
        "10": lambda: search_events(input("Enter calendar name: ").strip()),
        "11": lambda: update_event(input("Enter calendar name: ").strip()),
    }

    while True:
        print("\n🗓️  Google Calendar Manager")
        print("1️⃣  List Calendars")
        print("2️⃣  Create Calendar")
        print("3️⃣  Add Event with Multiple Dates/Times")
        print("4️⃣  Add Multiple Unique Events")
        print("5️⃣  Import Bulk Events from CSV")
        print("6️⃣  Sync Event Colors with Calendar Color")
        print("7️⃣  Inspect Calendar Color")
        print("8️⃣  Add Event from Template")
        print("9️⃣  Add Recurring Event")
        print("🔟  Search Events")
        print("1️⃣1️⃣  Update Event")
        print("🛑  Type 'exit' to quit.")
        
        choice = input("👉 Enter your choice: ").strip().lower()

        if choice == "exit":
            logging.info("👋 Exiting Google Calendar Manager.")
            print("👋 Goodbye!")
            break
        elif choice in menu_options:
            try:
                menu_options[choice]()
            except Exception as e:
                logging.error(f"❌ An error occurred while executing option {choice}: {e}")
                print(f"❌ An error occurred: {e}")
        else:
            logging.warning("❌ Invalid choice. Please try again.")
            print("❌ Invalid choice! Please try again.")


# ✅ Run the Script
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logging.warning("🛑 Program interrupted by user.")
        print("\n🛑 Program interrupted by user.")
    except Exception as e:
        logging.critical(f"❌ Critical error occurred: {e}")
        print(f"❌ Critical error occurred: {e}")
