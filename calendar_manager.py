# Import necessary modules and libraries for functionality
from datetime import datetime, timedelta  # Handle dates and time durations
import logging  # Enable logging for debugging and monitoring
import os  # Interact with the operating system (e.g., file paths, environment variables)
import sys  # Access system-specific parameters and functions
import csv  # Read and write CSV files
import json  # Parse JSON data
from dotenv import load_dotenv  # Load environment variables from a .env file
from googleapiclient.discovery import build  # Interact with Google APIs
from google.oauth2.credentials import Credentials  # Handle Google OAuth credentials
from google_auth_oauthlib.flow import (
    InstalledAppFlow,
)  # Manage OAuth 2.0 authentication flow
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError

# ✅ Load environment variables from the .env file
# Environment variables store sensitive data (like API keys) securely outside the codebase.
load_dotenv()

# ✅ Logging Configuration with Enhanced Error Handling
# Logs help track events, errors, and important information in the application.

# Define the directory and file for logging
LOG_DIR = "logs"
LOG_FILE = os.path.join(LOG_DIR, "calendar_manager.log")

# Ensure the log directory exists. If it doesn't, create it.
os.makedirs(LOG_DIR, exist_ok=True)

# Set up logging configuration
logging.basicConfig(
    level=logging.DEBUG,  # Log all messages from DEBUG level and above (DEBUG, INFO, WARNING, ERROR)
    format="%(asctime)s [%(levelname)s] %(message)s",  # Define the log message format
    handlers=[
        logging.FileHandler(
            LOG_FILE, encoding="utf-8"
        ),  # Save logs to a file with UTF-8 encoding
        logging.StreamHandler(stream=sys.stdout),  # Display logs in the console
    ],
)

# Log an initial debug message to confirm that logging is set up correctly
logging.debug("✅ Logging system initialized successfully.")

# ✅ Google Calendar Credentials
# Environment variables are used to securely store credentials and configuration for the Google Calendar API.

try:
    # Fetch required credentials and configurations from environment variables
    CLIENT_ID = os.getenv("GOOGLE_CALENDAR_CLIENT_ID")  # Google Calendar Client ID
    CLIENT_SECRET = os.getenv("GOOGLE_CALENDAR_CLIENT_SECRET")  # Client Secret
    TOKEN = os.getenv("GOOGLE_CALENDAR_TOKEN")  # Access Token for Google Calendar API
    REFRESH_TOKEN = os.getenv("GOOGLE_CALENDAR_REFRESH_TOKEN")  # Refresh Token
    TOKEN_URI = os.getenv(
        "GOOGLE_CALENDAR_TOKEN_URI", "https://oauth2.googleapis.com/token"
    )  # Token endpoint
    AUTH_URI = os.getenv(
        "GOOGLE_CALENDAR_AUTH_URI", "https://accounts.google.com/o/oauth2/auth"
    )  # Auth endpoint
    SCOPES = json.loads(
        os.getenv(
            "GOOGLE_CALENDAR_SCOPES", '["https://www.googleapis.com/auth/calendar"]'
        )
    )  # API scopes
    DEFAULT_TIMEZONE = os.getenv(
        "DEFAULT_TIMEZONE", "America/Chicago"
    )  # Default timezone for calendar events

    # Load additional configurations for calendars, colors, and templates
    calendars = json.loads(
        os.getenv("CALENDAR_NAMES", "{}")
    )  # Calendar names with associated colors
    color_map = json.loads(
        os.getenv("COLOR_MAP", "{}")
    )  # Color map for calendar events
    event_templates = json.loads(
        os.getenv("EVENT_TEMPLATES", "{}")
    )  # Predefined event templates

    # ✅ Validate Critical Environment Variables
    # Ensure that all required environment variables are present and correctly set.
    required_env_vars = [
        "GOOGLE_CALENDAR_CLIENT_ID",
        "GOOGLE_CALENDAR_CLIENT_SECRET",
        "GOOGLE_CALENDAR_TOKEN",
        "GOOGLE_CALENDAR_REFRESH_TOKEN",
        "DEFAULT_TIMEZONE",
    ]

    # Loop through each required variable and check if it is set
    for var in required_env_vars:
        if not os.getenv(var):
            # Raise an error if a required variable is missing
            raise EnvironmentError(f"❌ Missing required environment variable: {var}")

    # Log successful environment variable validation
    logging.debug("✅ Environment variables loaded and validated successfully.")

# Handle errors that might occur while loading or validating environment variables
except (json.JSONDecodeError, EnvironmentError) as e:
    # Log the error details
    logging.error(f"❌ Error with environment setup: {e}")
    # Exit the program because critical environment variables are missing or misconfigured
    raise SystemExit(
        "❌ Critical environment variables are missing or misconfigured. Exiting..."
    )


# ✅ Authentication for Google Calendar API
def authenticate_google_calendar():
    """
    Authenticate and return a Google Calendar API service instance.

    - Loads credentials from token.json if it exists.
    - Automatically refreshes expired tokens.
    - Falls back to a full OAuth flow if no valid credentials remain.
    """
    creds = None

    # 1️⃣ Try loading existing credentials
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        logging.debug("✅ Loaded credentials from token.json.")

        # 2️⃣ If expired, attempt a refresh
        if creds.expired and creds.refresh_token:
            try:
                logging.info("🔄 Token expired, attempting refresh...")
                creds.refresh(Request())
                with open("token.json", "w") as t:
                    t.write(creds.to_json())
                logging.info("✅ Token refreshed and saved to token.json.")
            except RefreshError as e:
                logging.warning(f"❌ Refresh failed ({e}), will re‐authenticate.")
                creds = None

    # 3️⃣ If we still don’t have valid creds, run the full OAuth flow
    if not creds or not creds.valid:
        try:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
            with open("token.json", "w") as t:
                t.write(creds.to_json())
            logging.info("✅ New credentials obtained and saved to token.json.")
        except Exception as e:
            logging.error(f"❌ OAuth flow failed: {e}")
            raise SystemExit("❌ Failed to authenticate with Google Calendar API.")

    # 4️⃣ Build and return the service
    logging.debug("✅ Google Calendar authentication successful.")
    return build("calendar", "v3", credentials=creds)


# initialize
service = authenticate_google_calendar()


# ✅ Validate Environment Variables
def validate_env_variables():
    """
    Validate that all required environment variables are set.

    Environment variables store sensitive information like API keys and tokens.
    This function ensures that all necessary variables are present and correctly set.
    """
    # Define a list of required environment variables
    required_vars = [
        "GOOGLE_CALENDAR_CLIENT_ID",  # Google API Client ID
        "GOOGLE_CALENDAR_CLIENT_SECRET",  # Google API Client Secret
        "GOOGLE_CALENDAR_TOKEN",  # Access token for Google Calendar API
        "GOOGLE_CALENDAR_REFRESH_TOKEN",  # Refresh token for long-term authentication
        "DEFAULT_TIMEZONE",  # Default timezone for calendar events
    ]

    # Check if any of the required variables are missing
    missing_vars = [var for var in required_vars if not os.getenv(var)]

    if missing_vars:
        # If there are missing variables, log and raise an error
        missing_vars_str = ", ".join(missing_vars)
        logging.error(f"❌ Missing required environment variables: {missing_vars_str}")
        raise EnvironmentError(
            f"❌ Missing required environment variables: {missing_vars_str}"
        )

    # Log a success message if all variables are validated
    logging.debug("✅ Environment variables validated successfully.")


# ✅ List Calendars
def list_calendars():
    """
    List all calendars in the user's Google Calendar account.

    This function fetches and displays all calendars associated with the authenticated Google account.
    """
    try:
        # ✅ Fetch the list of calendars from the user's Google account using the Google Calendar API
        calendars_list = service.calendarList().list().execute()

        # ✅ Check if the response contains any calendars
        if not calendars_list.get("items"):
            print("❌ No calendars found.")
            logging.warning("No calendars found in the user's calendar list.")
            return

        # ✅ Loop through the list of calendars and display their names and IDs
        print("\n📅 Available Calendars:")
        for calendar in calendars_list["items"]:
            print(
                f"- {calendar['summary']} (ID: {calendar['id']})"
            )  # Display calendar name and ID

        logging.debug("✅ Calendars listed successfully.")

    except Exception as e:
        # ✅ Catch and log any errors that happen while fetching calendars
        logging.error(f"❌ Failed to list calendars: {e}")
        print(f"❌ Error listing calendars: {e}")


# ✅ Create Calendar
def create_calendar(name):
    """
    Create a new calendar with the specified name.

    Args:
        name (str): The name of the new calendar.

    Returns:
        str: The ID of the created calendar if successful, otherwise None.
    """
    # ✅ Define the calendar properties: name and timezone
    calendar = {
        "summary": name,  # Set the calendar's display name
        "timeZone": DEFAULT_TIMEZONE,  # Use the default timezone from the environment variable
    }

    try:
        # ✅ Send a request to the Google Calendar API to create a new calendar
        created_calendar = service.calendars().insert(body=calendar).execute()

        # ✅ Print and log the success message with the new calendar's name and ID
        print(
            f"✅ Calendar created: {created_calendar['summary']} (ID: {created_calendar['id']})"
        )
        logging.debug(
            f"✅ Calendar '{name}' created successfully with ID: {created_calendar['id']}."
        )

        return created_calendar["id"]  # Return the ID of the created calendar

    except Exception as e:
        # ✅ Handle errors that may occur during calendar creation
        logging.error(f"❌ Error creating calendar '{name}': {e}")
        print(f"❌ Error creating calendar '{name}': {e}")

        return None


# ✅ Create Event
def create_event(calendar_name, summary, start_time, end_time):
    """
    Add an event to a specific calendar with proper color matching.

    Args:
        calendar_name (str): The name of the calendar where the event will be added.
        summary (str): A brief description or title for the event.
        start_time (str): Event start time in ISO 8601 format (e.g., "2024-06-01T10:00:00").
        end_time (str): Event end time in ISO 8601 format (e.g., "2024-06-01T11:00:00").
    """
    # ✅ Retrieve the calendar ID using the calendar's name
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.warning(f"❌ Calendar '{calendar_name}' not found!")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    try:
        # ✅ Validate that the start time is earlier than the end time
        start = datetime.fromisoformat(start_time)
        end = datetime.fromisoformat(end_time)
        if start >= end:
            logging.warning("❌ Error: Start time must be earlier than end time.")
            print("❌ Error: Start time must be earlier than end time.")
            return

        # ✅ Match the event's color based on the calendar's configuration
        # Get the color hex from calendar configuration, fallback to default if not found
        color_hex = calendars.get(calendar_name, "#236192")
        # Map the hex color to the corresponding Google Calendar color ID
        color_id = color_map.get(color_hex, "1")

        # ✅ Define the event's details
        event = {
            "summary": summary,  # The event's title or brief description
            "start": {
                "dateTime": start_time,
                "timeZone": DEFAULT_TIMEZONE,
            },  # Event start time with timezone
            "end": {
                "dateTime": end_time,
                "timeZone": DEFAULT_TIMEZONE,
            },  # Event end time with timezone
            "colorId": color_id,  # Apply the appropriate color ID to the event
            "reminders": {
                "useDefault": False,  # Override default reminders
                "overrides": [
                    {
                        "method": "email",
                        "minutes": 30,
                    },  # Email reminder 30 minutes before event
                    {
                        "method": "popup",
                        "minutes": 10,
                    },  # Popup reminder 10 minutes before event
                ],
            },
        }

        # ✅ Send the event details to the Google Calendar API to create the event
        created_event = (
            service.events().insert(calendarId=calendar_id, body=event).execute()
        )

        # ✅ Print and log a success message with a link to the created event
        print(f"✅ Event created: {created_event['htmlLink']}")
        logging.debug(
            f"✅ Event '{summary}' created successfully in calendar '{calendar_name}'."
        )

    except ValueError as ve:
        # ✅ Handle errors if the datetime format is incorrect
        logging.error(f"❌ Invalid datetime format: {ve}")
        print(f"❌ Invalid datetime format: {ve}")

    except Exception as e:
        # ✅ Catch and log any other errors during event creation
        logging.error(f"❌ Error creating event: {e}")
        print(f"❌ Error creating event: {e}")


# ✅ Prompt for Datetime
def prompt_for_datetime(prompt_text):
    """
    Prompt the user for a date and time in a structured format.

    Args:
        prompt_text (str): A message to guide the user (e.g., "Enter start date and time").

    Returns:
        str: An ISO 8601 formatted datetime string (e.g., "2024-06-01T10:00:00").

    Explanation:
        - The function asks the user to input year, month, day, hour, and minute individually.
        - It validates if the entered values form a valid date and time.
        - If the input is invalid (e.g., non-numeric or out-of-range values), it retries the prompt.
    """
    print(f"{prompt_text}:")  # Display the instruction to the user
    try:
        # ✅ Ask the user to input each part of the date and time
        year = int(input("  Year (e.g., 2024): ").strip())  # Get the year from the user
        month = int(input("  Month (1-12): ").strip())  # Get the month
        day = int(input("  Day (1-31): ").strip())  # Get the day
        hour = int(input("  Hour (0-23): ").strip())  # Get the hour
        minute = int(input("  Minute (0-59): ").strip())  # Get the minute

        # ✅ Create a datetime object using the provided inputs
        datetime_obj = datetime(year, month, day, hour, minute)

        # ✅ Return the datetime as an ISO 8601 formatted string
        return datetime_obj.isoformat()

    except ValueError as e:
        # ✅ Handle invalid input (e.g., letters instead of numbers, out-of-range values)
        logging.error(f"❌ Invalid datetime input: {e}")  # Log the error for debugging
        print(
            f"❌ Invalid datetime input: {e}"
        )  # Inform the user about the invalid input

        # ✅ Retry the prompt recursively until valid input is received
        return prompt_for_datetime(prompt_text)


# ✅ Get Calendar ID
def get_calendar_id(calendar_name):
    """
    Retrieve the calendar ID for a given calendar name.

    Args:
        calendar_name (str): The name of the calendar to search for.

    Returns:
        str: The calendar ID if found, otherwise None.

    Explanation:
        - The function fetches all calendars linked to the user's Google account.
        - It then loops through each calendar and checks if its name matches the given calendar_name.
        - If a match is found, the function returns the calendar's unique ID.
        - If no match is found or if an error occurs, it handles the situation gracefully.
    """
    try:
        # ✅ Fetch the list of calendars from the user's account
        calendars_list = service.calendarList().list().execute()

        # ✅ Loop through each calendar in the fetched list
        for calendar in calendars_list.get("items", []):
            # ✅ Check if the calendar name matches the given calendar_name
            if calendar["summary"] == calendar_name:
                # ✅ Log and return the calendar ID if found
                logging.debug(
                    f"✅ Found calendar '{calendar_name}' with ID: {calendar['id']}"
                )
                return calendar["id"]

        # ✅ If no calendar matches the given name, log a warning and inform the user
        logging.warning(f"❌ Calendar '{calendar_name}' not found.")
        print(f"❌ Calendar '{calendar_name}' not found.")
        return None  # Return None if the calendar isn't found

    except Exception as e:
        # ✅ Handle unexpected errors gracefully (e.g., API errors, connectivity issues)
        logging.error(f"❌ Error fetching calendar ID: {e}")  # Log the error
        print(f"❌ Error fetching calendar ID: {e}")  # Inform the user about the error
        return None  # Return None to indicate failure


# ✅ Add Event with Multiple Dates
def add_event_with_multiple_dates(calendar_name):
    """
    Add multiple occurrences of a single event with different dates and times.

    Args:
        calendar_name (str): The name of the calendar to add events to.

    Explanation:
        - Prompts the user for an event summary.
        - Allows adding multiple occurrences of the event with different dates and times.
        - Validates each date range before adding it to the event list.
        - Creates all the events in the specified calendar.
    """
    # ✅ Get the calendar ID based on the provided calendar name
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        return  # Exit if the calendar isn't found

    # ✅ Ask the user for a general event summary/title
    summary = input("Enter event summary (e.g., Twins vs Yankees): ").strip()
    events = []  # Create an empty list to store all event occurrences

    # ✅ Start a loop to collect multiple occurrences of the event
    while True:
        print("\nEnter details for an event occurrence:")

        # ✅ Prompt the user for start and end times
        start_time = prompt_for_datetime("Enter start date and time")
        end_time = prompt_for_datetime("Enter end date and time")

        try:
            # ✅ Validate that the start time is earlier than the end time
            start_dt = datetime.fromisoformat(start_time)
            end_dt = datetime.fromisoformat(end_time)
            if start_dt >= end_dt:
                print("❌ Error: Start time must be earlier than end time.")
                logging.warning("❌ Start time must be earlier than end time.")
                continue  # Restart the loop if the validation fails

        except ValueError as e:
            # ✅ Handle invalid datetime inputs
            print(f"❌ Invalid datetime format: {e}")
            logging.error(f"❌ Invalid datetime format: {e}")
            continue  # Restart the loop if the validation fails

        # ✅ Add the valid event occurrence to the list
        events.append(
            {"summary": summary, "start_time": start_time, "end_time": end_time}
        )

        # ✅ Ask the user if they want to add more occurrences
        more = input("Add another date/time for this event? (y/n): ").strip().lower()
        if more != "y":
            break  # Exit the loop if the user is done

    # ✅ Loop through the list of events and create them in the calendar
    for event in events:
        try:
            create_event(
                calendar_name, event["summary"], event["start_time"], event["end_time"]
            )
        except Exception as e:
            print(f"❌ Error creating event '{event['summary']}': {e}")
            logging.error(f"❌ Error creating event '{event['summary']}': {e}")

    # ✅ Confirm successful addition of events
    print(
        f"✅ {len(events)} occurrence(s) of '{summary}' added successfully to calendar '{calendar_name}'!"
    )
    logging.debug(
        f"✅ {len(events)} occurrences of '{summary}' added to calendar '{calendar_name}'."
    )


# ✅ Add Multiple Unique Events
def add_multiple_unique_events(calendar_name):
    """
    Continuously add unique events to a calendar until the user stops.

    Args:
        calendar_name (str): The name of the calendar where events will be added.

    Explanation:
        - Allows the user to add multiple unique events one by one.
        - Each event requires a summary, start time, and end time.
        - Validates that the start time is earlier than the end time.
        - Events are collected in a list and then added to the calendar one by one.
    """
    # ✅ Get the calendar ID based on the provided calendar name
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        return  # Exit if the calendar isn't found

    events = []  # Create an empty list to store all unique events

    # ✅ Start a loop to continuously add events until the user stops
    while True:
        # ✅ Prompt for event summary
        summary = input("\nEnter event summary (or type 'done' to finish): ").strip()
        if summary.lower() == "done":
            break  # Exit the loop if the user types 'done'

        # ✅ Prompt for start and end times
        start_time = prompt_for_datetime("Enter start date and time")
        end_time = prompt_for_datetime("Enter end date and time")

        try:
            # ✅ Validate that the start time is earlier than the end time
            start_dt = datetime.fromisoformat(start_time)
            end_dt = datetime.fromisoformat(end_time)
            if start_dt >= end_dt:
                print("❌ Error: Start time must be earlier than end time.")
                logging.warning("❌ Start time must be earlier than end time.")
                continue  # Restart the loop if the validation fails

        except ValueError as e:
            # ✅ Handle invalid datetime inputs
            print(f"❌ Invalid datetime format: {e}")
            logging.error(f"❌ Invalid datetime format: {e}")
            continue  # Restart the loop if the validation fails

        # ✅ Add the valid event to the list
        events.append(
            {"summary": summary, "start_time": start_time, "end_time": end_time}
        )

    # ✅ Loop through the list of unique events and create them in the calendar
    for event in events:
        try:
            create_event(
                calendar_name, event["summary"], event["start_time"], event["end_time"]
            )
        except Exception as e:
            print(f"❌ Error creating event '{event['summary']}': {e}")
            logging.error(f"❌ Error creating event '{event['summary']}': {e}")

    # ✅ Confirm successful addition of events
    print(
        f"✅ {len(events)} unique event(s) added successfully to calendar '{calendar_name}'!"
    )
    logging.debug(
        f"✅ {len(events)} unique events added to calendar '{calendar_name}'."
    )


# ✅ Import Events from CSV
def import_from_csv(calendar_name, csv_file):
    """
    Import events from a CSV file into a Google Calendar.

    Args:
        calendar_name (str): The name of the calendar to import events into.
        csv_file (str): Path to the CSV file containing event data.

    Explanation:
        - Reads a CSV file containing event data (Summary, Start Date, Start Time, End Date, End Time).
        - Validates each row to ensure required fields are provided.
        - Validates that start time is earlier than end time.
        - Adds valid events to the calendar one by one.
    """
    # ✅ Fetch the calendar ID based on the calendar name
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found.")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    # ✅ Retrieve the default color ID for the calendar
    color_id = get_calendar_color_id(calendar_name)

    try:
        # ✅ Open the CSV file for reading
        with open(csv_file, mode="r", newline="", encoding="utf-8") as file:
            reader = csv.DictReader(file)  # Read CSV rows as dictionaries

            # ✅ Ensure the CSV has valid headers
            if not reader.fieldnames:
                logging.error("❌ CSV file is empty or headers are missing.")
                print("❌ CSV file is empty or headers are missing.")
                return

            logging.info(f"✅ CSV Headers: {reader.fieldnames}")
            print("✅ CSV Headers:", reader.fieldnames)

            # ✅ Iterate through each row in the CSV file
            for row in reader:
                # ✅ Extract event details from the current row
                summary = row.get("Summary", "").strip()
                start_date = row.get("Start Date", "").strip()
                start_time = row.get("Start Time", "").strip()
                end_date = row.get("End Date", "").strip()
                end_time = row.get("End Time", "").strip()

                # ✅ Validate required fields
                if not all([summary, start_date, start_time, end_date, end_time]):
                    logging.warning(f"❌ Skipping invalid row (missing fields): {row}")
                    print(f"❌ Skipping invalid row (missing fields): {row}")
                    continue

                try:
                    # ✅ Create ISO 8601 formatted datetime strings
                    start_datetime = f"{start_date}T{start_time}"
                    end_datetime = f"{end_date}T{end_time}"

                    # ✅ Validate that start time is earlier than end time
                    start_dt = datetime.fromisoformat(start_datetime)
                    end_dt = datetime.fromisoformat(end_datetime)

                    if start_dt >= end_dt:
                        logging.warning(
                            f"❌ Skipping row with invalid time range: {row}"
                        )
                        print(f"❌ Skipping row with invalid time range: {row}")
                        continue

                    # ✅ Create the event payload to send to Google Calendar
                    event = {
                        "summary": summary,
                        "start": {
                            "dateTime": start_datetime,
                            "timeZone": DEFAULT_TIMEZONE,
                        },
                        "end": {"dateTime": end_datetime, "timeZone": DEFAULT_TIMEZONE},
                        "colorId": color_id,
                        "reminders": {
                            "useDefault": False,
                            "overrides": [
                                {
                                    "method": "email",
                                    "minutes": 30,
                                },  # Email reminder 30 minutes before
                                {
                                    "method": "popup",
                                    "minutes": 10,
                                },  # Popup reminder 10 minutes before
                            ],
                        },
                    }

                    # ✅ Add the event to Google Calendar
                    created_event = (
                        service.events()
                        .insert(calendarId=calendar_id, body=event)
                        .execute()
                    )
                    logging.info(f"✅ Event added: {summary} on {start_date}")
                    print(f"✅ Event added: {summary} on {start_date}")

                except ValueError as e:
                    # ✅ Handle invalid datetime formats
                    logging.error(
                        f"❌ Skipping row with invalid datetime format: {row}. Error: {e}"
                    )
                    print(
                        f"❌ Skipping row with invalid datetime format: {row}. Error: {e}"
                    )
                except KeyError as e:
                    # ✅ Handle missing keys in the row
                    logging.error(f"❌ Missing key in row: {row}. Error: {e}")
                    print(f"❌ Missing key in row: {row}. Error: {e}")

    except FileNotFoundError:
        # ✅ Handle if the CSV file doesn't exist
        logging.error("❌ CSV file not found. Please provide a valid path.")
        print("❌ CSV file not found. Please provide a valid path.")
    except Exception as e:
        # ✅ Handle any other unexpected errors
        logging.error(f"❌ Error processing CSV: {e}")
        print(f"❌ Error processing CSV: {e}")


# ✅ Add Event Using a Template
def add_event_using_template(calendar_name):
    """
    Add an event using a predefined template.

    Args:
        calendar_name (str): The name of the calendar where the event will be added.

    Explanation:
        - Displays a list of predefined event templates.
        - Allows the user to choose one template by number.
        - Prompts the user to enter the event start date and time.
        - Automatically calculates the end time based on the template's duration.
        - Creates and adds the event to the specified calendar.
    """
    # ✅ Fetch the calendar ID based on the calendar name
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found.")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    # ✅ Display a list of available event templates
    print("\n📝 Available Event Templates:")
    for idx, template_name in enumerate(event_templates.keys(), start=1):
        print(f"{idx}. {template_name}")

    # ✅ Prompt the user to select a template by number
    choice = input("Select a template by number: ").strip()
    try:
        template_index = int(choice) - 1
        template_keys = list(event_templates.keys())
        selected_template = template_keys[template_index]
    except (ValueError, IndexError):
        logging.warning("❌ Invalid choice! Please select a valid template.")
        print("❌ Invalid choice! Please select a valid template.")
        return

    # ✅ Fetch template details
    template = event_templates[selected_template]
    summary = template.get("summary", "Untitled Event")
    duration = template.get(
        "duration", 60
    )  # Default to 60 minutes if duration isn't provided

    # ✅ Prompt for event start date and time
    start_time = prompt_for_datetime("Enter event start date and time")
    try:
        start_dt = datetime.fromisoformat(start_time)
        end_dt = start_dt + timedelta(
            minutes=duration
        )  # Calculate end time based on duration

        start_time_str = start_dt.isoformat()
        end_time_str = end_dt.isoformat()
    except ValueError as e:
        logging.error(f"❌ Invalid datetime input: {e}")
        print(f"❌ Invalid datetime input: {e}")
        return

    # ✅ Create the event in the calendar
    try:
        create_event(calendar_name, summary, start_time_str, end_time_str)
        logging.info(
            f"✅ Event '{summary}' added using template '{selected_template}'."
        )
        print(f"✅ Event '{summary}' added using template '{selected_template}'.")
    except Exception as e:
        logging.error(f"❌ Failed to create event: {e}")
        print(f"❌ Failed to create event: {e}")


# ✅ Add Recurring Event
def add_recurring_event(calendar_name):
    """
    Add a recurring event to a specific calendar.

    Args:
        calendar_name (str): The name of the calendar where the recurring event will be added.

    Explanation:
        - Allows the user to create an event that repeats on a schedule (e.g., daily, weekly, monthly, yearly).
        - Prompts for event summary, start time, and end time.
        - Asks the user for a recurrence pattern and an end condition (by date or number of occurrences).
        - Creates and adds the recurring event to the specified calendar.
    """
    # ✅ Fetch the calendar ID using the calendar's name
    # This ensures we are adding the event to the correct calendar
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found.")  # Log the error
        print(f"❌ Calendar '{calendar_name}' not found!")  # Show an error to the user
        return

    # ✅ Gather basic event details from the user
    summary = input("Enter event summary (e.g., Weekly Standup Meeting): ").strip()
    start_time = prompt_for_datetime("Enter start date and time")
    end_time = prompt_for_datetime("Enter end date and time")

    try:
        # ✅ Validate the datetime inputs
        # Ensure the start time is earlier than the end time
        start_dt = datetime.fromisoformat(start_time)
        end_dt = datetime.fromisoformat(end_time)
        if start_dt >= end_dt:
            logging.warning(
                "❌ Start time must be earlier than end time."
            )  # Log warning
            print(
                "❌ Error: Start time must be earlier than end time."
            )  # Show an error to the user
            return
    except ValueError as e:
        # Handle invalid datetime formats entered by the user
        logging.error(f"❌ Invalid datetime format: {e}")
        print(f"❌ Invalid datetime format: {e}")
        return

    # ✅ Ask the user for a recurrence pattern (Daily, Weekly, Monthly, Yearly)
    print("\nChoose a recurrence pattern:")
    print("1️⃣ Daily")
    print("2️⃣ Weekly")
    print("3️⃣ Monthly")
    print("4️⃣ Yearly")
    recurrence_choice = input("👉 Enter your choice: ").strip()

    # ✅ Map the user's choice to a valid recurrence frequency
    frequency_mapping = {"1": "DAILY", "2": "WEEKLY", "3": "MONTHLY", "4": "YEARLY"}
    frequency = frequency_mapping.get(recurrence_choice)

    if not frequency:
        # Handle invalid recurrence choice
        logging.warning("❌ Invalid recurrence choice.")  # Log the error
        print("❌ Invalid recurrence choice.")  # Show an error to the user
        return

    # ✅ Ask the user how the recurrence should end: by a date or a number of occurrences
    end_condition = (
        input(
            "Should the recurrence end by date (d) or after a number of occurrences (n)? "
        )
        .strip()
        .lower()
    )

    if end_condition == "d":
        # ✅ End recurrence by a specific date
        end_date = input("Enter end date (YYYY-MM-DD): ").strip()
        recurrence_rule = (
            f"RRULE:FREQ={frequency};UNTIL={end_date.replace('-', '')}T000000Z"
        )
    elif end_condition == "n":
        # ✅ End recurrence after a specific number of occurrences
        count = input("Enter the number of occurrences: ").strip()
        if not count.isdigit():
            logging.warning("❌ Invalid number of occurrences.")  # Log warning
            print("❌ Invalid number of occurrences.")  # Show an error
            return
        recurrence_rule = f"RRULE:FREQ={frequency};COUNT={count}"
    else:
        # Handle invalid end condition choice
        logging.warning("❌ Invalid end condition choice.")  # Log the error
        print("❌ Invalid end condition choice.")  # Show an error
        return

    # ✅ Build the recurring event payload
    event = {
        "summary": summary,  # Title or description of the event
        "start": {
            "dateTime": start_time,
            "timeZone": DEFAULT_TIMEZONE,
        },  # Event start time
        "end": {"dateTime": end_time, "timeZone": DEFAULT_TIMEZONE},  # Event end time
        "recurrence": [
            recurrence_rule
        ],  # Add the recurrence rule (Daily, Weekly, etc.)
        "reminders": {
            "useDefault": False,
            "overrides": [
                {
                    "method": "email",
                    "minutes": 30,
                },  # Email reminder 30 minutes before event
                {
                    "method": "popup",
                    "minutes": 10,
                },  # Popup reminder 10 minutes before event
            ],
        },
    }

    # ✅ Send the event data to Google Calendar API to create the recurring event
    try:
        created_event = (
            service.events().insert(calendarId=calendar_id, body=event).execute()
        )
        logging.info(
            f"✅ Recurring event created: {created_event['htmlLink']}"
        )  # Log success
        print(
            f"✅ Recurring event created: {created_event['htmlLink']}"
        )  # Show success to the user
    except Exception as e:
        # Handle errors during event creation
        logging.error(f"❌ Error creating recurring event: {e}")
        print(f"❌ Error creating recurring event: {e}")


# ✅ Search Events
def search_events(calendar_name):
    """
    Search events in a calendar by keyword or date range.

    Args:
        calendar_name (str): The name of the calendar to search events in.

    Explanation:
        - The user can search events either by keyword or within a specific date range.
        - If the search is by keyword, it looks for matching text in event summaries.
        - If the search is by date range, it filters events within the specified start and end dates.
        - Matching events are displayed with their summary, start time, and event ID.
    """
    # ✅ Fetch the calendar ID using the calendar's name
    # This ensures we are working with the correct calendar
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(
            f"❌ Calendar '{calendar_name}' not found."
        )  # Log an error if calendar not found
        print(f"❌ Calendar '{calendar_name}' not found!")  # Show an error to the user
        return

    # ✅ Present search options to the user
    # The user can choose to search by keyword or by date range
    print("\n🔍 Search Options:")
    print("1️⃣ Search by keyword")  # Search for events based on text in their summaries
    print("2️⃣ Search by date range")  # Search for events in a specific date range
    choice = input("👉 Enter your choice: ").strip()

    # ✅ Initialize variables for search criteria
    search_query = None  # Used for keyword search
    time_min = None  # Start date for date range search
    time_max = None  # End date for date range search

    if choice == "1":
        # ✅ Option 1: Search by keyword
        search_query = input("Enter keyword to search in event summaries: ").strip()
    elif choice == "2":
        # ✅ Option 2: Search by date range
        # The user will input a start and end date for the search
        time_min = input("Enter start date (YYYY-MM-DD): ").strip()
        time_max = input("Enter end date (YYYY-MM-DD): ").strip()
    else:
        # ✅ Handle invalid search choice
        logging.warning("❌ Invalid choice.")  # Log a warning for an invalid choice
        print("❌ Invalid choice.")  # Show an error message to the user
        return

    try:
        # ✅ Fetch events from Google Calendar API based on search criteria
        events_result = (
            service.events()
            .list(
                calendarId=calendar_id,  # Search within the selected calendar
                q=search_query,  # Filter by keyword if provided
                timeMin=(
                    f"{time_min}T00:00:00Z" if time_min else None
                ),  # Start of date range
                timeMax=(
                    f"{time_max}T23:59:59Z" if time_max else None
                ),  # End of date range
                singleEvents=True,  # Ensures recurring events are expanded into single instances
                orderBy="startTime",  # Sort events by their start time
            )
            .execute()
        )

        # ✅ Retrieve the list of events from the API response
        events = events_result.get("items", [])
        if not events:
            # ✅ Handle case where no events match the criteria
            logging.info("❌ No matching events found.")  # Log info for no results
            print("❌ No matching events found.")  # Inform the user
            return

        # ✅ Display matching events to the user
        print("\n🔗 Matching Events:")
        for i, event in enumerate(events, start=1):
            # ✅ Get the event's start time (dateTime or date if it's an all-day event)
            start = event["start"].get("dateTime", event["start"].get("date"))
            print(f"{i}. {event['summary']} | Start: {start} | ID: {event['id']}")

        # ✅ Return the list of events for further processing, if needed
        return events

    except Exception as e:
        # ✅ Handle unexpected errors during the event search
        logging.error(f"❌ Error searching events: {e}")  # Log the exception
        print(
            f"❌ Error searching events: {e}"
        )  # Display the error message to the user


# ✅ Update Event
def update_event(calendar_name):
    """
    Update an existing event in a calendar.

    Args:
        calendar_name (str): The name of the calendar where the event exists.

    Explanation:
        - First, the user searches for events in the calendar using `search_events`.
        - The user selects an event by its number from the search results.
        - The user can update the event's summary, start time, end time, and description.
        - If no updates are provided, the existing values are retained.
    """
    # ✅ Step 1: Fetch the calendar ID using the calendar name
    # The calendar ID is needed to identify and modify events in the correct calendar.
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(
            f"❌ Calendar '{calendar_name}' not found!"
        )  # Log an error if calendar is not found
        print(f"❌ Calendar '{calendar_name}' not found!")  # Inform the user
        return

    # ✅ Step 2: Search for events in the selected calendar
    # The `search_events` function allows the user to list events and select one for updating.
    events = search_events(calendar_name)
    if not events:
        return  # Exit if no events are found or if the search was unsuccessful

    # ✅ Step 3: Ask the user to select an event to update
    event_index = input("\nEnter the number of the event you want to update: ").strip()
    if (
        not event_index.isdigit()
        or int(event_index) < 1
        or int(event_index) > len(events)
    ):
        logging.warning(
            "❌ Invalid event selection."
        )  # Log a warning for invalid selection
        print("❌ Invalid event selection.")  # Show an error to the user
        return

    # ✅ Step 4: Retrieve the selected event details
    selected_event = events[
        int(event_index) - 1
    ]  # Get the selected event using its index
    event_id = selected_event["id"]  # Extract the event ID for updating

    # ✅ Step 5: Prompt the user for updated event details
    print("\n🛠️ Update Event Details:")
    summary = input(
        f"New summary (leave blank to keep '{selected_event['summary']}'): "
    ).strip()
    start_time = prompt_for_datetime(
        "Enter new start date and time (leave blank to keep current)"
    )
    end_time = prompt_for_datetime(
        "Enter new end date and time (leave blank to keep current)"
    )
    description = input("Enter new description (leave blank to keep current): ").strip()

    # ✅ Step 6: Build the updated event data structure
    updated_event = {
        "summary": summary
        or selected_event["summary"],  # Keep the current summary if none is provided
        "start": selected_event["start"],  # Default to the current start time
        "end": selected_event["end"],  # Default to the current end time
        "description": description
        or selected_event.get("description", ""),  # Update description if provided
    }

    # ✅ Step 7: Update start and end times if the user provided new ones
    if start_time:
        updated_event["start"] = {"dateTime": start_time, "timeZone": DEFAULT_TIMEZONE}
    if end_time:
        updated_event["end"] = {"dateTime": end_time, "timeZone": DEFAULT_TIMEZONE}

    # ✅ Step 8: Validate the time range if both start and end times are updated
    try:
        if start_time and end_time:
            start_dt = datetime.fromisoformat(start_time)  # Parse start time
            end_dt = datetime.fromisoformat(end_time)  # Parse end time
            if start_dt >= end_dt:
                logging.error(
                    "❌ Error: Start time must be earlier than end time."
                )  # Log the error
                print(
                    "❌ Error: Start time must be earlier than end time."
                )  # Inform the user
                return
    except ValueError as e:
        # ✅ Handle invalid datetime formats
        logging.error(f"❌ Invalid datetime format: {e}")
        print(f"❌ Invalid datetime format: {e}")
        return

    # ✅ Step 9: Send the updated event details to Google Calendar API
    try:
        service.events().update(
            calendarId=calendar_id,  # Use the fetched calendar ID
            eventId=event_id,  # Specify the event to update using its ID
            body=updated_event,  # Pass the updated event details
        ).execute()

        logging.info("✅ Event updated successfully.")  # Log success message
        print("✅ Event updated successfully.")  # Inform the user

    except Exception as e:
        # ✅ Handle errors during the event update process
        logging.error(f"❌ Error updating event: {e}")  # Log the error
        print(f"❌ Error updating event: {e}")  # Display the error to the user


# ✅ Get Calendar Color ID
def get_calendar_color_id(calendar_name):
    """
    Retrieve the default color ID of a calendar.

    Args:
        calendar_name (str): The name of the calendar to fetch the color ID.

    Explanation:
        - Each Google Calendar can have a default color assigned to it.
        - This function fetches the `colorId` for the given calendar.
        - If no color is assigned, it defaults to "1".
    """
    try:
        # ✅ Step 1: Fetch all calendars associated with the user account
        calendars_list = service.calendarList().list().execute()

        # ✅ Step 2: Loop through all calendars to find the matching calendar by name
        for calendar in calendars_list["items"]:
            if calendar["summary"] == calendar_name:
                # ✅ Step 3: Retrieve the `colorId` from the calendar
                # If no color is set, default to "1"
                color_id = calendar.get("colorId", "1")
                logging.info(
                    f"🎨 Retrieved colorId '{color_id}' for calendar '{calendar_name}'."
                )
                return color_id  # Return the found colorId

    except Exception as e:
        # ✅ Step 4: Handle any errors during the API call
        logging.error(
            f"❌ Failed to retrieve calendar color for '{calendar_name}': {e}"
        )
        print(f"❌ Failed to retrieve calendar color for '{calendar_name}': {e}")

    # ✅ Step 5: If the calendar isn't found or an error occurs, return the default color "1"
    return "1"


# ✅ Sync Event Colors
def sync_event_colors(calendar_name):
    """
    Sync event colors to match the calendar's assigned color.

    Args:
        calendar_name (str): The name of the calendar whose events need color syncing.

    Explanation:
        - Fetches all events from the specified calendar.
        - Compares each event's color with the calendar's default color.
        - If an event's color does not match, it updates the event to use the calendar's default color.
    """
    # ✅ Step 1: Get the calendar ID for the specified calendar name
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        logging.error(f"❌ Calendar '{calendar_name}' not found!")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    # ✅ Step 2: Fetch the calendar's designated color from the configuration
    # `calendars` and `color_map` are predefined mappings in the environment variables
    color_hex = calendars.get(
        calendar_name, "#236192"
    )  # Default to `#236192` if not in .env
    color_id = color_map.get(
        color_hex, "1"
    )  # Map the hex color to a Google Calendar `colorId`, defaulting to "1"

    try:
        # ✅ Step 3: Fetch all events from the calendar
        events_result = service.events().list(calendarId=calendar_id).execute()
        events = events_result.get(
            "items", []
        )  # Extract the list of events from the response

        updated_count = 0  # Counter to track how many events were updated

        # ✅ Step 4: Loop through each event in the calendar
        for event in events:
            event_id = event.get("id")  # Unique identifier for the event
            event_summary = event.get(
                "summary", "No Summary"
            )  # Event title, fallback to 'No Summary'
            event_color_id = event.get(
                "colorId", None
            )  # Current event colorId, if available

            # ✅ Step 5: Check if the event color matches the calendar's default color
            if event_color_id != color_id:
                # Update the event's color to match the calendar's default color
                event["colorId"] = color_id
                service.events().update(
                    calendarId=calendar_id, eventId=event_id, body=event
                ).execute()

                logging.info(f"✅ Updated color for event: {event_summary}")
                print(f"✅ Updated color for event: {event_summary}")
                updated_count += 1  # Increment the updated event counter

        # ✅ Step 6: Display a summary of the changes
        logging.info(f"🎨 Finished syncing colors. {updated_count} events updated.")
        print(f"🎨 Finished syncing colors. {updated_count} events updated.")

    except Exception as e:
        # ✅ Step 7: Handle errors gracefully
        logging.error(f"❌ Error syncing event colors: {e}")
        print(f"❌ Error syncing event colors: {e}")


# ✅ Inspect Calendar Color
def inspect_calendar_color(calendar_name):
    """
    Inspect the `colorId` of a calendar.

    Args:
        calendar_name (str): The name of the calendar to inspect.

    Explanation:
        - Each Google Calendar can have a specific `colorId`.
        - This function retrieves and displays the `colorId` of the given calendar.
        - If the calendar isn't found or an error occurs, it handles the issue gracefully.
    """
    # ✅ Step 1: Get the calendar ID using its name
    calendar_id = get_calendar_id(calendar_name)
    if not calendar_id:
        # Log and print an error if the calendar doesn't exist
        logging.error(f"❌ Calendar '{calendar_name}' not found!")
        print(f"❌ Calendar '{calendar_name}' not found!")
        return

    try:
        # ✅ Step 2: Fetch calendar details from Google Calendar API
        calendar = service.calendarList().get(calendarId=calendar_id).execute()

        # ✅ Step 3: Extract the colorId from the calendar details
        color_id = calendar.get(
            "colorId", "Default"
        )  # Default to "Default" if not present

        # ✅ Step 4: Display the retrieved colorId
        logging.info(f"🎨 Calendar '{calendar_name}' has colorId: {color_id}")
        print(f"🎨 Calendar '{calendar_name}' has colorId: {color_id}")

    except Exception as e:
        # ✅ Step 5: Handle any errors during API interaction
        logging.error(f"❌ Failed to retrieve calendar color: {e}")
        print(f"❌ Failed to retrieve calendar color: {e}")


# ✅ Main Menu
def main():
    """
    Main menu for the script.

    Explanation:
        - This function provides an interactive menu for the user.
        - Each menu option corresponds to a specific calendar functionality.
        - The user can interactively choose an option, and the corresponding function will run.
        - The user can exit the menu by typing 'exit'.
    """
    # ✅ Step 1: Define menu options and their corresponding functions
    menu_options = {
        "1": list_calendars,  # List all calendars
        "2": lambda: create_calendar(
            input("Enter calendar name: ").strip()
        ),  # Create a new calendar
        "3": lambda: add_event_with_multiple_dates(
            input("Enter calendar name: ").strip()
        ),  # Add events with multiple dates
        "4": lambda: add_multiple_unique_events(
            input("Enter calendar name: ").strip()
        ),  # Add multiple unique events
        "5": lambda: import_from_csv(
            input("Enter calendar name: ").strip(),
            input("Enter the path to the CSV file: ").strip(),
        ),  # Import events from a CSV file
        "6": lambda: sync_event_colors(
            input("Enter calendar name: ").strip()
        ),  # Sync event colors
        "7": lambda: inspect_calendar_color(
            input("Enter calendar name: ").strip()
        ),  # Inspect calendar color
        "8": lambda: add_event_using_template(
            input("Enter calendar name: ").strip()
        ),  # Add event using a template
        "9": lambda: add_recurring_event(
            input("Enter calendar name: ").strip()
        ),  # Add a recurring event
        "10": lambda: search_events(
            input("Enter calendar name: ").strip()
        ),  # Search events
        "11": lambda: update_event(
            input("Enter calendar name: ").strip()
        ),  # Update an event
    }

    # ✅ Step 2: Display the menu options and handle user input
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

        # ✅ Step 3: Get the user's choice
        choice = input("👉 Enter your choice: ").strip().lower()

        # ✅ Step 4: Handle user input
        if choice == "exit":
            # Exit gracefully
            logging.info("👋 Exiting Google Calendar Manager.")
            print("👋 Goodbye!")
            break
        elif choice in menu_options:
            # ✅ Step 5: Execute the chosen menu option
            try:
                menu_options[choice]()
            except Exception as e:
                # Log and print any errors that occur while running the selected function
                logging.error(
                    f"❌ An error occurred while executing option {choice}: {e}"
                )
                print(f"❌ An error occurred: {e}")
        else:
            # Handle invalid input
            logging.warning("❌ Invalid choice. Please try again.")
            print("❌ Invalid choice! Please try again.")


# ✅ Run the Script
if __name__ == "__main__":
    """
    Entry point of the program.

    Explanation:
        - The `main()` function is executed when the script is run directly.
        - Handles graceful exits and unexpected errors.
    """
    try:
        # ✅ Start the main menu
        main()
    except KeyboardInterrupt:
        # ✅ Handle user interruption (Ctrl+C)
        logging.warning("🛑 Program interrupted by user.")
        print("\n🛑 Program interrupted by user.")
    except Exception as e:
        # ✅ Handle unexpected critical errors
        logging.critical(f"❌ Critical error occurred: {e}")
        print(f"❌ Critical error occurred: {e}")
