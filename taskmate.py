import os
import sys
import asyncio
import json
import logging
from dotenv import load_dotenv
from datetime import datetime, timedelta
from dateutil.parser import parse as parse_date
from typing import List, Dict, Any
from msal import PublicClientApplication, SerializableTokenCache
from openai import OpenAI, OpenAIError, RateLimitError, APIConnectionError
from rich.console import Console
from rich.prompt import Prompt
from rich.table import Table
from cachetools import TTLCache
import aiohttp
from dateparser import parse as parse_natural_date
import pytz
from prompt_toolkit import PromptSession
from prompt_toolkit.styles import Style
from prompt_toolkit.formatted_text import HTML
from prompt_toolkit.completion import WordCompleter
from prompt_toolkit.history import FileHistory
from prompt_toolkit.key_binding import KeyBindings
from pydantic import BaseModel

def get_custom_prompt(current_datetime):
    return f'Calendar Assistant ({current_datetime})> '

def get_history():
    return FileHistory('.calendar_assistant_history')

def get_key_bindings():
    kb = KeyBindings()
    
    @kb.add('c-d')
    def _(event):
        "Exit when 'c-d' is pressed."
        event.app.exit()
    
    @kb.add('c-l')
    def _(event):
        "Clear the screen when 'c-l' is pressed."
        event.app.current_buffer.text = ''
    
    return kb

logging.basicConfig(
    filename='chatbot.log',
    filemode='a',
    format='%(asctime)s %(levelname)s:%(message)s',
    level=logging.DEBUG
)
logger = logging.getLogger(__name__)

console = Console()

load_dotenv()



OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
OPENAI_PACKAGE_VERSION = os.getenv('OPENAI_PACKAGE_VERSION', '1.0.0')

CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')

AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPES = ['Calendars.ReadWrite']

missing_vars = []
if not OPENAI_API_KEY:
    missing_vars.append('OPENAI_API_KEY')
if not CLIENT_ID:
    missing_vars.append('CLIENT_ID')
if not TENANT_ID:
    missing_vars.append('TENANT_ID')

if missing_vars:
    console.print(f"[bold red]Error: Missing environment variables: {', '.join(missing_vars)}[/bold red]")
    sys.exit(1)

client = OpenAI(api_key=OPENAI_API_KEY)

token_cache = SerializableTokenCache()

app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    token_cache=token_cache
)


def get_access_token() -> str:
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        result = app.acquire_token_interactive(SCOPES)

    if 'access_token' in result:
        return result['access_token']
    else:
        error_description = result.get('error_description', 'Unknown error')
        console.print(f"[bold red]Error acquiring token: {error_description}[/bold red]")
        logger.error(f"Error acquiring token: {error_description}")
        sys.exit(1)


conversation_history: List[Dict[str, Any]] = []


def trim_conversation_history(history: List[Dict[str, Any]], max_length: int = 10) -> List[Dict[str, Any]]:
    return history[-max_length:]


def parse_natural_language_date(date_str: str, tz_str: str = 'America/New_York') -> datetime:
    try:
        timezone = pytz.timezone(tz_str)
    except pytz.UnknownTimeZoneError:
        logger.error(f"Unknown timezone: {tz_str}")
        raise ValueError(f"Unknown timezone: {tz_str}")

    logger.debug(f"Parsing date string: '{date_str}' with RELATIVE_BASE '{datetime.now(timezone).isoformat()}'")

    parsed_date = parse_natural_date(
        date_str,
        settings={
            'RETURN_AS_TIMEZONE_AWARE': True,
            'RELATIVE_BASE': datetime.now(timezone),
            'TIMEZONE': tz_str,
            'TO_TIMEZONE': tz_str
        }
    )

    if not parsed_date:
        logger.error(f"Unable to parse date: {date_str}")
        raise ValueError(f"Unable to parse date: {date_str}")

    # Ensure the parsed date is timezone-aware
    if parsed_date.tzinfo is None:
        parsed_date = timezone.localize(parsed_date)

    logger.debug(f"Parsed date: {parsed_date.isoformat()}")
    return parsed_date


def format_datetime_us(dt_str: str) -> str:
    dt = parse_date(dt_str)
    if dt:
        return dt.strftime('%A, %B %d, %Y at %I:%M %p')
    return ''


id_cache = TTLCache(maxsize=1000, ttl=3600)


async def cache_calendars() -> List[Dict[str, Any]]:
    access_token = get_access_token()
    headers = {'Authorization': f'Bearer {access_token}'}
    url = 'https://graph.microsoft.com/v1.0/me/calendars?$select=id,name,isDefaultCalendar'
    calendars = []
    async with aiohttp.ClientSession() as session:
        while url:
            try:
                async with session.get(url, headers=headers) as response:
                    if response.status == 200:
                        data = await response.json()
                        calendars.extend(data.get('value', []))
                        url = data.get('@odata.nextLink')
                    else:
                        error_text = await response.text()
                        console.print(f"[bold red]Failed to get calendars: {response.status} {error_text}[/bold red]")
                        logger.error(f"Failed to get calendars: {response.status} {error_text}")
                        return []
            except aiohttp.ClientError as e:
                console.print(f"[bold red]Request exception while fetching calendars: {e}[/bold red]")
                logger.exception("Request exception while fetching calendars")
                return []
    
    if not calendars:
        logger.error("No calendars retrieved from Microsoft 365")
        console.print("[bold red]Failed to retrieve any calendars. Please check your account settings.[/bold red]")
        return []
    
    for calendar in calendars:
        id_cache[calendar['name'].lower()] = calendar['id']
        if calendar.get('isDefaultCalendar'):
            id_cache['default'] = calendar['id']
    logger.debug(f"Cached calendars: {id_cache}")
    return calendars

def get_default_calendar_id() -> str:
    default_id = id_cache.get('default')
    if default_id:
        return default_id
    # If no default calendar is set, return the first calendar in the cache
    return next(iter(id_cache.values()), None)


async def make_api_call(method: str, url: str, headers: Dict[str, str], **kwargs) -> Dict[str, Any]:
    async with aiohttp.ClientSession() as session:
        func = getattr(session, method.lower())
        try:
            if 'json' in kwargs:
                logger.debug(f"API Request JSON: {json.dumps(kwargs['json'], indent=4)}")
            async with func(url, headers=headers, **kwargs) as response:
                text = await response.text()
                logger.debug(f"API Response Status: {response.status}, Body: {text}")
                if response.status in (200, 201, 204):
                    return {} if response.status == 204 else json.loads(text)
                else:
                    logger.error(f"API call failed: {response.status} {text}")
                    return {'error': f"{response.status}: {text}"}
        except Exception as e:
            logger.exception("API call exception")
            return {'error': str(e)}


def format_event_preview(event_data: Dict[str, Any]) -> None:
    table = Table(title="Event Preview")
    table.add_column("Field", style="bold")
    table.add_column("Value")

    for field in ['subject', 'body', 'start', 'end', 'location', 'attendees',
                  'recurrence', 'reminderMinutesBeforeStart', 'categories', 'isAllDay']:
        value = event_data.get(field)

        if field in ('start', 'end'):
            if event_data.get('isAllDay', False):
                date = value.get('date', '')
                value = f"{date} (All day)"
            else:
                date_time = value.get('dateTime', '')
                time_zone = value.get('timeZone', 'UTC')
                try:
                    formatted_dt = format_datetime_us(date_time)
                    value = f"{formatted_dt} ({time_zone})"
                except Exception as e:
                    value = f"{date_time} ({time_zone})"
                    logger.error(f"Error formatting datetime for field '{field}': {e}")
        elif field == 'location' and isinstance(value, dict):
            value = value.get('displayName', '')
        elif field == 'attendees':
            if isinstance(value, list) and value:
                attendees_formatted = ', '.join([att.get('emailAddress', {}).get('address', '') for att in value])
                value = attendees_formatted if attendees_formatted else 'None'
            else:
                value = 'None'
        elif field == 'recurrence' and not value:
            value = 'None'
        elif field == 'categories' and not value:
            value = 'None'
        elif field == 'reminderMinutesBeforeStart' and not value:
            value = 'None'
        elif field == 'isAllDay':
            value = 'Yes' if value else 'No'

        table.add_row(field.capitalize(), str(value))

    console.print(table)


def validate_event_data(event_data: Dict[str, Any]) -> bool:
    required_fields = ['subject', 'start', 'end']
    for field in required_fields:
        if field not in event_data:
            console.print(f"[bold red]Error: '{field}' is a required field for events.[/bold red]")
            return False
    
    # Additional validation checks
    if not event_data['subject'].strip():
        console.print("[bold red]Error: 'subject' cannot be empty.[/bold red]")
        return False
    
    try:
        start = parse_natural_language_date(event_data['start'])
        end = parse_natural_language_date(event_data['end'])
        if end <= start:
            console.print("[bold red]Error: End time must be after start time.[/bold red]")
            return False
    except ValueError as ve:
        console.print(f"[bold red]Error: Invalid date format - {ve}[/bold red]")
        return False
    
    return True


class CalendarEvent(BaseModel):
    subject: str
    start: str
    end: str
    location: str = ""
    attendees: List[str] = []
    reminderMinutesBeforeStart: int = 15
    categories: List[str] = []
    isAllDay: bool = False
    body: str = ""

import time

async def create_event(event_data: dict, tz_str: str = 'America/New_York', max_retries: int = 3) -> None:
    try:
        event_data = CalendarEvent(**event_data)
    except ValueError as e:
        console.print(f"[bold red]Error: Invalid event data provided - {e}[/bold red]")
        return

    start_datetime = parse_natural_language_date(event_data.start, tz_str=tz_str)
    end_datetime = parse_natural_language_date(event_data.end, tz_str=tz_str)

    new_event = {
        'subject': event_data.subject,
        'body': {
            'contentType': 'text',
            'content': event_data.body
        },
        'isAllDay': event_data.isAllDay,
        'start': {
            'dateTime': start_datetime.isoformat() if not event_data.isAllDay else start_datetime.date().isoformat(),
            'timeZone': tz_str
        },
        'end': {
            'dateTime': end_datetime.isoformat() if not event_data.isAllDay else end_datetime.date().isoformat(),
            'timeZone': tz_str
        }
    }

    if event_data.location:
        new_event['location'] = {'displayName': event_data.location}

    if event_data.attendees:
        new_event['attendees'] = [{'emailAddress': {'address': email}} for email in event_data.attendees]

    if event_data.reminderMinutesBeforeStart is not None:
        new_event['reminderMinutesBeforeStart'] = event_data.reminderMinutesBeforeStart

    if event_data.categories:
        new_event['categories'] = event_data.categories

    logger.debug(f"Event payload: {json.dumps(new_event, indent=2)}")

    console.print("\n[bold yellow]Event Preview:[/bold yellow]")
    format_event_preview(new_event)

    calendars = await cache_calendars()
    console.print("\n[bold cyan]Available Calendars:[/bold cyan]")
    for idx, calendar in enumerate(calendars, start=1):
        console.print(f"{idx}. {calendar['name']}")
    
    calendar_completer = WordCompleter([str(i) for i in range(1, len(calendars)+1)])
    session = PromptSession()
    while True:
        calendar_choice = await session.prompt_async("Select a calendar (number): ", completer=calendar_completer)
        try:
            calendar_index = int(calendar_choice) - 1
            if 0 <= calendar_index < len(calendars):
                selected_calendar = calendars[calendar_index]
                break
            else:
                console.print("[bold red]Invalid selection. Please try again.[/bold red]")
        except ValueError:
            console.print("[bold red]Please enter a valid number.[/bold red]")

    calendar_id = selected_calendar['id']

    for attempt in range(max_retries):
        try:
            access_token = get_access_token()
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events'
            response = await make_api_call('POST', url, headers, json=new_event)

            if 'error' in response:
                error_code = response['error'].get('code', '')
                if error_code == 'UnableToDeserializePostBody':
                    logger.error(f"Deserialization error. Payload: {json.dumps(new_event, indent=2)}")
                    console.print("[bold red]Error: The event data couldn't be processed. Please check the event details.[/bold red]")
                    return
                raise Exception(f"Failed to create event: {response['error']}")

            console.print(f"[bold green]Event '{event_data.subject}' created successfully in calendar '{selected_calendar['name']}'.[/bold green]")
            logger.info(f"Event '{event_data.subject}' created successfully in calendar '{selected_calendar['name']}'.")
            return

        except Exception as e:
            logger.error(f"Attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_retries - 1:
                wait_time = (2 ** attempt) * 0.5  # Exponential backoff
                console.print(f"[bold yellow]Retrying in {wait_time:.1f} seconds... (Attempt {attempt + 2} of {max_retries})[/bold yellow]")
                time.sleep(wait_time)
            else:
                console.print(f"[bold red]Failed to create event after {max_retries} attempts: {str(e)}[/bold red]")
                return


async def create_events(calendar_id: str, events_data: Any, tz_str: str = 'America/New_York') -> None:
    if isinstance(events_data, str):
        try:
            events_data = json.loads(events_data)
            logger.debug("Deserialized 'events_data' from string to list.")
        except json.JSONDecodeError:
            console.print(
                "[bold red]Error: 'events' should be a valid JSON string representing a list of events.[/bold red]")
            logger.error(f"Invalid JSON for 'events': {events_data}")
            return

    if not isinstance(events_data, list):
        console.print("[bold red]Error: 'events' should be a list of event objects.[/bold red]")
        logger.error(f"'events' is not a list: {events_data}")
        return

    console.print(f"\n[bold yellow]Preview of Events to be Created:[/bold yellow]")
    for idx, event in enumerate(events_data, start=1):
        console.print(f"\n[bold yellow]Event {idx}:[/bold yellow]")
        format_event_preview(event)

    confirmation = Prompt.ask("Do you want to proceed with creating ALL these events? (yes/no)",
                              choices=["yes", "no"], default="no")

    if confirmation.lower() == "yes":
        for event in events_data:
            await create_event(calendar_id, event, tz_str=tz_str)
            await asyncio.sleep(1)

        console.print(f"[bold green]Created {len(events_data)} event(s) successfully.[/bold green]")
        logger.info(f"Created {len(events_data)} event(s).")
    else:
        console.print("[bold red]Event creation canceled by the user.[/bold red]")
        logger.info("User canceled event creation.")

async def create_calendar(calendar_name: str, color: str = None) -> None:
    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    payload = {
        'name': calendar_name,
        'color': color if color else 'auto'
    }
    
    url = 'https://graph.microsoft.com/v1.0/me/calendars'
    
    response = await make_api_call('POST', url, headers, json=payload)
    
    if 'error' in response:
        console.print(f"[bold red]Failed to create calendar: {response['error']}[/bold red]")
        logger.error(f"Failed to create calendar '{calendar_name}': {response['error']}")
    else:
        console.print(f"[bold green]Calendar '{calendar_name}' created successfully.[/bold green]")
        logger.info(f"Calendar '{calendar_name}' created successfully.")



async def update_event(calendar_id: str, event_id: str, updates: Dict[str, Any], apply_to_series: bool,
                      tz_str: str = 'America/New_York') -> None:
    if not event_id or not isinstance(event_id, str):
        console.print("[bold red]Error: 'event_id' is required and must be a valid string.[/bold red]")
        logger.error("Attempted to update event with invalid 'event_id'.")
        return

    try:
        is_all_day = updates.get('isAllDay')
        
        if 'start' in updates:
            start_datetime = parse_natural_language_date(updates['start'], tz_str=tz_str)
            updates['start'] = {
                'dateTime': start_datetime.isoformat() if not is_all_day else start_datetime.date().isoformat(),
                'timeZone': tz_str
            }

        if 'end' in updates:
            end_datetime = parse_natural_language_date(updates['end'], tz_str=tz_str)
            updates['end'] = {
                'dateTime': end_datetime.isoformat() if not is_all_day else (end_datetime.date() + timedelta(days=1)).isoformat(),
                'timeZone': tz_str
            }

        access_token = get_access_token()
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events/{event_id}'
        
        if apply_to_series:
            url += '/instances'
        
        response = await make_api_call('PATCH', url, headers, json=updates)

        if 'error' in response:
            raise Exception(f"Failed to update event: {response['error']}")

        console.print(f"[bold green]Event '{event_id}' updated successfully.[/bold green]")
        logger.info(f"Event '{event_id}' updated successfully.")

    except Exception as e:
        console.print(f"[bold red]Error updating event: {str(e)}[/bold red]")
        logger.error(f"Error updating event: {str(e)}")

async def delete_calendar(calendar_name: str) -> None:
    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    # Get the calendar ID from the cache or fetch it first
    calendar_id = id_cache.get(calendar_name.lower())
    
    if not calendar_id:
        console.print(f"[bold red]Calendar '{calendar_name}' not found.[/bold red]")
        logger.error(f"Calendar '{calendar_name}' not found in cache.")
        return

    url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}'
    
    response = await make_api_call('DELETE', url, headers)
    
    if 'error' in response:
        console.print(f"[bold red]Failed to delete calendar: {response['error']}[/bold red]")
        logger.error(f"Failed to delete calendar '{calendar_name}': {response['error']}")
    else:
        console.print(f"[bold green]Calendar '{calendar_name}' deleted successfully.[/bold green]")
        logger.info(f"Calendar '{calendar_name}' deleted successfully.")
        
async def create_calendar(calendar_name: str, color: str = None) -> None:
    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    payload = {
        'name': calendar_name,
        'color': color if color else 'auto'
    }
    
    url = 'https://graph.microsoft.com/v1.0/me/calendars'
    
    response = await make_api_call('POST', url, headers, json=payload)
    
    if 'error' in response:
        console.print(f"[bold red]Failed to create calendar: {response['error']}[/bold red]")
        logger.error(f"Failed to create calendar '{calendar_name}': {response['error']}")
    else:
        console.print(f"[bold green]Calendar '{calendar_name}' created successfully.[/bold green]")
        logger.info(f"Calendar '{calendar_name}' created successfully.")
        
    # Update the cache with the new calendar
    await cache_calendars()


async def delete_event(calendar_id: str, event_id: str, cancel_occurrence: bool, cancel_series: bool) -> None:
    if not event_id or not isinstance(event_id, str):
        console.print("[bold red]Error: 'event_id' is required and must be a valid string.[/bold red]")
        logger.error("Attempted to delete event with invalid 'event_id'.")
        return

    try:
        access_token = get_access_token()
        headers = {'Authorization': f'Bearer {access_token}'}
        url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events/{event_id}'
        
        if cancel_occurrence:
            url += '/instances'
        elif cancel_series:
            url += '/series'
        
        response = await make_api_call('DELETE', url, headers)

        if 'error' in response:
            raise Exception(f"Failed to delete event: {response['error']}")

        console.print(f"[bold green]Event '{event_id}' deleted successfully.[/bold green]")
        logger.info(f"Event '{event_id}' deleted successfully.")

    except Exception as e:
        console.print(f"[bold red]Error deleting event: {str(e)}[/bold red]")
        logger.error(f"Error deleting event: {str(e)}")


async def list_events(calendar_id: str, filters: Dict[str, Any], tz_str: str = 'America/New_York') -> None:
    access_token = get_access_token()
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events'
    query_params = []
    if 'start_datetime' in filters and filters['start_datetime'] and 'end_datetime' in filters and filters[
        'end_datetime']:
        try:
            start_datetime = parse_natural_language_date(filters['start_datetime'], tz_str=tz_str).isoformat()
            end_datetime = parse_natural_language_date(filters['end_datetime'], tz_str=tz_str).isoformat()
            query_params.append(f"start/dateTime ge '{start_datetime}'")
            query_params.append(f"end/dateTime le '{end_datetime}'")
            logger.debug(f"Filter start_datetime: {start_datetime}")
            logger.debug(f"Filter end_datetime: {end_datetime}")
        except ValueError as ve:
            console.print(f"[bold red]{ve}[/bold red]")
            logger.error(f"Date parsing error in filters: {ve}")
            return
    if 'categories' in filters and filters['categories']:
        categories = ','.join([f"'{cat}'" for cat in filters['categories']])
        query_params.append(f"categories/any(c:c in ({categories}))")
    if 'subject_contains' in filters and filters['subject_contains']:
        subject = filters['subject_contains']
        query_params.append(f"contains(subject,'{subject}')")
    if query_params:
        filter_query = ' and '.join(query_params)
        url += f"?$filter={filter_query}"
        logger.debug(f"Filter query: {filter_query}")
    response = await make_api_call('GET', url, headers)
    if 'error' in response:
        console.print(f"[bold red]Failed to list events: {response['error']}[/bold red]")
        logger.error(f"Failed to list events: {response['error']}")
    else:
        events = response.get('value', [])
        if not events:
            console.print("[bold yellow]No events found.[/bold yellow]")
            logger.info("No events found with the given filters.")
            return
        table = Table(title="Events")
        table.add_column("Subject")
        table.add_column("Start")
        table.add_column("End")
        for event in events:
            table.add_row(
                event.get('subject', ''),
                format_datetime_us(event.get('start', {}).get('dateTime', '')),
                format_datetime_us(event.get('end', {}).get('dateTime', ''))
            )
        console.print(table)
        logger.info(f"Listed {len(events)} event(s).")


async def get_event(calendar_id: str, event_id: str, properties: List[str], tz_str: str = 'America/New_York') -> None:
    access_token = get_access_token()
    headers = {'Authorization': f'Bearer {access_token}'}
    select_params = ','.join(properties) if properties else ''
    url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}/events/{event_id}'
    if select_params:
        url += f"?$select={select_params}"
    response = await make_api_call('GET', url, headers)
    if 'error' in response:
        console.print(f"[bold red]Failed to get event: {response['error']}[/bold red]")
        logger.error(f"Failed to get event {event_id}: {response['error']}")
    else:
        event = response
        table = Table(title="Event Details")
        table.add_column("Property", style="bold")
        table.add_column("Value")
        for prop in properties:
            value = event.get(prop, '')
            if isinstance(value, dict):
                if prop in ['start', 'end']:
                    if event.get('isAllDay', False):
                        date = value.get('date', '')
                        value = f"{date} (All day)"
                    else:
                        date_time = value.get('dateTime', '')
                        time_zone = value.get('timeZone', 'UTC')
                        try:
                            formatted_dt = format_datetime_us(date_time)
                            value = f"{formatted_dt} ({time_zone})"
                        except Exception as e:
                            value = f"{date_time} ({time_zone})"
                            logger.error(f"Error formatting datetime for property '{prop}': {e}")
                else:
                    value = json.dumps(value, indent=2)
            elif isinstance(value, list):
                if prop == 'attendees':
                    attendees_formatted = ', '.join([att.get('emailAddress', {}).get('address', '') for att in value])
                    value = attendees_formatted if attendees_formatted else 'None'
                else:
                    value = ', '.join([str(item) for item in value]) if value else 'None'
            elif prop == 'isAllDay':
                value = 'Yes' if value else 'No'
            elif not value:
                value = 'None'
            table.add_row(prop.capitalize(), str(value))
        console.print(table)
        logger.info(f"Retrieved details for event '{event_id}'.")


async def list_calendars(filter_query: str, order_by: str) -> None:
    access_token = get_access_token()
    headers = {'Authorization': f'Bearer {access_token}'}
    url = 'https://graph.microsoft.com/v1.0/me/calendars?$select=id,name'
    query_params = []
    if filter_query:
        query_params.append(f"$filter={filter_query}")
    if order_by:
        query_params.append(f"$orderby={order_by}")
    if query_params:
        url += '&' + '&'.join(query_params)
        logger.debug(f"Calendars filter/order query: {'&'.join(query_params)}")

    response = await make_api_call('GET', url, headers)

    if 'error' in response:
        console.print(f"[bold red]Failed to list calendars: {response['error']}[/bold red]")
        logger.error(f"Failed to list calendars: {response['error']}")
    elif 'value' in response:
        calendars = response['value']
        if not calendars:
            console.print("[bold yellow]No calendars found.[/bold yellow]")
            logger.info("No calendars found.")
            return
        table = Table(title="Calendars")
        table.add_column("Name")
        for calendar in calendars:
            table.add_row(calendar.get('name', ''))
        console.print(table)
        logger.info(f"Listed {len(calendars)} calendar(s).")
    else:
        console.print(f"[bold red]Unexpected response format: {response}[/bold red]")
        logger.error(f"Unexpected response format when listing calendars: {response}")


async def get_calendar(calendar_id: str, properties: List[str]) -> None:
    access_token = get_access_token()
    headers = {'Authorization': f'Bearer {access_token}'}
    select_params = ','.join(properties) if properties else ''
    url = f'https://graph.microsoft.com/v1.0/me/calendars/{calendar_id}'
    if select_params:
        url += f"?$select={select_params}"
    response = await make_api_call('GET', url, headers)
    if 'error' in response:
        console.print(f"[bold red]Failed to get calendar: {response['error']}[/bold red]")
        logger.error(f"Failed to get calendar {calendar_id}: {response['error']}")
    else:
        calendar = response
        table = Table(title="Calendar Details")
        table.add_column("Property", style="bold")
        table.add_column("Value")
        for prop in properties:
            value = calendar.get(prop, '')
            if isinstance(value, dict):
                value = json.dumps(value, indent=2)
            elif isinstance(value, list):
                value = ', '.join([str(item) for item in value]) if value else 'None'
            elif not value:
                value = 'None'
            table.add_row(prop.capitalize(), str(value))
        console.print(table)
        logger.info(f"Retrieved details for calendar '{calendar_id}'.")


async def search_events(query: str, calendar_names: List[str], start_datetime: str, end_datetime: str,
                         categories: List[str], tz_str: str = 'America/New_York') -> None:
    access_token = get_access_token()
    headers = {'Authorization': f'Bearer {access_token}'}
    search_url = 'https://graph.microsoft.com/v1.0/me/events'
    filters = []
    if query:
        filters.append(f"contains(subject,'{query}') or contains(body/content,'{query}')")
    if calendar_names:
        calendar_ids = [id_cache.get(name.lower()) for name in calendar_names if id_cache.get(name.lower())]
        calendar_ids = [cid for cid in calendar_ids if cid]
        if calendar_ids:
            calendar_id_list = ','.join([f"'{cid}'" for cid in calendar_ids])
            filters.append(f"calendar/id in ({calendar_id_list})")
    if start_datetime and end_datetime:
        try:
            start_dt = parse_natural_language_date(start_datetime, tz_str=tz_str).isoformat()
            end_dt = parse_natural_language_date(end_datetime, tz_str=tz_str).isoformat()
            filters.append(f"start/dateTime ge '{start_dt}' and end/dateTime le '{end_dt}'")
            logger.debug(f"Search filter start_datetime: {start_dt}")
            logger.debug(f"Search filter end_datetime: {end_dt}")
        except ValueError as ve:
            console.print(f"[bold red]{ve}[/bold red]")
            logger.error(f"Date parsing error in search filters: {ve}")
            return
    if categories:
        categories_quoted = ','.join([f"'{cat}'" for cat in categories])
        filters.append(f"categories/any(c:c in ({categories_quoted}))")
    filter_query = ' and '.join(filters)
    if filter_query:
        search_url += f"?$filter={filter_query}"
        logger.debug(f"Search filter query: {filter_query}")
    response = await make_api_call('GET', search_url, headers)
    if 'error' in response:
        console.print(f"[bold red]Failed to search events: {response['error']}[/bold red]")
        logger.error(f"Failed to search events: {response['error']}")
    else:
        events = response.get('value', [])
        if not events:
            console.print("[bold yellow]No events matched the search criteria.[/bold yellow]")
            logger.info("No events matched the search criteria.")
            return
        table = Table(title="Search Results")
        table.add_column("Subject")
        table.add_column("Start")
        table.add_column("End")
        for event in events:
            table.add_row(
                event.get('subject', ''),
                format_datetime_us(event.get('start', {}).get('dateTime', '')),
                format_datetime_us(event.get('end', {}).get('dateTime', ''))
            )
        console.print(table)
        logger.info(f"Search returned {len(events)} event(s).")


async def manage_categories(action: str, category_name: str, new_category_name: str = None, color: str = None) -> None:
    action_completer = WordCompleter(['create', 'update', 'delete'])
    color_completer = WordCompleter(['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightPink', 'lightRed', 'lightYellow'])
    
    session = PromptSession()
    if not action:
        action = await session.prompt_async("Choose an action (create/update/delete): ", completer=action_completer)
    if not category_name:
        category_name = await session.prompt_async("Enter category name: ")
    if action == 'update' and not new_category_name:
        new_category_name = await session.prompt_async("Enter new category name: ")
    if (action == 'create' or action == 'update') and not color:
        color = await session.prompt_async("Choose a color: ", completer=color_completer)

    access_token = get_access_token()
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    url = 'https://graph.microsoft.com/v1.0/me/outlook/masterCategories'
    if action == "create":
        payload = {
            "name": category_name,
            "color": color if color else "preset0"
        }
        response = await make_api_call('POST', url, headers, json=payload)
        if 'error' in response:
            console.print(f"[bold red]Failed to create category: {response['error']}[/bold red]")
            logger.error(f"Failed to create category '{category_name}': {response['error']}")
        else:
            console.print(f"[bold green]Category '{category_name}' created successfully.[/bold green]")
            logger.info(f"Category '{category_name}' created successfully.")

    elif action == "update":
        response = await make_api_call('GET', url, headers)
        if 'error' in response:
            console.print(f"[bold red]Failed to retrieve categories: {response['error']}[/bold red]")
            logger.error(f"Failed to retrieve categories for update: {response['error']}")
            return
        categories = response.get('value', [])
        category = next((cat for cat in categories if cat.get('name', '').lower() == category_name.lower()), None)
        if not category:
            console.print(f"[bold red]Category '{category_name}' not found.[/bold red]")
            logger.error(f"Category '{category_name}' not found for update.")
            return
        category_id = category.get('id')
        update_url = f"{url}/{category_id}"
        payload = {}
        if new_category_name:
            payload["name"] = new_category_name
        if color:
            payload["color"] = color
        if not payload:
            console.print("[bold yellow]No updates provided.[/bold yellow]")
            logger.warning("No updates provided for category management.")
            return
        response = await make_api_call('PATCH', update_url, headers, json=payload)
        if 'error' in response:
            console.print(f"[bold red]Failed to update category: {response['error']}[/bold red]")
            logger.error(f"Failed to update category '{category_name}': {response['error']}")
        else:
            console.print(f"[bold green]Category '{category_name}' updated successfully.[/bold green]")
            logger.info(f"Category '{category_name}' updated successfully.")

    elif action == "delete":
        response = await make_api_call('GET', url, headers)
        if 'error' in response:
            console.print(f"[bold red]Failed to retrieve categories: {response['error']}[/bold red]")
            logger.error(f"Failed to retrieve categories for deletion: {response['error']}")
            return
        categories = response.get('value', [])
        category = next((cat for cat in categories if cat.get('name', '').lower() == category_name.lower()), None)
        if not category:
            console.print(f"[bold red]Category '{category_name}' not found.[/bold red]")
            logger.error(f"Category '{category_name}' not found for deletion.")
            return
        category_id = category.get('id')
        delete_url = f"{url}/{category_id}"
        response = await make_api_call('DELETE', delete_url, headers)
        if 'error' in response:
            console.print(f"[bold red]Failed to delete category: {response['error']}[/bold red]")
            logger.error(f"Failed to delete category '{category_name}': {response['error']}")
        else:
            console.print(f"[bold green]Category '{category_name}' deleted successfully.[/bold green]")
            logger.info(f"Category '{category_name}' deleted successfully.")
    else:
        console.print("[bold red]Invalid action. Please choose from create, update, or delete.[/bold red]")
        logger.warning(f"Invalid category management action: {action}")


def get_function_schemas() -> List[Dict[str, Any]]:
    return [
        {
        "name": "create_calendar",
        "description": "Create a new calendar with the specified name and optional color.",
        "parameters": {
            "type": "object",
            "properties": {
            "calendar_name": {
                "type": "string",
                "description": "The name of the calendar to create."
            },
            "color": {
                "type": "string",
                "description": "The optional color for the calendar. Defaults to 'auto'.",
                "enum": ["auto", "lightBlue", "lightGreen", "lightOrange", "lightGray", "lightPink", "lightRed", "lightYellow"]
            }
            },
            "required": ["calendar_name"]
        }
        },
        {
            "name": "delete_calendar",
            "description": "Delete a calendar by its name.",
            "parameters": {
                "type": "object",
                "properties": {
                "calendar_name": {
                    "type": "string",
                    "description": "The name of the calendar to delete."
                    }
                },
                "required": ["calendar_name"]
        }
        },
        {
            "name": "create_event",
            "description": "Create a new calendar event.",
            "parameters": {
                "type": "object",
                "properties": {
                    "calendar_name": {"type": "string"},
                    "event": {
                        "type": "object",
                        "properties": {
                            "subject": {"type": "string"},
                            "body": {"type": "string"},
                            "start": {"type": "string", "format": "date-time"},
                            "end": {"type": "string", "format": "date-time"},
                            "location": {"type": "string"},
                            "attendees": {
                                "type": "array",
                                "items": {"type": "string", "format": "email"}
                            },
                            "recurrence": {"type": "string"},
                            "categories": {
                                "type": "array",
                                "items": {"type": "string"}
                            },
                            "reminderMinutesBeforeStart": {"type": "integer"},
                            "isReminderOn": {"type": "boolean"},
                            "isAllDay": {"type": "boolean"}
                        },
                        "required": ["subject", "start", "end"]
                    }
                },
                "required": ["calendar_name", "event"]
            }
        },
        {
            "name": "create_events",
            "description": "Create multiple calendar events at once.",
            "parameters": {
                "type": "object",
                "properties": {
                    "calendar_name": {"type": "string"},
                    "events": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "subject": {"type": "string"},
                                "body": {"type": "string"},
                                "start": {"type": "string", "format": "date-time"},
                                "end": {"type": "string", "format": "date-time"},
                                "location": {"type": "string"},
                                "attendees": {
                                    "type": "array",
                                    "items": {"type": "string", "format": "email"}
                                },
                                "recurrence": {"type": "string"},
                                "categories": {
                                    "type": "array",
                                    "items": {"type": "string"}
                                },
                                "reminderMinutesBeforeStart": {"type": "integer"},
                                "isReminderOn": {"type": "boolean"}
                            },
                            "required": ["subject", "start", "end"]
                        }
                    }
                },
                "required": ["calendar_name", "events"]
            }
        },
        {
            "name": "update_event",
            "description": "Update an existing calendar event.",
            "parameters": {
                "type": "object",
                "properties": {
                    "calendar_name": {"type": "string"},
                    "event_id": {"type": "string"},
                    "updates": {
                        "type": "object",
                        "properties": {
                            "subject": {"type": "string"},
                            "body": {"type": "string"},
                            "start": {"type": "string", "format": "date-time"},
                            "end": {"type": "string", "format": "date-time"},
                            "location": {"type": "string"},
                            "attendees": {
                                "type": "array",
                                "items": {"type": "string", "format": "email"}
                            },
                            "recurrence": {"type": "string"},
                            "categories": {
                                "type": "array",
                                "items": {"type": "string"}
                            },
                            "reminderMinutesBeforeStart": {"type": "integer"},
                            "isReminderOn": {"type": "boolean"},
                            "isAllDay": {"type": "boolean"}
                        }
                    },
                    "apply_to_series": {"type": "boolean"}
                },
                "required": ["calendar_name", "event_id", "updates"]
            }
        },
        {
            "name": "delete_event",
            "description": "Delete a calendar event.",
            "parameters": {
                "type": "object",
                "properties": {
                    "calendar_name": {"type": "string"},
                    "event_id": {"type": "string"},
                    "cancel_occurrence": {"type": "boolean"},
                    "cancel_series": {"type": "boolean"}
                },
                "required": ["calendar_name", "event_id"]
            }
        },
        {
            "name": "list_events",
            "description": "List calendar events with optional filters.",
            "parameters": {
                "type": "object",
                "properties": {
                    "calendar_name": {"type": "string"},
                    "start_datetime": {"type": "string", "format": "date-time"},
                    "end_datetime": {"type": "string", "format": "date-time"},
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"}
                    },
                    "subject_contains": {"type": "string"}
                },
                "required": ["calendar_name"]
            }
        },
        {
            "name": "get_event",
            "description": "Get detailed information about a specific event.",
            "parameters": {
                "type": "object",
                "properties": {
                    "calendar_name": {"type": "string"},
                    "event_id": {"type": "string"},
                    "properties": {
                        "type": "array",
                        "items": {"type": "string"}
                    }
                },
                "required": ["calendar_name", "event_id"]
            }
        },
        {
            "name": "list_calendars",
            "description": "List all available calendars with optional filtering and ordering.",
            "parameters": {
                "type": "object",
                "properties": {
                    "filter": {"type": "string"},
                    "order_by": {"type": "string"}
                },
                "required": []
            }
        },
        {
            "name": "get_calendar",
            "description": "Get detailed information about a specific calendar.",
            "parameters": {
                "type": "object",
                "properties": {
                    "calendar_name": {"type": "string"},
                    "properties": {
                        "type": "array",
                        "items": {"type": "string"}
                    }
                },
                "required": ["calendar_name"]
            }
        },
        {
            "name": "search_events",
            "description": "Search for events across calendars with various filters.",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string"},
                    "calendar_names": {
                        "type": "array",
                        "items": {"type": "string"}
                    },
                    "start_datetime": {"type": "string", "format": "date-time"},
                    "end_datetime": {"type": "string", "format": "date-time"},
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"}
                    }
                },
                "required": []
            }
        },
        {
            "name": "manage_categories",
            "description": "Manage event categories (create, update, delete).",
            "parameters": {
                "type": "object",
                "properties": {
                    "action": {"type": "string"},
                    "category_name": {"type": "string"},
                    "new_category_name": {"type": "string"},
                    "color": {"type": "string"}
                },
                "required": ["action", "category_name"]
            }
        }
    ]


SYSTEM_MESSAGE = """
## Microsoft 365 Calendar Assistant Instructions

You are an AI assistant designed to manage Microsoft 365 calendars. Your primary goal is to assist the user by executing tasks precisely, improving their input where necessary, and ensuring that all calendar-related tasks are handled smoothly.

### Core Capabilities:

* **Create Events:** Set up single or recurring events, specifying details such as date, time, title, description, location, and reminders.
* **Update Events:** Modify existing events, including changes to date, time, title, description, location, and reminders.
* **Delete Events:** Remove events from the calendar as requested.
* **List Events:** Display events using various filters to help users find what they need.
* **Get Event Details:** Provide comprehensive information about specific events upon request.
* **Manage Calendars:** List, organize, add, or remove calendars.
* **Search Events:** Find events across all calendars based on specified criteria.
* **Manage Categories:** Organize and manage event categories for better classification and filtering.

### Current Date and Time:

You are aware of the current date and time, which will be provided to you in each user interaction. Use this information to provide context-aware responses and to handle relative time expressions accurately.

### Input Improvement & Clarification:

When users provide input that is unclear, incomplete, or scattered, your role is to:

* **Revise and Enhance:** Reorganize and refine user input to ensure clarity and accuracy, preserving the original intent.
* **Clarify Ambiguities:** Ask for further clarification when user input is incomplete or unclear, but provide helpful suggestions to guide the user.
* **Propose Improvements:** Where applicable, suggest ways to optimize or improve user commands to make them more concise or actionable.

For instance, if a user input is vague or lacks important information (e.g., missing date or time for an event), ask follow-up questions to gather the necessary details. If the request includes a mix of instructions (e.g., creating and listing events in a single request), separate the tasks and clarify each part before proceeding.

### Handling User Requests:

When processing user instructions, follow these structured guidelines:

#### Drafting Events:

For each event request, compile a clear and concise draft with the following:

* **Date:** Specify the date of the event.
* **Start Time-End Time or All Day:** Indicate whether the event is all-day or has a defined start and end time.
* **Event Title:** Ensure each word in the title is capitalized.
* **Description:** Provide a brief summary (one to three sentences) of the event.
* **Location:** Include where the event will take place.
* **Reminder:** Set a reminder with options such as 5 minutes, 15 minutes, 30 minutes, 1 hour, 2 hours, 1 day, or 1 week before the event.
* **Repeat:** Clarify whether the event repeats daily, weekly, monthly, or yearly.

#### Confirm Major Actions:

Always confirm major actions, such as creating, updating, or deleting events, with the user before proceeding.

#### Handle Multiple Instructions:

When a user provides multiple instructions in one message, break them down and address each request separately, ensuring all components are understood and executed appropriately.

### Communication Style:

* Provide clear and concise responses while ensuring that actions are carried out according to the user’s instructions.
* When necessary, offer improvements or suggestions to enhance the user’s input or make processes more efficient. 

"""


def main():
    loop = asyncio.get_event_loop()
    try:
        loop.run_until_complete(async_main())
    finally:
        loop.close()

async def async_main():
    console.print("[bold blue]Welcome to the Microsoft 365 Calendar Assistant![/bold blue]")
    console.print("[bold yellow]Type your commands and press Enter. For multiline input, use Shift+Enter.[/bold yellow]")
    console.print("[bold yellow]Authenticating with Microsoft 365...[/bold yellow]")
    try:
        access_token = get_access_token()
        logger.info("Authentication successful.")
        console.print("[bold green]Authentication successful.[/bold green]")
    except Exception as e:
        logger.exception("Authentication failed.")
        console.print(f"[bold red]Authentication failed: {e}[/bold red]")
        return

    console.print("[bold yellow]Fetching your calendars...[/bold yellow]")
    calendars = await cache_calendars()
    if not calendars:
        console.print("[bold red]Could not retrieve calendars.[/bold red]")
        logger.error("Could not retrieve calendars.")
        return
    console.print("[bold green]Calendars cached.[/bold green]")
    logger.info("Calendars cached.")

    default_calendar_id = get_default_calendar_id()
    if default_calendar_id:
        default_calendar_name = next((cal['name'] for cal in calendars if cal['id'] == default_calendar_id), 'Unknown')
        console.print(f"[bold green]Default calendar set: {default_calendar_name}[/bold green]")
    else:
        console.print("[bold yellow]No default calendar found.[/bold yellow]")

    functions = get_function_schemas()

    session = PromptSession(
        history=get_history(),
        key_bindings=get_key_bindings(),
    )

    while True:
        try:
            current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            user_input = await session.prompt_async(
                get_custom_prompt(current_datetime),
                multiline=False,
            )
        except EOFError:
            console.print("[bold blue]No input received. Exiting...[/bold blue]")
            logger.info("No input received. Exiting application.")
            break

        if user_input.lower() in ('exit', 'quit', '/exit'):
            console.print("[bold blue]Goodbye![/bold blue]")
            logger.info("User exited the application.")
            return

        conversation_history.append({
            'role': 'user',
            'content': user_input,
            'timestamp': datetime.utcnow().isoformat()
        })

        trimmed_history = trim_conversation_history(conversation_history)

        console.print("[bold magenta]Assistant is thinking...[/bold magenta]")

        try:
            current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            system_message_with_time = f"{SYSTEM_MESSAGE}\n\nCurrent date and time: {current_datetime}"
            response = await asyncio.to_thread(
                client.chat.completions.create,
                model="gpt-4o-mini",
                messages=[{"role": "system", "content": system_message_with_time}] + trimmed_history,
                functions=functions,
                function_call="auto",
                temperature=0.2,
            )
            logger.debug(f"OpenAI API response: {response}")
        except OpenAIError as e:
            if isinstance(e, RateLimitError):
                console.print("[bold red]Rate limit exceeded. Please wait a moment before trying again.[/bold red]")
            elif isinstance(e, APIConnectionError):
                console.print("[bold red]Connection to OpenAI failed. Please check your internet connection.[/bold red]")
            else:
                console.print(f"[bold red]An error occurred while communicating with OpenAI: {str(e)}[/bold red]")
            logger.error(f"OpenAI API error: {str(e)}")
            continue

        if not response or not response.choices:
            console.print("[bold red]No response from OpenAI API.[/bold red]")
            logger.error("No response received from OpenAI API.")
            continue

        assistant_message = response.choices[0].message
        logger.debug(f"Assistant message: {assistant_message}")

        if not assistant_message:
            console.print("[bold red]Received empty response from AI assistant. This may be due to a temporary issue. Please try your request again.[/bold red]")
            logger.error("Empty response received from OpenAI API")
            continue

        if assistant_message.function_call:
            function_call = assistant_message.function_call
            function_name = function_call.name
            logger.debug(f"Function call detected: {function_name}")

            try:
                function_args = json.loads(function_call.arguments)
                logger.debug(f"Function arguments after json.loads: {function_args}")
            except json.JSONDecodeError:
                logger.error(f"Failed to parse function arguments: {function_call.arguments}")
                console.print("[bold red]Error: Failed to parse function arguments.[/bold red]")
                conversation_history.append({
                    'role': 'assistant',
                    'content': "I'm sorry, I couldn't parse the function arguments.",
                    'timestamp': datetime.utcnow().isoformat()
                })
                continue

            for key, value in function_args.items():
                if isinstance(value, str) and key == 'events':
                    try:
                        function_args[key] = json.loads(value)
                        logger.debug(f"Deserialized 'events' argument: {function_args[key]}")
                    except json.JSONDecodeError:
                        logger.error(f"Failed to parse 'events' argument: {value}")
                        console.print("[bold red]Error: 'events' should be a valid JSON array.[/bold red]")
                        function_args[key] = []

                if key == 'events' and isinstance(function_args[key], list):
                    for idx, event in enumerate(function_args[key]):
                        if isinstance(event, str):
                            try:
                                function_args[key][idx] = json.loads(event)
                                logger.debug(f"Deserialized individual event {idx + 1}: {function_args[key][idx]}")
                            except json.JSONDecodeError:
                                logger.error(f"Failed to parse individual event {idx + 1}: {event}")
                                console.print(f"[bold red]Error: Event {idx + 1} is not a valid JSON object.[/bold red]")
                                function_args[key][idx] = {}

            logger.debug(f"Function call arguments: {function_args}")

            if not isinstance(function_args, dict):
                logger.error(f"Function arguments are not a dictionary: {function_args}")
                console.print("[bold red]Error: Function arguments are not in the expected format.[/bold red]")
                conversation_history.append({
                    'role': 'assistant',
                    'content': "I'm sorry, the function arguments are not in the expected format.",
                    'timestamp': datetime.utcnow().isoformat()
                })
                continue

            try:
                if function_name == "create_event":
                    event_data = function_args.get('event', {})
                    await create_event(event_data)
                    conversation_history.append({
                        'role': 'assistant',
                        'content': f"Event '{event_data.get('subject', 'Untitled')}' has been created.",
                        'timestamp': datetime.utcnow().isoformat()
                    })

                elif function_name == "create_events":
                    events_data = function_args.get('events', [])
                    logger.debug(f"Type of 'events_data': {type(events_data)}")
                    logger.debug(f"Content of 'events_data': {events_data}")
                    await create_events(events_data)

                elif function_name == "update_event":
                    calendar_name = function_args.get('calendar_name')
                    calendar_name = calendar_name.lower() if calendar_name else None
                    event_id = function_args.get('event_id')
                    updates = function_args.get('updates', {})
                    apply_to_series = function_args.get('apply_to_series', False)
                    calendar_id = id_cache.get(calendar_name, get_default_calendar_id()) if calendar_name else get_default_calendar_id()
                    if not calendar_id:
                        console.print(f"[bold red]Calendar '{calendar_name or 'Default'}' not found and no default calendar available.[/bold red]")
                        logger.error(f"Calendar '{calendar_name or 'Default'}' not found and no default calendar available.")
                        continue
                    await update_event(calendar_id, event_id, updates, apply_to_series)

                elif function_name == "delete_event":
                    calendar_name = function_args.get('calendar_name')
                    calendar_name = calendar_name.lower() if calendar_name else None
                    event_id = function_args.get('event_id')
                    cancel_occurrence = function_args.get('cancel_occurrence', False)
                    cancel_series = function_args.get('cancel_series', False)
                    calendar_id = id_cache.get(calendar_name, get_default_calendar_id()) if calendar_name else get_default_calendar_id()
                    if not calendar_id:
                        console.print(f"[bold red]Calendar '{calendar_name or 'Default'}' not found and no default calendar available.[/bold red]")
                        logger.error(f"Calendar '{calendar_name or 'Default'}' not found and no default calendar available.")
                        continue
                    await delete_event(calendar_id, event_id, cancel_occurrence, cancel_series)

                elif function_name == "list_events":
                    calendar_name = function_args.get('calendar_name')
                    calendar_name = calendar_name.lower() if calendar_name else None
                    filters = {
                        'start_datetime': function_args.get('start_datetime'),
                        'end_datetime': function_args.get('end_datetime'),
                        'categories': function_args.get('categories'),
                        'subject_contains': function_args.get('subject_contains')
                    }
                    calendar_id = id_cache.get(calendar_name, get_default_calendar_id()) if calendar_name else get_default_calendar_id()
                    if not calendar_id:
                        console.print(f"[bold red]Calendar '{calendar_name or 'Default'}' not found and no default calendar available.[/bold red]")
                        logger.error(f"Calendar '{calendar_name or 'Default'}' not found and no default calendar available.")
                        continue
                    await list_events(calendar_id, filters)

                elif function_name == "get_event":
                    calendar_name = function_args.get('calendar_name')
                    calendar_name = calendar_name.lower() if calendar_name else None
                    event_id = function_args.get('event_id')
                    properties = function_args.get('properties', [])
                    calendar_id = id_cache.get(calendar_name, get_default_calendar_id()) if calendar_name else get_default_calendar_id()
                    if not calendar_id:
                        console.print(f"[bold red]Calendar '{calendar_name or 'Default'}' not found and no default calendar available.[/bold red]")
                        logger.error(f"Calendar '{calendar_name or 'Default'}' not found and no default calendar available.")
                        continue
                    await get_event(calendar_id, event_id, properties)

                elif function_name == "list_calendars":
                    filter_query = function_args.get('filter', '')
                    order_by = function_args.get('order_by', '')
                    await list_calendars(filter_query, order_by)

                elif function_name == "get_calendar":
                    calendar_name = function_args.get('calendar_name', '').lower()
                    properties = function_args.get('properties', [])
                    calendar_id = id_cache.get(calendar_name, get_default_calendar_id())
                    if not calendar_id:
                        console.print(f"[bold red]Calendar '{calendar_name}' not found and no default calendar available.[/bold red]")
                        logger.error(f"Calendar '{calendar_name}' not found and no default calendar available.")
                        continue
                    await get_calendar(calendar_id, properties)

                elif function_name == "search_events":
                    query = function_args.get('query', '')
                    calendar_names = function_args.get('calendar_names', [])
                    start_datetime = function_args.get('start_datetime')
                    end_datetime = function_args.get('end_datetime')
                    categories = function_args.get('categories', [])
                    await search_events(query, calendar_names, start_datetime, end_datetime, categories)

                elif function_name == "manage_categories":
                    action = function_args.get('action')
                    category_name = function_args.get('category_name')
                    new_category_name = function_args.get('new_category_name')
                    color = function_args.get('color')
                    await manage_categories(action, category_name, new_category_name, color)

                elif function_name == "create_calendar":
                    calendar_name = function_args.get('calendar_name')
                    color = function_args.get('color')
                    await create_calendar(calendar_name, color)

                else:
                    console.print(f"[bold yellow]Unknown function '{function_name}'.[/bold yellow]")
                    logger.warning(f"Unknown function called: {function_name}")

            except Exception as e:
                logger.exception(f"Error executing function {function_name}")
                console.print(f"[bold red]Error executing function {function_name}: {str(e)}[/bold red]")
                if function_name == "create_event":
                    console.print("[bold yellow]Tip: Make sure you've provided all required event details (subject, start time, end time).[/bold yellow]")
                elif function_name in ["update_event", "delete_event", "get_event"]:
                    console.print("[bold yellow]Tip: Ensure you've specified the correct calendar and event ID.[/bold yellow]")
                continue

            conversation_history.append({
                'role': 'assistant',
                'content': f"Executed function `{function_name}`.",
                'timestamp': datetime.utcnow().isoformat()
            })
        else:
            if assistant_message.content:
                console.print("[bold cyan]Assistant:[/bold cyan] " + assistant_message.content)
                conversation_history.append({
                    'role': 'assistant',
                    'content': assistant_message.content,
                    'timestamp': datetime.utcnow().isoformat()
                })
                logger.info("Assistant provided a text response.")
            else:
                console.print("[bold red]Error: Assistant response is empty.[/bold red]")
                logger.error("Assistant response is empty.")


def test_date_parsing():
    test_inputs = ["today", "tomorrow", "next Monday", "three weeks from today"]
    tz_str = 'America/New_York'

    for input_str in test_inputs:
        try:
            parsed = parse_natural_language_date(input_str, tz_str=tz_str)
            logger.debug(f"Test Input: '{input_str}' => Parsed Date: {parsed}")
            console.print(f"[bold green]Input: '{input_str}' => Parsed Date: {parsed}[/bold green]")
        except ValueError as ve:
            console.print(f"[bold red]Input: '{input_str}' => Error: {ve}[/bold red]")


if __name__ == '__main__':
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    try:
        main()
    except KeyboardInterrupt:
        console.print("[bold blue]Interrupted by user. Exiting...[/bold blue]")
        logger.info("Application interrupted by user.")
    except Exception as e:
        logger.exception("Unhandled exception in main")
        console.print(f"[bold red]An error occurred: {e}[/bold red]")

# Removed calendar and action completers
