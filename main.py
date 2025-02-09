import datetime
import logging
import os
import random
import re

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from pyairtable import Api

AIRTABLE_API_KEY = os.environ.get("AIRTABLE_API_KEY")
AIRTABLE_BASE_ID = os.environ.get("AIRTABLE_BASE_ID")
SESSIONS_TABLE = os.environ.get("SESSIONS_TABLE", "Sessions")
CURATION_STATUS_TABLE = os.environ.get("CURATION_STATUS_TABLE", "Statuses")
SPEAKERS_TABLE = os.environ.get("SPEAKERS_TABLE", "Speaker")

GOOGLE_PRESENTATION_ID = os.environ.get("GOOGLE_PRESENTATION_ID")
TEMPLATE_SLIDE_ID = os.environ.get("TEMPLATE_SLIDE_ID")
GOOGLE_SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE")

# Google API scopes
SCOPES = [
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
]

# Set up logging
logging.basicConfig(level=logging.INFO)

# Define a mapping for status-to-color for speakers/moderators.
status_to_color_map = {
    'Idea': 'red',
    'Inbound': 'red',
    'Confirmed, Needs Reg': 'orange',
    'Ask Needed': 'red',
    'Inquired': 'red',
    'To Inquire': 'red',
    'Partner Idea': 'red',
    'LL Keynote Idea': 'red',
    'Keynote Target': 'red',
    'Flag For Next Year': 'red',
    'In Progress': 'yellow',
    'Sent Registration': 'yellow',
    'Sent Registration ': 'yellow',
    'Registered': 'green',
    'Registered, Missing Headshot/Bio': 'orange',
    'Meeting - to set': 'yellow',
    'Meeting - being set ': 'yellow',
    'Meeting - upcoming ': 'yellow',
    'Meeting - complete': 'yellow',
    'Signed, No Content Request': 'red',
    'Signed, No Content Request, but helpful': 'red',
    'Waiting for Content Ideas': 'yellow',
    'Need to Review Proposals': 'yellow',
    'Need Representation in Program': 'yellow',
    'No Further Action Needed': 'green',
    'Program Representation Secured': 'green',
}


# ================================================
# Helper: Google Slides Service
# ================================================
def get_slides_service():
    creds = Credentials.from_service_account_file(GOOGLE_SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('slides', 'v1', credentials=creds)
    return service


# ================================================
# Helper: Airtable – get a record column value
# (Used to resolve the curation status from its table)
# ================================================
def get_airtable_record_column_value(base_id: str, api: Api, table_name: str, record_id: str, column_with_name: str):
    table = api.table(base_id, table_name)
    record = table.get(record_id)
    return record.get("fields", {}).get(column_with_name)


# ================================================
# Helper: Adjust Representation
# ================================================
def adjust_representation(s: str) -> str:
    s = s.strip()
    match = re.match(r'^([^\(]+?)(?:\s*\(\s*(.*?)\s*\))?$', s)
    if match:
        name = match.group(1).strip()
        inside_parens = match.group(2)
        if inside_parens is not None:
            parts = [part.strip() for part in inside_parens.split(',') if part.strip()]
            if parts:
                return f"{name} ({', '.join(parts)})"
            else:
                return name
        else:
            return name
    else:
        return s


# ================================================
# Helper: Add spaces after commas (outside parentheses)
# ================================================
def add_spaces_after_commas_between_people(s: str) -> str:
    result = ''
    paren_depth = 0
    i = 0
    while i < len(s):
        c = s[i]
        if c == '(':
            paren_depth += 1
            result += c
            i += 1
        elif c == ')':
            if paren_depth > 0:
                paren_depth -= 1
            result += c
            i += 1
        elif c == ',' and paren_depth == 0:
            result += ', '
            i += 1
            while i < len(s) and s[i] == ' ':
                i += 1
        else:
            result += c
            i += 1
    return result


# ================================================
# Helper: Split people and indices
# ================================================
def split_people_and_indices(s: str):
    """
    Returns a list of dictionaries with:
      - personString: the substring for that person,
      - startIndex: starting index,
      - endIndex: ending index.
    """
    results = []
    start_index = 0
    paren_depth = 0
    i = 0
    while i < len(s):
        c = s[i]
        if c == '(':
            paren_depth += 1
            i += 1
        elif c == ')':
            if paren_depth > 0:
                paren_depth -= 1
            i += 1
        elif c == ',' and paren_depth == 0:
            if i + 1 < len(s) and s[i + 1] == ' ':
                end_index = i
                person_string = s[start_index:end_index].strip()
                results.append({
                    "personString": person_string,
                    "startIndex": start_index,
                    "endIndex": end_index
                })
                i += 2
                start_index = i
            else:
                i += 1
        else:
            i += 1
    if start_index < len(s):
        person_string = s[start_index:].strip()
        results.append({
            "personString": person_string,
            "startIndex": start_index,
            "endIndex": len(s)
        })
    return results


# ================================================
# Helper: Generate new IDs (for create mode)
# ================================================
def generate_new_ids(initial_id: str):
    parts = initial_id.split('_')
    try:
        last_number = int(parts[2])
    except (IndexError, ValueError):
        last_number = random.randint(1, 1000)
    id_plus_one = f"{parts[0]}_{parts[1]}_{last_number + 1}" if len(parts) >= 2 else initial_id + "_1"
    id_plus_six = f"{parts[0]}_{parts[1]}_{last_number + 6}" if len(parts) >= 2 else initial_id + "_6"
    return {"idPlusOne": id_plus_one, "idPlusSix": id_plus_six}


# ================================================
# Main Function: Update (or create) a slide in Google Slides
# ================================================
def update_presentation_with_slide(service,
                                   presentation_id: str,
                                   slide_object_id: str,
                                   status: str,
                                   date_time: str,
                                   note: str,
                                   title: str,
                                   channel: str,
                                   description: str,
                                   speakers: str,
                                   speaker_note: str,
                                   speaker_colors: str,
                                   moderators: str,
                                   moderator_colors: str,
                                   partners: str = "",
                                   partners_colors: str = "",
                                   slide_id: str = "") -> dict:
    # Map statuses to background colors for text
    COLORS = {
        "green": {"red": 0.6, "green": 1, "blue": 0.6},
        "yellow": {"red": 1, "green": 1, "blue": 0.6},
        "red": {"red": 1, "green": 0.6, "blue": 0.6},
        "orange": {"red": 1, "green": 0.647, "blue": 0}
    }

    # Process colors and people strings
    speaker_colors_list = [c.strip() for c in speaker_colors.split(",")]
    moderator_colors_list = [c.strip() for c in moderator_colors.split(",")]
    speakers = add_spaces_after_commas_between_people(speakers)
    moderators = add_spaces_after_commas_between_people(moderators)

    # Choose background color based on curation status
    if status.strip() == '(5) Confirmed':
        background_color = {"red": 0.83, "green": 0.898, "blue": 0.812}
    else:
        background_color = {"red": 0.804, "green": 0.87, "blue": 0.98}

    requests_list = []
    table_id = None
    speaker_notes_id = None
    main_slide_id = None
    table_element = None
    slide_data = None

    # --- UPDATE mode: if slide_id is provided, fetch slide details ---
    if slide_id:
        main_slide_id = slide_id
        # Retrieve the full presentation to find the slide
        presentation = service.presentations().get(presentationId=presentation_id).execute()
        slide_data = None
        for page in presentation.get("slides", []):
            if page.get("objectId") == slide_id:
                slide_data = page
                break
        if not slide_data:
            raise Exception(f"Slide with id {slide_id} not found.")
        if "pageElements" not in slide_data or not slide_data["pageElements"]:
            raise Exception("No page elements found on the slide.")
        # Find the table element
        table_element = next((pe for pe in slide_data["pageElements"] if "table" in pe), None)
        if not table_element:
            raise Exception("No table element found on the slide.")
        table_id = table_element["objectId"]
        # Locate the speaker notes shape (assume index 1 in notesPage)
        notes_page = slide_data.get("slideProperties", {}).get("notesPage", {})
        notes_elements = notes_page.get("pageElements", [])
        if len(notes_elements) > 1:
            speaker_notes_id = notes_elements[1]["objectId"]
        else:
            logging.warning("Speaker notes element not found; speaker notes updates may fail.")
    else:
        # --- CREATE mode: duplicate the template slide ---
        num = (str(random.random() + 1))[7:]
        main_slide_id = f"{slide_object_id}_{num}"
        base_ids = generate_new_ids(slide_object_id)
        table_id = f"{base_ids['idPlusOne']}_{num}"
        speaker_notes_id = f"{base_ids['idPlusSix']}_{num}"
        duplicate_request = {
            "duplicateObject": {
                "objectId": slide_object_id,
                "objectIds": {
                    slide_object_id: main_slide_id,
                    base_ids['idPlusOne']: table_id,
                    base_ids['idPlusSix']: speaker_notes_id
                }
            }
        }
        requests_list.append(duplicate_request)

    # --- Helper: Add a text-update request for a table cell ---
    def add_text_update(cell_location: dict, text: str):
        # In update mode, check if the cell already has text and delete it.
        if slide_id and table_element and table_element.get("table"):
            try:
                row_idx = cell_location["rowIndex"]
                col_idx = cell_location["columnIndex"]
                cell = table_element["table"]["tableRows"][row_idx]["tableCells"][col_idx]
                has_text = False
                if cell.get("text", {}).get("textElements"):
                    for te in cell["text"]["textElements"]:
                        if te.get("textRun", {}).get("content", "").strip() != "":
                            has_text = True
                            break
                if has_text:
                    requests_list.append({
                        "deleteText": {
                            "objectId": table_id,
                            "cellLocation": cell_location,
                            "textRange": {"type": "ALL"}
                        }
                    })
            except Exception as e:
                logging.warning(f"Error checking cell content, skipping deletion: {e}")
        # Then insert the new text.
        requests_list.append({
            "insertText": {
                "objectId": table_id,
                "cellLocation": cell_location,
                "insertionIndex": 0,
                "text": text
            }
        })

    # --- Update table cells ---
    add_text_update({"rowIndex": 0, "columnIndex": 1}, title)
    add_text_update({"rowIndex": 1, "columnIndex": 1}, date_time)
    add_text_update({"rowIndex": 2, "columnIndex": 1}, channel)
    add_text_update({"rowIndex": 3, "columnIndex": 1}, description)
    add_text_update({"rowIndex": 4, "columnIndex": 1}, moderators)
    add_text_update({"rowIndex": 5, "columnIndex": 1}, speakers)

    # --- Update slide background color ---
    requests_list.append({
        "updatePageProperties": {
            "objectId": main_slide_id,
            "pageProperties": {
                "pageBackgroundFill": {
                    "solidFill": {
                        "color": {"rgbColor": background_color}
                    }
                }
            },
            "fields": "pageBackgroundFill.solidFill.color"
        }
    })

    # --- Update speaker notes ---
    if speaker_notes_id:
        # In update mode, clear existing speaker notes if any
        if slide_id and slide_data:
            notes_page = slide_data.get("slideProperties", {}).get("notesPage", {})
            for pe in notes_page.get("pageElements", []):
                if pe.get("objectId") == speaker_notes_id and pe.get("shape", {}).get("text", {}).get("textElements"):
                    notes_content = "".join(
                        [te.get("textRun", {}).get("content", "") for te in pe["shape"]["text"]["textElements"]])
                    if notes_content.strip() not in ["", "\n"]:
                        requests_list.append({
                            "deleteText": {
                                "objectId": speaker_notes_id,
                                "textRange": {"type": "ALL"}
                            }
                        })
                    break
        requests_list.append({
            "insertText": {
                "objectId": speaker_notes_id,
                "insertionIndex": 0,
                "text": speaker_note
            }
        })

    # --- Reset text background colors for the entire cells before applying new formatting ---
    # Reset speakers cell (row index 5, column index 1)
    requests_list.append({
        "updateTextStyle": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 5, "columnIndex": 1},
            "textRange": {"type": "ALL"},
            "style": {
                "backgroundColor": {}  # Clear any background color so that it becomes transparent
            },
            "fields": "backgroundColor"
        }
    })
    # Reset moderators cell (row index 4, column index 1)
    requests_list.append({
        "updateTextStyle": {
            "objectId": table_id,
            "cellLocation": {"rowIndex": 4, "columnIndex": 1},
            "textRange": {"type": "ALL"},
            "style": {
                "backgroundColor": {}  # Clear any background color so that it becomes transparent
            },
            "fields": "backgroundColor"
        }
    })

    # --- Update text styles for speakers ---
    speaker_entries = split_people_and_indices(speakers)
    for idx, entry in enumerate(speaker_entries):
        if idx < len(speaker_colors_list):
            color = speaker_colors_list[idx]
            if color != "unknown" and color in COLORS:
                requests_list.append({
                    "updateTextStyle": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": 5, "columnIndex": 1},
                        "textRange": {
                            "type": "FIXED_RANGE",
                            "startIndex": entry["startIndex"],
                            "endIndex": entry["endIndex"]
                        },
                        "style": {
                            "backgroundColor": {
                                "opaqueColor": {"rgbColor": COLORS[color]}
                            }
                        },
                        "fields": "backgroundColor"
                    }
                })

    # --- Update text styles for moderators ---
    moderator_entries = split_people_and_indices(moderators)
    for idx, entry in enumerate(moderator_entries):
        if idx < len(moderator_colors_list):
            color = moderator_colors_list[idx]
            if color != "unknown" and color in COLORS:
                requests_list.append({
                    "updateTextStyle": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": 4, "columnIndex": 1},
                        "textRange": {
                            "type": "FIXED_RANGE",
                            "startIndex": entry["startIndex"],
                            "endIndex": entry["endIndex"]
                        },
                        "style": {
                            "backgroundColor": {
                                "opaqueColor": {"rgbColor": COLORS[color]}
                            }
                        },
                        "fields": "backgroundColor"
                    }
                })

    logging.info("Batch update requests: " + str(requests_list))
    body = {"requests": requests_list}
    response = service.presentations().batchUpdate(
        presentationId=presentation_id,
        body=body
    ).execute()
    return response


# ================================================
# Main Job: Process Airtable records and update slides
# ================================================
def main():
    # 1. Get current time and compute timestamp one hour ago (UTC)
    now = datetime.datetime.now(datetime. UTC)
    one_hour_ago = now - datetime.timedelta(hours=1)
    one_hour_ago_iso = one_hour_ago.isoformat() + "Z"
    logging.info("Fetching sessions modified after " + one_hour_ago_iso)
    print("Fetching sessions modified after " + one_hour_ago_iso)

    # 2. Query Airtable “Sessions” table
    # Using Airtable’s filterByFormula to select records modified in the last hour and having a Slide ID.
    filter_formula = f"AND(IS_AFTER({{Last Modified}}, '{one_hour_ago_iso}'), {{Slide ID}})"

    api = Api(AIRTABLE_API_KEY)
    sessions_table = api.table(AIRTABLE_BASE_ID, SESSIONS_TABLE)
    records = sessions_table.all(formula=filter_formula)
    logging.info(f"Found {len(records)} record(s) to process.")
    print(f"Found {len(records)} record(s) to process.")

    slides_service = get_slides_service()

    for record in records[:1]:
        fields = record.get("fields", {})

        # --- Slide ID ---
        slide_id = fields.get("Slide ID")
        if not slide_id:
            continue  # Skip records with no Slide ID

        # --- Curation Status ---
        curation_status_field = fields.get("Curation Status")
        if isinstance(curation_status_field, list) and curation_status_field:
            cs_id = curation_status_field[0]
            cs_name = get_airtable_record_column_value(AIRTABLE_BASE_ID, api, CURATION_STATUS_TABLE, cs_id, "Status")
            curation_status_name = cs_name if cs_name else "No Status Available"
        else:
            curation_status_name = "No Status Available"

        # --- Speakers & Moderators ---
        # Use the detailed fields that include name, title, and company.
        speakers_raw = fields.get("Speaker Name & Title (from Speaker)", [])
        moderators_raw = fields.get("Moderator Name & Title", [])
        # Apply adjust_representation to each element.
        speakers_list = [adjust_representation(s) for s in speakers_raw]
        moderators_list = [adjust_representation(s) for s in moderators_raw]
        # Join with commas (with no extra space) as required.
        speakers_str = ",".join(speakers_list)
        moderators_str = ",".join(moderators_list)

        # Use the status fields (if available) to map to colors.
        speakers_statuses = fields.get("S Status (from Speaker)", [])
        moderators_statuses = fields.get("S Status (from Moderator)", [])
        speaker_colors_str = ",".join([status_to_color_map.get(s.strip(), 'unknown') for s in speakers_statuses])
        moderator_colors_str = ",".join([status_to_color_map.get(s.strip(), 'unknown') for s in moderators_statuses])

        # --- Additional Fields ---
        date_time = fields.get("S25 Start Date/Time", "")
        note = fields.get("Notes", "")
        title = fields.get("Session Title (<100 characters)", "")
        channel = fields.get("W Channel Text", "")
        description = fields.get("Description (<2500 characters)", "")

        # --- Update (or create) the slide ---
        try:
            result = update_presentation_with_slide(
                service=slides_service,
                presentation_id=GOOGLE_PRESENTATION_ID,
                slide_object_id=TEMPLATE_SLIDE_ID,
                status=curation_status_name,
                date_time=date_time,
                note=note,
                title=title,
                channel=channel,
                description=description,
                speakers=speakers_str,
                speaker_note=note,  # Using the "Notes" field as speaker notes (adjust if needed)
                speaker_colors=speaker_colors_str,
                moderators=moderators_str,
                moderator_colors=moderator_colors_str,
                partners="",
                partners_colors="",
                slide_id=slide_id  # Update mode since Slide ID exists
            )
            logging.info(f"Slide update result for record {record.get('id')}: {result}")
            print(f"Slide update result for record {record.get('id')}: {result}")
        except Exception as e:
            logging.error(f"Error updating slide for record {record.get('id')}: {e}")
            print(f"Error updating slide for record {record.get('id')}: {e}")


if __name__ == "__main__":
    main()
