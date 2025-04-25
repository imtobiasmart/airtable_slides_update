# Airtable to Google Slides Sync

A Python application that synchronizes session data from Airtable to Google Slides, automatically updating presentation slides when records are modified.

## Overview

This application monitors an Airtable base for recently modified session records and updates corresponding slides in a Google Slides presentation. It's designed for event management where session information (speakers, moderators, titles, descriptions, etc.) needs to be visualized in a consistent slide format.

## Features

- Automatically syncs Airtable session data to Google Slides
- Updates existing slides or creates new ones from a template
- Color-codes speakers and moderators based on their status
- Handles formatting for names, titles, and companies
- Updates slide background colors based on session status
- Captures session notes in speaker notes section

## Requirements

- Python 3.6+
- Google Cloud account with Slides API enabled
- Airtable account and API key
- Service account credentials with Google Slides access

## Dependencies

- `google-auth` - Google authentication
- `google-api-python-client` - Google API client
- `pyairtable` - Airtable API client

## Installation

1. Clone this repository:
   ```
   git clone <repository-url>
   cd airtable-slides-sync
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Set up environment variables (see Configuration section)

## Configuration

Set the following environment variables:

### Airtable Configuration
- `AIRTABLE_API_KEY` - Your Airtable API key
- `AIRTABLE_BASE_ID` - The ID of your Airtable base
- `SESSIONS_TABLE` - Name of the sessions table (default: "Sessions")
- `CURATION_STATUS_TABLE` - Name of the status table (default: "Statuses")
- `SPEAKERS_TABLE` - Name of the speakers table (default: "Speaker")

### Google Slides Configuration
- `GOOGLE_PRESENTATION_ID` - ID of your Google Slides presentation
- `TEMPLATE_SLIDE_ID` - ID of the template slide to duplicate
- `GOOGLE_SERVICE_ACCOUNT_FILE` - Path to your Google service account credentials JSON file

## Airtable Structure Requirements

Your Airtable base should have the following tables and fields:

### Sessions Table
- `Slide ID` - ID of the corresponding slide in Google Slides
- `Curation Status` - Linked to a status record
- `Speaker Name & Title (from Speaker)` - List of speaker names with titles
- `Moderator Name & Title` - List of moderator names with titles
- `S Status (from Speaker)` - Speaker status values for color coding
- `S Status (from Moderator)` - Moderator status values for color coding
- `S25 Start Date/Time` - Session date and time
- `Notes` - Session notes
- `Session Title (<100 characters)` - Session title
- `W Channel Text` - Session channel or track
- `Description (<2500 characters)` - Session description
- `Last Modified` - Timestamp of last modification

### Status Table
- `Status` - Name of the status

## Usage

Run the script manually or set it up as a scheduled job:

```
python main.py
```

The script will:
1. Check for Airtable session records modified in the last hour
2. For each modified record with a Slide ID, retrieve the relevant data
3. Update the corresponding Google Slide with the latest information

## Google Slides Template Structure

Your template slide should contain:
- A table with the following structure:
  - Row 0, Column 1: Session Title
  - Row 1, Column 1: Date and Time
  - Row 2, Column 1: Channel
  - Row 3, Column 1: Description
  - Row 4, Column 1: Moderators
  - Row 5, Column 1: Speakers
- A speaker notes section for session notes

## Status-to-Color Mapping

The script maps status values to background colors for speakers and moderators:
- Green: Registered, No Further Action Needed, Program Representation Secured
- Yellow: In Progress, Meeting states, Waiting for Content Ideas, etc.
- Red: Idea, Inbound, Ask Needed, etc.
- Orange: Confirmed but Needs Registration, Missing Information

## Logging

The script logs information and errors to the standard output and to the default Python logging configuration.

## Error Handling

The script provides error messages when:
- A slide ID is provided but the slide is not found
- Required page elements are missing
- API requests fail
