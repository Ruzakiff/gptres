from google.oauth2 import service_account
from googleapiclient.discovery import build
import json
from openai import OpenAI

# Replace with your service account file path and document ID
SERVICE_ACCOUNT_FILE = 'rags-416420-02f4be967127.json'
DOCUMENT_ID = '1kFoE75_emnMUTS_G0ibnStu-WyA_4fmi69JuKepOQm4'

# Authenticate and build the service
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=['https://www.googleapis.com/auth/documents']
)
service = build('docs', 'v1', credentials=credentials)

# Fetch the document
def fetch_document(service, document_id):
    return service.documents().get(documentId=document_id).execute()

# Parse the document using GPT-4
client = OpenAI()

def parse_document(text):
    messages = [
        {"role": "system", "content": f"Given the following text from a resume, identify and structure the company name, job title, hours worked, time period, and related bullets for each job experience. The text is:\n\n{text}\n\n PERSERVE ONLY START/END indexs of items parsed, Return the structured data in JSON format with the following keys:START/END INDEX(can be multiple in one key) company_name, job_title, hours_worked, time_period, details (list of bullets). "}]
    response = client.chat.completions.create(
        model="gpt-4o-2024-05-13",
        response_format={ "type": "json_object" },
        messages=messages,
        max_tokens=4095,
        temperature=0,
    )
    return response.choices[0].message.content

# Process the document content and formatting
def process_document(document):
    content = document.get('body').get('content')
    text_with_indices = ""
    document_text = ""  # Store the entire document text
    formatting_details = []

    for element in content:
        if 'paragraph' in element:
            paragraph = element.get('paragraph')
            paragraph_style = paragraph.get('paragraphStyle', {})
            for paragraph_element in paragraph.get('elements'):
                text_run = paragraph_element.get('textRun')
                if text_run:
                    text_content = text_run.get('content')
                    start_index = paragraph_element.get('startIndex')
                    end_index = paragraph_element.get('endIndex')
                    text_with_indices += f"{{start:{start_index},end:{end_index},text:{json.dumps(text_content)}}}"
                    document_text += text_content  # Append to the document text
                    text_style = text_run.get('textStyle', {})
                    bullet_info = paragraph.get('bullet', {})
                    font_size_info = text_style.get('fontSize')
                    font_size = font_size_info.get('magnitude') if font_size_info else "Default or not set"
                    formatting_detail = {
                        'text': text_content.strip(),
                        'start_index': start_index,
                        'end_index': end_index,
                        'text_style': text_style,
                        'bullet_info': bullet_info,
                        'font_size': font_size,
                        'font_family': text_style.get('weightedFontFamily', {}).get('fontFamily'),
                        'font_weight': text_style.get('weightedFontFamily', {}).get('weight'),
                        'bold': text_style.get('bold', False),
                        'italic': text_style.get('italic', False),
                        'underline': text_style.get('underline', False),
                        'strikethrough': text_style.get('strikethrough', False),
                        'foreground_color': text_style.get('foregroundColor', {}).get('color', {}).get('rgbColor'),
                        'background_color': text_style.get('backgroundColor', {}).get('color', {}).get('rgbColor'),
                        'link': text_style.get('link', {}).get('url'),
                        'baseline_offset': text_style.get('baselineOffset', 'NONE'),
                        'small_caps': text_style.get('smallCaps', False),
                        'paragraph_style': paragraph_style,
                        'tags': []  # Initialize tags list
                    }
                    if 'bullet' in paragraph:
                        bullet = paragraph.get('bullet')
                        list_id = bullet.get('listId')
                        nesting_level = bullet.get('nestingLevel', 0)
                        formatting_detail.update({
                            'list_id': list_id,
                            'nesting_level': nesting_level
                        })
                    formatting_details.append(formatting_detail)

    print(text_with_indices)
    parsed_json = json.loads(parse_document(text_with_indices))
    print(parsed_json)

    # Tag formatting details based on parsed JSON
    for section_key, experiences in parsed_json.items():
        if isinstance(experiences, list):
            for experience in experiences:
                company_name = experience.get('company_name')
                for key, value in experience.items():
                    tag_detail_in_formatting_details(value, formatting_details, key, document_text)

    return parsed_json, formatting_details

def tag_detail_in_formatting_details(value, formatting_details, tag_type, document_text):
    # Check if value is a list and handle it appropriately
    if isinstance(value, list):
        for item in value:
            tag_detail_in_formatting_details(item, formatting_details, tag_type, document_text)
    elif isinstance(value, dict):
        # Extract the actual text using the start and end indices if it's a company name
        if tag_type == 'company_name':
            # Find the text run that matches the start and end indices
            company_name = ""
            for fmt_detail in formatting_details:
                if fmt_detail['start_index'] >= value['start'] and fmt_detail['end_index'] <= value['end']:
                    company_name += fmt_detail['text']
        else:
            company_name = None

        # Proceed with the original logic if value is a dictionary
        for fmt_detail in formatting_details:
            start_index = fmt_detail['start_index']
            end_index = fmt_detail['end_index']
            if start_index < value.get('end', float('inf')) and end_index > value.get('start', -1):
                if tag_type not in fmt_detail['tags']:
                    fmt_detail['tags'].append(tag_type)
                if company_name:
                    fmt_detail['tags'].append(f"company_name:{company_name}")
    else:
        # Handle unexpected value types (e.g., neither list nor dict)
        print(f"Unexpected type of value in tag_detail_in_formatting_details: {type(value)}")

# Update Google Document
def update_google_doc(service, document_id, requests):
    result = service.documents().batchUpdate(
        documentId=document_id, body={'requests': requests}).execute()
    return result

# Replace text in section
def replace_text_in_section(service, document_id, section_name, old_text, new_text):
    document = fetch_document(service, document_id)
    sections = process_document(document)
    requests = []
    start_index = None
    end_index = None
    if section_name in sections:
        for item in sections[section_name]:
            if item['company_name'] == old_text:
                company_name_formatting = item.get('company_name_formatting')
                if company_name_formatting:
                    start_index = company_name_formatting['start_index']
                    end_index = company_name_formatting['end_index']
                    requests.append({
                        'replaceAllText': {
                            'containsText': {
                                'text': old_text,
                                'matchCase': True
                            },
                            'replaceText': new_text
                        }
                    })
                    print('Request added for Google Doc update.')
                break
    return requests, start_index, end_index

# Adjust layout
def adjust_layout(service, document_id, start_index, end_index, alignment):
    requests = [
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                },
                'paragraphStyle': {
                    'alignment': alignment,
                    'indentStart': {
                        'magnitude': 0,
                        'unit': 'PT'
                    },
                    'indentEnd': {
                        'magnitude': 0,
                        'unit': 'PT'
                    },
                    'spaceAbove': {
                        'magnitude': 0,
                        'unit': 'PT'
                    },
                    'spaceBelow': {
                        'magnitude': 0,
                        'unit': 'PT'
                    }
                },
                'fields': 'alignment,indentStart,indentEnd,spaceAbove,spaceBelow'
            }
        }
    ]
    update_google_doc(service, document_id, requests)

# Adjust text position
def adjust_text_position(service, document_id, start_index, end_index):
    requests = [
        {
            'deleteContentRange': {
                'range': {
                    'startIndex': end_index,
                    'endIndex': end_index + 1
                }
            }
        }
    ]
    update_google_doc(service, document_id, requests)

def map_text_runs_to_parsed_units(parsed_units, formatting_details):
    changes = []

    for item in parsed_units.values():
        if isinstance(item, list):
            for sub_item in item:
                if isinstance(sub_item, dict):
                    for key, value in sub_item.items():
                        sub_item_text = ""
                        sub_item_formatting = []
                        for fmt in formatting_details:
                            if value in fmt['text']:
                                sub_item_text += fmt['text']
                                sub_item_formatting.append(fmt)
                        if sub_item_text == value:
                            changes.append((sub_item, f"{key}_formatting", sub_item_formatting))

    # Apply changes after iteration
    for change in changes:
        change[0][change[1]] = change[2]

    return parsed_units

def collect_formatting_details(parsed_units):
    for unit in parsed_units:
        unit['all_formatting'] = []
        for formatting in unit.get('formatting', []):
            unit['all_formatting'].extend(formatting)
    return parsed_units

def edit_text_run_by_tags_or_content(service, document_id, search_tags=None, search_text=None, new_text=None, new_formatting=None):
    document = fetch_document(service, document_id)
    parsed_json, formatting_details = process_document(document)
    
    requests = []
    found = False

    # Iterate through formatting details to find matching text runs
    for detail in formatting_details:
        tag_match = all(tag in detail['tags'] for tag in search_tags) if search_tags else False
        text_match = search_text in detail['text'] if search_text else False

        if tag_match or text_match:
            found = True
            start_index = detail['start_index']
            end_index = detail['end_index']

            # If new text is provided, generate a request to replace the text
            if new_text:
                requests.append({
                    'replaceAllText': {
                        'containsText': {
                            'text': detail['text'],
                            'matchCase': True
                        },
                        'replaceText': new_text
                    }
                })

            # If new formatting is provided, generate requests to update formatting
            if new_formatting:
                # Example of updating bold formatting; extend as needed for other properties
                if 'bold' in new_formatting:
                    requests.append({
                        'updateTextStyle': {
                            'range': {
                                'startIndex': start_index,
                                'endIndex': end_index
                            },
                            'textStyle': {
                                'bold': new_formatting['bold']
                            },
                            'fields': 'bold'
                        }
                    })

            # Break after the first match if only one update is needed
            break

    if not found:
        print("No matching text run found.")
        return

    # Execute all requests to update the document
    if requests:
        update_google_doc(service, document_id, requests)
        print("Document updated successfully.")
    else:
        print("No updates to perform.")

def print_experience_indices(parsed_json):
    experiences = parsed_json.get('experiences', parsed_json.get('job_experiences', []))
    for i, experience in enumerate(experiences):
        # Initialize with high start and low end to find the min and max respectively
        start_indices = []
        end_indices = []
        
        
        # Loop through each key in the experience dictionary
        for key, value in experience.items():
            if isinstance(value, dict):  # Single item dict
                start_indices.append(value['start'])
                end_indices.append(value['end'])
            elif isinstance(value, list):  # List of dicts
                for item in value:
                    start_indices.append(item['start'])
                    end_indices.append(item['end'])
        
        # Find the minimum start index and maximum end index
        if start_indices and end_indices:
            min_start = min(start_indices)
            max_end = max(end_indices)
            print(f"Company {i+1} starts at index {min_start} and ends at index {max_end}")

# Example usage:
#edit_text_run_by_tags_or_content(service, DOCUMENT_ID, search_tags=['company_name'], new_text="New Company Name", new_formatting={'bold': True})

if __name__ == "__main__":
    # Fetch the document
    document = fetch_document(service, DOCUMENT_ID)
    parsed_json, formatting_details = process_document(document)
    print_experience_indices(parsed_json)
    # Print details for all text runs
    for detail in formatting_details:
        print(f"Text: {detail['text']}, Start Index: {detail['start_index']}, End Index: {detail['end_index']}, Tags: {detail['tags']}")
