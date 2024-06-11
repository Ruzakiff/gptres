from google.oauth2 import service_account
from googleapiclient.discovery import build
import json

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

# Parse the experience section using GPT-4
from openai import OpenAI
client = OpenAI()

def parse_experience_section(text):
    messages = [
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": f"Given the following text from a resume, identify and structure the company name, job title, hours worked, time period, and related bullets for each job experience. The text is:\n\n{text}\n\nReturn the structured data in JSON format with the following keys: company_name, job_title, hours_worked, time_period, details (list of bullets)."}
    ]
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
    sections = {
        'contact_info': [],
        'summary': [],
        'experience': [],
        'education': [],
        'skills': [],
        'projects': [],
        'certifications': [],
        'others': []
    }
    
    current_section = 'others'
    experience_text = ""
    formatting_details = []

    for element in content:
        if 'paragraph' in element:
            paragraph = element.get('paragraph')
            bullet = paragraph.get('bullet')
            bullet_info = f"Bullet: {bullet}" if bullet else "No bullet"
            for paragraph_element in paragraph.get('elements'):
                text_run = paragraph_element.get('textRun')
                if text_run:
                    text = text_run.get('content').strip()
                    text_style = text_run.get('textStyle')
                    start_index = paragraph_element.get('startIndex')
                    end_index = paragraph_element.get('endIndex')
                    formatting_details.append({
                        'text': text,
                        'start_index': start_index,
                        'end_index': end_index,
                        'text_style': text_style,
                        'bullet_info': bullet_info,
                        'font_size': text_style.get('fontSize', {}).get('magnitude'),
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
                        'small_caps': text_style.get('smallCaps', False)
                    })
                    
                    # Determine the section based on keywords
                    if text.lower() in ['contact information', 'contact info']:
                        current_section = 'contact_info'
                    elif text.lower() in ['summary', 'professional summary']:
                        current_section = 'summary'
                    elif text.lower() in ['experience', 'work experience', 'professional experience']:
                        current_section = 'experience'
                        experience_text = ""
                    elif text.lower() in ['education', 'academic background']:
                        current_section = 'education'
                    elif text.lower() in ['skills', 'technical skills']:
                        current_section = 'skills'
                    elif text.lower() in ['projects', 'personal projects']:
                        current_section = 'projects'
                    elif text.lower() in ['certifications', 'licenses']:
                        current_section = 'certifications'
                    
                    if current_section == 'experience':
                        experience_text += text + "\n"
                    else:
                        sections[current_section].append({
                            'text': text,
                            'start_index': start_index,
                            'end_index': end_index,
                            'text_style': text_style
                        })
    
    if experience_text:
        j = clean_output(parse_experience_section(experience_text))
        parsed_experience = json.loads(j)
        for job in parsed_experience['work_experience']:
            for i, detail in enumerate(job['details']):
                for fmt in formatting_details:
                    if fmt['text'] == detail:
                        job['details'][i] = fmt
            # Add the company name formatting details
            for fmt in formatting_details:
                if fmt['text'] == job['company_name']:
                    job['company_name_formatting'] = fmt
        sections['experience'] = parsed_experience['work_experience']
    
    return sections

def clean_output(output_str):
    # Find the first opening curly bracket
    start_index = output_str.find("{")
    # Find the last closing curly bracket
    end_index = output_str.rfind("}")
    # Extract everything between the first opening and last closing curly bracket
    cleaned_str = output_str[start_index:end_index+1]
    return cleaned_str


def update_google_doc(service, document_id, requests):
    result = service.documents().batchUpdate(
        documentId=document_id, body={'requests': requests}).execute()
    return result

def replace_text_in_section(service, document_id, section_name, old_text, new_text):
    document = fetch_document(service, document_id)
    sections = process_document(document)
    requests = []
    start_index = None
    end_index = None
    if section_name in sections:
        for item in sections[section_name]:
            if item['company_name'] == old_text:
                # Use the company_name_formatting to get the start and end index
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

# Example usage
document = fetch_document(service, DOCUMENT_ID)
sections = process_document(document)

# Access a specific job and its details
experience_section = sections['experience']
for job in experience_section:
    if job['company_name'] == 'First National Bank':
        print(job)
        # job['job_title'] will give you the job title
        # job['hours_worked'] will give you the hours worked
        # job['time_period'] will give you the time period
        # job['details'] will give you the list of bullets with formatting details

# Replace text in a specific job
requests, start_index, end_index = replace_text_in_section(service, DOCUMENT_ID, 'experience', 'First National Bank', 'New Longer Company Name')
if requests:
    update_google_doc(service, DOCUMENT_ID, requests)
    if start_index is not None and end_index is not None:
        adjust_layout(service, DOCUMENT_ID, start_index, end_index, 'END')
        adjust_text_position(service, DOCUMENT_ID, start_index, end_index)

if __name__ == "__main__":
    # Fetch the document
    document = fetch_document(service, DOCUMENT_ID)
    sections = process_document(document)

    # Access a specific job and its details
    experience_section = sections['experience']
    for job in experience_section:
        if job['company_name'] == 'First National Bank':
            print("Company Name:", job['company_name'])
            print("Formatting Details of Company Name:")
            print(job.get('company_name_formatting', 'No formatting details found'))
