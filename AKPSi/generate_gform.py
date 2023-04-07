import os
import openpyxl
import PyPDF2
from google.oauth2 import service_account
from googleapiclient.discovery import build
import sys
from link_rushees import create_candidate_page_dict, randomize_candidate_pages
from google.oauth2.credentials import Credentials


def create_randomized_pdf(original_pdf, output_pdf, randomized_candidate_page_dict, k):
    # Read the original PDF
    pdf_reader = PyPDF2.PdfFileReader(original_pdf)

    # Create a new PDF
    pdf_writer = PyPDF2.PdfFileWriter()

    # Process candidates
    randomized_candidate_page_dict = sorted(
        randomized_candidate_page_dict.items(), key=lambda x: x[1])

    print(randomized_candidate_page_dict)
    # Add the first 2 pages to the new PDF
    pdf_writer.addPage(pdf_reader.getPage(0))
    pdf_writer.addPage(pdf_reader.getPage(1))

    # Add the first k entries
    for candidate, page_number in randomized_candidate_page_dict[:k - 3]:
        random_page_number = candidate_page_dict[candidate]
        pdf_writer.addPage(pdf_reader.getPage(random_page_number))

    # # Add the kth page to the new PDF
    pdf_writer.addPage(pdf_reader.getPage(k - 1))

    # Add the remaining candidates
    for candidate, page_number in randomized_candidate_page_dict[k - 3:]:
        random_page_number = candidate_page_dict[candidate]
        pdf_writer.addPage(pdf_reader.getPage(random_page_number))

    # Save the new PDF
    with open(output_pdf, 'wb') as output:
        pdf_writer.write(output)


def create_google_form(randomized_candidate_page_dict, k, credentials_file):
    SCOPES = ['https://www.googleapis.com/auth/forms']

    # Load credentials
    creds = service_account.Credentials.from_service_account_file(
        credentials_file, SCOPES)

    # Build the Google Forms API client
    service = build('forms', 'v1', credentials=creds)

    # Create the form
    form_metadata = {
        'title': 'Individual Interview Delibs',
        'description': 'Please select "yes", "no", or "maybe" for each candidate and select your options after the discussion of a candidate has been completed.\n\nIf rushees have a similar name to another rushee, they have been labeled with "(Namesake)." For some rushees, their namesake may not have moved on to this round.',
    }
    form = service.forms().create(body=form_metadata).execute()
    form_id = form.get('formId')

    # Add first page
    first_page_section = {
        'section': {
            'title': 'First-Years',
            'items': []
        }
    }
    for candidate in list(randomized_candidate_page_dict.keys())[:k-1]:
        first_page_section['section']['items'].append({
            'questionItem': {
                'question': {
                    'questionText': candidate,
                    'required': True,
                    'questionType': 'RADIO',
                    'options': ['Yes', 'No', 'Maybe']
                }
            }
        })
    service.forms().sections().create(formId=form_id, body=first_page_section).execute()

    # Add second page
    second_page_section = {
        'section': {
            'title': 'Second-Years',
            'items': []
        }
    }
    for candidate in list(randomized_candidate_page_dict.keys())[k-1:]:
        second_page_section['section']['items'].append({
            'questionItem': {
                'question': {
                    'questionText': candidate,
                    'required': True,
                    'questionType': 'RADIO',
                    'options': ['Yes', 'No', 'Maybe']
                }
            }
        })
    service.forms().sections().create(formId=form_id,
                                      body=second_page_section).execute()

    print(
        f"Google Form created: https://docs.google.com/forms/d/{form_id}/edit")


if __name__ == '__main__':
    if len(sys.argv) != 6:
        print("Usage: python randomize_candidate_pages.py excel_file pdf_file k output_pdf credentials")
        sys.exit(1)

    excel_file = sys.argv[1]
    pdf_file = sys.argv[2]
    k = int(sys.argv[3])
    output_pdf = sys.argv[4]
    credentials_file = sys.argv[5]

    candidate_page_dict = create_candidate_page_dict(excel_file, pdf_file, k)
    randomized_candidate_page_dict = randomize_candidate_pages(
        candidate_page_dict, k)

    create_randomized_pdf(pdf_file, output_pdf,
                          randomized_candidate_page_dict, k)
    create_google_form(randomized_candidate_page_dict, k, credentials_file)

    print("New randomized PDF created.")
