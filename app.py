import pandas as pd
import os
import PyPDF2
import re
import docx
import logging

logging.basicConfig()


def list_files(directory):
    files = os.listdir(directory)
    file_names = [
        file for file in files
        if os.path.isfile(os.path.join(directory, file)) and
           (file.lower().endswith('.pdf')
            or
            file.lower().endswith('.doc'))
    ]
    return file_names


def collect_all_text(file_names):
    text = []
    for file in file_names:
        if file.lower().endswith('.pdf'):
            try:
                text.append(pdf_to_text(os.path.join(directory, file)))
            except Exception as e:
                logging.error(f"Error while extracting text from PDF file {file}: {str(e)}")
        elif file.lower().endswith('.doc'):
            try:
                text.append(doc_to_text(os.path.join(directory, file)))
            except Exception as e:
                logging.error(f"Error while extracting text from DOC file {file}: {str(e)}")
    return text


def collect_all_emails(texts: list):
    emails = []
    for t in texts:
        try:
            emails.append(extract_emails(t))
        except Exception as e:
            logging.error(f"Error while extracting emails from text: {str(e)}")
    return emails


def doc_to_text(path: str):
    try:
        doc = docx.Document(path)
        text = ' '.join(paragraph.text for paragraph in doc.paragraphs)
        return text
    except Exception as e:
        logging.error(f"Error while extracting text from DOC file {path}: {str(e)}")
        return ""


def pdf_to_text(path: str):
    try:
        pdfFileObj = open(path, 'rb')
        reader = PyPDF2.PdfReader(pdfFileObj)
        pageObj = reader.pages[0]
        text = pageObj.extract_text()
        return text
    except Exception as e:
        logging.error(f"Error while extracting text from PDF file {path}: {str(e)}")
        return ""


def extract_emails(text):
    try:
        # text = text.replace(' ', '')
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'
        emails = re.findall(email_pattern, text)
        logging.info(f'{emails}')
        return emails
    except Exception as e:
        logging.error(f"Error while extracting emails from text: {str(e)}")
        return []


def to_excel(file_names, emails):
    main_emails = []
    other_emails = []
    for email in emails:
        if len(email) > 0:
            main_emails.append(email[0])
            try:
                other_emails.append(' '.join(email[1:]))
            except:
                other_emails.append("")
        else:
            main_emails.append("")
            other_emails.append("")
    try:
        cvv = {'file_name': file_names,
               'main email': main_emails,
               'other emails': other_emails}
        cv_df = pd.DataFrame(cvv)
        cv_df.to_excel('filename.xlsx')
    except Exception as e:
        logging.error(f"Error while saving to Excel file: {str(e)}")


if __name__ == "__main__":
    directory = '/direcory_of_pdf_files/'
    try:
        file_names = list_files(directory)
        print(file_names)
        texts = collect_all_text(file_names)
        emails = collect_all_emails(texts)
        to_excel(file_names, emails)
    except Exception as e:
        logging.error(f"An error occurred during script execution: {str(e)}")
