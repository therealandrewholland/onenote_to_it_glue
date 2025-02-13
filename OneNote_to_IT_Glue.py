import os
import re
import sys
import time
import uuid
import json
import string
import shutil
import zipfile
import traceback
from urllib.parse import unquote
import win32com.client
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement
from docx.enum.style import WD_STYLE_TYPE

#Global variable used to store what needs review
manual_check_needed = {}
client_name = ""

def getNotebookID(path_or_url, url_file=False):
        global client_name

        app = win32com.client.dynamic.Dispatch('OneNote.Application')

         #wait to finish opening notebook before continuing
        if url_file:
            #open url in OneNote
            app.NavigateToUrl(path_or_url, False)

            #wait to finish opening notebook before continuing
            time.sleep(10)
            while True:
                time.sleep(3)

                try:
                    open_notebooks = app.GetHierarchy("" , 1)
                    break
                except:
                    continue

            root = ET.fromstring(open_notebooks)
            namespaces = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}
            all_notebooks = root.findall('one:Notebook', namespaces)

            #complicated but recreates the client folder from the url to compare with the decoded url onenote saves
            #this part may need to be changed for documentation that does not follow the pattern:
            #https://aunalytics.sharepoint.com/sites/Aunalytics2/Document%20Library/[FOLDER]/[CLIENT NAME]/
            client_folder = unquote(path_or_url[8:].split("/")[7])
            for notebook in all_notebooks:
                if notebook.get('path').split("/")[7] == client_folder:
                    notebook_id = notebook.get('ID')
                    break

            if client_name == None:
                client_name = client_folder

        else:
            app = win32com.client.dynamic.Dispatch('OneNote.Application')

            #store older open notebooks / sections
            old_open_notebooks = app.GetHierarchy("" , 1)

            app.OpenHierarchy(path_or_url, "", "", 0)

            #store newer open notebooks / sections after opening another
            new_open_notebooks = app.GetHierarchy("" , 1)

            #get ID from the difference between old / new (this lets it handle both notebooks and sections)
            result = ""
            for i in range(len(new_open_notebooks)):
                if i >= len(old_open_notebooks) or new_open_notebooks[i] != old_open_notebooks[i]:
                    result += new_open_notebooks[i]

            #this part may need to be changed for documentation that does not follow the pattern:
            #C:\Users\andrew.holland\OneDrive - Aunalytics\[FOLDER]\[CLIENT NAME]\
            if client_name == None:
                client_name = path_or_url.split("\\")[5]

            # Remove uneeded information around id
            notebook_id = result.split("\"")[1]

        return notebook_id

#url can be path to .one or formatted URL that starts with "onenote:"
def getID(url_or_path):
    global manual_check_needed
    global client_name

    app = win32com.client.dynamic.Dispatch('OneNote.Application')

    if url_or_path.startswith("onenote:"):
        #This will fetch the highest ID for the .one file
        notebook_id = getNotebookID(url_or_path, url_file=True)
    else:
        #for akzo. This will fetch the highest ID for the .one file
        notebook_id = getNotebookID(url_or_path, url_file=False)

    #wait for all sections to sync
    time.sleep(10)
    main_sections = app.GetHierarchy(notebook_id, 1)
    root = ET.fromstring(main_sections)
    namespaces = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}
    all_sections = root.findall('one:Section', namespaces)

    # This dictionary will store the "Production" section id and only overwrite it if there is a "Production 2" or some variation
    id_list = {
        "name": None,
        "ID": None
    }

    production_dict = {}

    for section in all_sections:
        name = section.get('name')
        
        production_dict[name] = section.get('ID')
        # For Akzo. If there is only one section, grabs its id instead of the "Production" section
        if "Production" in name or len(all_sections) == 1:
            if id_list["name"] == None or "2" in name:
                id_list["name"] = name
                id_list["ID"] = section.get('ID')

    #lets user select section to export when multiple found
    printable_keys = production_dict.keys()
    if len(printable_keys) > 1:
        option_dict = {}
        i = 0
        print("Many exportable sections found. Please select one to export.")
        for key in printable_keys:
            i += 1
            print(f"{i}. {key}")
            option_dict[str(i)] = production_dict[key]

        user_input = 0     
        while user_input not in option_dict.keys():
            user_input = input(f"Enter an option (1 to {i}):")

        print("Exporting...")
        target_section_id = option_dict[user_input]

    target_section_id = id_list["ID"]

    return target_section_id, notebook_id


def createDoc(target_section_id, client):
    app = win32com.client.dynamic.Dispatch('OneNote.Application')
    docx_path = os.path.join(os.getcwd(), f"temp\\{client}.docx")
    app.Publish(target_section_id, docx_path, 5, "")


def getHeadingsFromNotebook(target_section_id):
    app = win32com.client.dynamic.Dispatch('OneNote.Application')

    # Get all notebook info
    notebooks_xml = app.GetHierarchy(target_section_id, 4)
    root = ET.fromstring(notebooks_xml)
    namespace = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}

    #Used for checking full notebook as raw .xml
    #print(notebooks_xml)

    #Track heading text as key and "Heading" level as value
    heading_levels_dict = {}
    for page in root.findall('one:Page', namespace):
        page_name = page.get('name')
        page_level = page.get('pageLevel')

        heading_levels_dict[page_name] = page_level

    return heading_levels_dict


def getHeadingsFromDocx(document, target_section_id):
    heading_tracking_listA = []
    heading_tracking_listB =[]
    heading_tracking_number = 0
    heading_levels_dict = getHeadingsFromNotebook(target_section_id)

    # Iterate through .docx to identify sections / subsetions by
    for paragraph in document.paragraphs:
        if paragraph.style.name in ['Heading 1', 'Heading 2', 'Heading 3']:
            paragraph.style = "Normal"
        text = paragraph.text

        if (text in heading_levels_dict):
            heading_tracking_listA.append(heading_tracking_number)
        if (':' in text and (text.endswith(" AM") or text.endswith(" PM"))):
            heading_tracking_listB.append(heading_tracking_number-2)

        heading_tracking_number += 1

        #Used for checking value of every parahraph in .docx
        #print(paragraph.text)
        #print("- - - - - - - - - - -")
    heading_tracking_list = list(set(heading_tracking_listA) & set(heading_tracking_listB))

    return heading_tracking_list, heading_levels_dict


def formatHeadings(client, target_section_id):
    global manual_check_needed
    global client_name

    document = Document(f"temp\\{client}.docx")

    if len(document.paragraphs) == 1:
        failed_list = findCorruptSections(target_section_id)
        string = f"Error exporting from OneNote file. Recreate corrupted sections: {failed_list}"
        print(string)
        manual_check_needed[client_name].append(string)

        
    heading_tracking_list, heading_levels_dict = getHeadingsFromDocx(document, target_section_id)

    # For each paragraph that needs to be a heading, add level (i.e. "Heading 1", "Heading 2")
    for i in heading_tracking_list:
        paragraph = document.paragraphs[i]
        key = paragraph.text
        heading_level = heading_levels_dict[key]
        heading_convert = {
            '1': 'Heading 1',
            '2': 'Heading 2',
            '3': 'Heading 3'
        }

        #Used for checking headings and their level
        #print(key)
        #print(heading_convert[heading_level])
        #print("- -- -- -- -- -- -")

        document.paragraphs[i].style = heading_convert[heading_level]

    # .docx files show weird when uploaded to IT Glue
    document.save(f"temp\\final_docx\\{client}.docx")
    os.remove(f"temp\\{client}.docx")

# TO DO
def extractDocx(document_name, extract_folder_name):
    with zipfile.ZipFile(document_name, 'r') as zip_ref:
        zip_ref.extractall(extract_folder_name)

    return os.path.join(extract_folder_name, 'word', 'styles.xml')

def rezipDocx(document_name, extract_folder_name):
    with zipfile.ZipFile(document_name, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
        for foldername, subfolders, filenames in os.walk(extract_folder_name):
            for filename in filenames:
                filepath = os.path.join(foldername, filename)
                arcname = os.path.relpath(filepath, extract_folder_name)
                docx_zip.write(filepath, arcname)

#pull Headings styles from template .Zip and add them to original
def createStyles(client):
    document = Document(f"temp\\{client}.docx")

    template_path = ""
    if getattr(sys, 'frozen', False):
        # If the application is frozen (running as a .exe)
        base_path = sys._MEIPASS
    else:
        # If the application is running normally
        base_path = os.path.dirname(__file__)

    template_path = os.path.join(base_path, 'template.docx')

    template_document_xml_path = extractDocx(template_path, "temp\\extracted_contents\\template")
    document_xml_path = extractDocx(f"temp\\{client}.docx", "temp\\extracted_contents\\test")

    os.remove(document_xml_path)
    os.remove(f"temp\\{client}.docx")
    shutil.copy2(template_document_xml_path, document_xml_path)

    rezipDocx(f"temp\\{client}.docx", "temp\\extracted_contents\\test")

    document = Document(f"temp\\{client}.docx")

    cleanup_list = ["temp\\extracted_contents\\template", "temp\\extracted_contents\\test"]
    for directory in cleanup_list:
        shutil.rmtree(directory)

#Fixes issue with formatted docx not working in IT Glue by opening them in Word and converting to .doc
def convertToDoc(client, formatted_client):
    current_directory = os.getcwd()

    word = win32com.client.dynamic.Dispatch("Word.Application")
    word.Visible = False

    document = word.Documents.Open(f"{current_directory}\\temp\\final_docx\\{formatted_client}.docx")

    #Adds client name to start of file so the file name shows correctly in IT Glue
    range = document.Content
    range.InsertBefore(client + ".OneNote\n")
    range = document.Range(0, len(client) + 1)
    range.Style = "Heading 1"

    document.SaveAs(f"{current_directory}\\formatted_docs\\{client}.doc", FileFormat=0)

    document.Close()
    word.Quit()
    os.remove(f"temp\\final_docx\\{formatted_client}.docx")

def findCorruptSections(target_section_id):
    app = win32com.client.dynamic.Dispatch('OneNote.Application')
    notebook_info = app.GetHierarchy(target_section_id, 4)
    root = ET.fromstring(notebook_info)
    namespace = "{http://schemas.microsoft.com/office/onenote/2013/onenote}Page"

    os.makedirs("temp\\find_corrupt_section\\")
    failed_list = []

    for page in root.iter(namespace):
        page_name = page.get('name')
        formatted_page_name = page_name.lower().translate(str.maketrans('', '', string.punctuation))
        page_id = page.get('ID')
        docx_path = os.path.join(os.getcwd(), f"temp\\find_corrupt_section\\{formatted_page_name}.docx")
        try:
            app.Publish(page_id, docx_path, 5, "")
            if isDocxCorrupt(docx_path):
                failed_list.append(page_name)
        except:
            failed_list.append(page_name)

    shutil.rmtree("temp\\find_corrupt_section\\")

    return failed_list

def isDocxCorrupt(docx_path):
    document = Document(docx_path)
    os.remove(docx_path)
    
    if len(document.paragraphs) == 1:
        return True

    return False

def convertToITGlue(url_or_path, user_input_client_name):
    global manual_check_needed
    global client_name

    notebook_id = None
    client_name = user_input_client_name
    manual_check_needed[client_name] = []

    try:
        #getID also adds a value to client_name
        target_section_id, notebook_id = getID(url_or_path)
        formatted_client = client_name.replace(',', '').lower().replace(' ', '_')

        if not os.path.exists(f"formatted_docs\\{client_name}.doc"):
            createDoc(target_section_id, formatted_client)
            createStyles(formatted_client)
            formatHeadings(formatted_client, target_section_id)

            if len(manual_check_needed[client_name]) == 0:
                convertToDoc(client_name, formatted_client)
                print(f"The OneNote for {client_name} has successfully been converted to a IT Glue formatted .doc")

    #Catch and print all errors for failed clients
    except Exception as e:
        print(f"{client_name} failed. Skipping. Error message:")
        print(e)
        traceback.print_exc()

    finally:
        #close any opened notebooks, even if the client fails to convert
        if not notebook_id == None:
            app = win32com.client.dynamic.Dispatch('OneNote.Application')
            try:
                app.CloseNotebook(notebook_id)
            except:
                print(f"Unable to close notebook or section for {client_name}")
                pass

    if manual_check_needed:
        with open(f"need_review\\{client_name}.json", 'w') as json_file:
            json.dump(manual_check_needed, json_file)

    manual_check_needed = {}
    client_name = ""
