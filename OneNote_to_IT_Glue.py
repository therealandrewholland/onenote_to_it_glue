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

#Global variable used to log store what needs review
manual_check_needed = {}
client_name = ""

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
    isEmpty = False
    
    if len(document.paragraphs) == 1:
        failed_list = findCorruptSections(target_section_id)
        string = f"Error exporting from OneNote file. Recreate corrupted sections: {failed_list}"
        print(string)
        manual_check_needed[client_name].append(string)

    elif len(document.paragraphs) == 3: #empty document
        isEmpty = True

    document.save(f"temp\\final_docx\\{client}.docx")
    os.remove(f"temp\\{client}.docx")
        
    return isEmpty
    
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
def convertToDoc(client, formatted_client, doc_path):
    current_directory = os.getcwd()

    word = win32com.client.dynamic.Dispatch("Word.Application")
    word.Visible = False

    document = word.Documents.Open(f"{current_directory}\\temp\\final_docx\\{formatted_client}.docx")
    document.SaveAs(doc_path, FileFormat=0)

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

def convertToITGlue(target_section_id, section_name, doc_path):
    global manual_check_needed
    global client_name

    notebook_id = None
    client_name = section_name
    manual_check_needed[client_name] = []
    formatted_client = re.sub(r'[<>:"|?*]', '', client_name).lower().replace(' ', '_')

    try:
        #getID also can also add a value to global client_name if its empty
        #target_section_id, notebook_id = getID(url_or_path)

        if not os.path.exists(doc_path):
            createDoc(target_section_id, formatted_client)
            createStyles(formatted_client)
            isEmpty = formatHeadings(formatted_client, target_section_id)

            if len(manual_check_needed[client_name]) == 0 and not isEmpty:
                convertToDoc(client_name, formatted_client, doc_path)
                print(f"The OneNote for {client_name} has successfully been converted to a IT Glue formatted .doc")

    #Catch and print all errors for failed clients
    except Exception as e:
        print(f"client_name: {client_name}")
        print(f"formatted_client: {formatted_client}")
        print(f"id: {target_section_id}")
        print(f"doc_path: {doc_path}")
        print("- - - - - ERROR - - - - -")
        print(f"{client_name} failed. Skipping. Error message:")
        print(e)
        traceback.print_exc()
        print("-------------------------------")

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
        with open(f"need_review\\manual_check_needed.json", 'w') as json_file:
            json.dump(manual_check_needed, json_file)

    #manual_check_needed = {}
    client_name = ""

def createDir(directory):
    directory = re.sub(r'[<>:"|?*]', '', directory)
    directory = directory[:1] + ':' + directory[1:]
    if not os.path.exists(directory):
        os.makedirs(directory)

def convertPages(section, section_dir):
    namespaces = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}
    
    all_pages = section.findall('one:Page', namespaces)
    main_dir = f"{section_dir}"
    sub_dir = f"{section_dir}"
    
    for page in all_pages:
        page_name = page.get("name").rstrip().replace("\\", " or ").replace(" / ", " or ").replace("/", " or ")
        page_id = page.get("ID")
        page_level = page.get("pageLevel")
        
        if page_level == "1":
            main_dir = f"{section_dir}\\{page_name}"
            page_dir = f"{section_dir}\\{page_name}.doc"
            
            if os.path.exists(main_dir) and os.path.exists(page_dir):
                os.remove(page_dir)
                #continue
                page_dir = f"{section_dir}\\{page_name}\\{page_name}.doc"
        elif page_level == "2":
            createDir(main_dir)
            sub_dir = f"{main_dir}\\{page_name}"
            page_dir = f"{main_dir}\\{page_name}.doc"

            if os.path.exists(sub_dir) and os.path.exists(page_dir):
                os.remove(page_dir)
                #continue
                page_dir = f"{main_dir}\\{page_name}\\{page_name}.doc"
                
        elif page_level == "3":
            createDir(sub_dir)
            page_dir = f"{sub_dir}\\{page_name}.doc"

        page_dir = re.sub(r'[<>:"|?*]', '', page_dir)
        page_dir = page_dir[:1] + ':' + page_dir[1:]
        print(page_dir)

        convertToITGlue(page_id, page_name, page_dir)
        
def convertSection(section_group, base_directory):
    namespaces = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}
    
    sections = section_group.findall('one:Section', namespaces)
    for section in sections:
        section_name = section.get("name").rstrip().replace("\\", " or ").replace(" / ", " or ").replace("/", " or ")
        section_id = section.get("ID")
        section_dir = f"{base_directory}\\{section_name}"
        createDir(section_dir)
        
        convertPages(section, section_dir)

def convertSectionGroup(parent, base_directory, getSTA=False):
    namespaces = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}

    section_groups = parent.findall('one:SectionGroup', namespaces)

    system_tools_applications_group = None
    for section_group in section_groups:
        section_group_name = section_group.get("name").rstrip().replace("\\", " or ").replace(" / ", " or ").replace("/", " or ")
        
        section_group_dir = f"{base_directory}\\{section_group_name}"
        createDir(section_group_dir)

        if getSTA and section_group_name == "Systems, Tools, and Applications":
            return section_group
        convertSection(section_group, section_group_dir)
        
def getPlayBookSections():
    current_directory = os.getcwd()
    base_directory = f"{current_directory}\\formatted_docs\\NOC_Playbook"
    createDir(base_directory)
    
    app = win32com.client.dynamic.Dispatch('OneNote.Application')
    open_notebooks = app.GetHierarchy("", 1)

    root = ET.fromstring(open_notebooks)
    namespaces = {'one': 'http://schemas.microsoft.com/office/onenote/2013/onenote'}
    all_notebooks = root.findall('one:Notebook', namespaces)        

    for notebook in all_notebooks:
        if notebook.get('name') == "NOC":
            notebook_id = notebook.get('ID')
            break

    noc_playbook = app.GetHierarchy(notebook_id, 4)
    root = ET.fromstring(noc_playbook)

    #all sections from highest level of playbook hierarchy
    convertSection(root, base_directory)
        
    #all section groups from highest level of playbook hierarchy
    convertSectionGroup(root, base_directory)

    #get "Systems, Tools, and Applications" section group
    system_tools_applications_group = convertSectionGroup(root, base_directory, getSTA=True)

    #if "Systems, Tools, and Applications" section group exists
    if system_tools_applications_group is not None:
        base_directory = f"{base_directory}\\Systems, Tools, and Applications"
        #get "Systems, Tools, and Applications" section group children
        convertSectionGroup(system_tools_applications_group, base_directory)


print(getPlayBookSections())
