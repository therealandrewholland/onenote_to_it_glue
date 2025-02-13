import OneNote_to_IT_Glue
import os
from urllib.parse import unquote

def initDir():
    dir_list = ["formatted_docs", "temp", "need_review", "temp\\extracted_contents", "temp\\final_docx"]

    for directory in dir_list:
        if not os.path.exists(directory):
            os.makedirs(directory)        

def checkInput(user_input):
    # Check if the input starts with "onenote:" or ends with ".one"
    if user_input.startswith("onenote:") or user_input.startswith("https:") or user_input.endswith(".one"):
        return True
    return False

def cleanOneNoteURL(url):
    # Remove "/:o:/r" from the URL
    cleaned_url = url.replace("/:o:/r", "")
    
    # Split using "?d=" and remove everything after
    if "?d=" in cleaned_url:
        cleaned_url = cleaned_url.split("?d=")[0]

    return cleaned_url

def main():
    initDir()
    
    while True:
        print("OneNote to IT Glue conversion tool")
        user_input = input("Please enter a OneNote URL or path:")

        if user_input == 'exit' or user_input == 'quit':
            break
        elif checkInput(user_input):
            if user_input.startswith("https:"):
                user_input = "onenote:" + user_input

            client_name = None
            if user_input.startswith("onenote:"):
                user_input = cleanOneNoteURL(user_input)
                try:
                    client_name = unquote(user_input[8:].split("/")[7])
                except:
                    pass
                print("Valid URL!")
            else:
                try:
                    client_name = path_or_url.split("\\")[5]
                except:
                    pass
                print("Valid path!")

            is_client_name = input(f"Is \"{client_name}\" the client? (y/n):")
            if is_client_name == "n":
                client_name = input("Please enter client name:")

            print("Converting OneNote...")
            onenote_to_it_glue.convertToITGlue(user_input, client_name)
        else:
            print("Invalid URL or path!")

if __name__ == "__main__":
    main()
