from customtkinter import filedialog
import  customtkinter as ctk
import os
from modules import *

# Keywords to filter specific filenames. "._" for example, as that indicates a corrupted file
file_filter_list: list[str] = [
    "._",
    "~$",
]

# Keywords to filter specific directory names that should always be skipped
directory_filter_list: list[str] = []

# Initialize UI object from UI class
ui_obj = ui.gui()

def main() -> None:
    database_logic.create_all_presentations_table()
    
    window = ctk.CTk()
    window.geometry("600x270")
    window.title("pptxSearch Tool")    
    
    window.columnconfigure(0, minsize=100)
    window.columnconfigure(1, weight=1)

    top_frame = ui_obj.container(master_frame=window, row=0)
    create_db_button = ui_obj.button(master_frame=top_frame, column=0, row=0, button_text="Create Database", button_command=create_database)
    db_info_label = ui_obj.label(master_frame=top_frame, column=1, text="")
    
    bottom_frame = ui_obj.container(master_frame=window, row=1)
    search_button = ui_obj.button(master_frame=bottom_frame, column=0, row=1, button_text="Search", button_command=search)
    query_entry = ui_obj.entry(master_frame=bottom_frame, column=1, row= 1)  
    
    # Check database content after initialization of the UI
    ui_obj.change_label_text(text=database_logic.get_database_content())

    window.mainloop()


def search() -> None:
    search_query: str = ui_obj.fetch_query_field()
    query_result = database_logic.query_database(search_query)
    write_pptx.create_presentation(query_result)
    os.startfile("search_results.pptx") 


def create_database() -> None:
    directory_path: str = filedialog.askdirectory()
    found_presentations = find_pptx_files(directory_path)
    for presentation in found_presentations:
        # All seven return values from the get_presentation function put into variables
        presentation_name, presentation_path, slide_number, hidden_slides, presentation_created, presentation_modified, all_text = read_pptx.get_presentation(presentation)
        if presentation_name:
            # Insert the extracted words into the database
            database_logic.write_to_database(word_list=all_text, presentation_name=presentation_name, presentation_path=presentation_path)
            # Insert the presentation metadata into the database
            database_logic.save_pptx_to_database(name=presentation_name, path=presentation_path, created=presentation_created, modified=presentation_modified, slides=slide_number, hidden=hidden_slides)
    ui_obj.change_label_text(text=database_logic.get_database_content())


def find_pptx_files(directory) -> list[str]:
    """Check the selected directory and all sub directories for .pptx files"""
    file_paths = []
    for root, dirs, files in os.walk(directory):
        for dir_filter in directory_filter_list:
            # Use filter list for directory names
            if dir_filter in dirs:
                dirs.remove(dir_filter)
        for file in files:
            if file.endswith(".pptx"):
                # Use filter list for file names
                for file_filter in file_filter_list:
                    if file.startswith(file_filter):
                        continue
                    else:
                        file_path = os.path.join(root, file)
                        file_modified_time:str = read_pptx.get_time(1, file_path)
                        dbcheck: int = database_logic.check_for_existing(file, file_modified_time)
                        if dbcheck == 0 and file_path not in file_paths:
                            file_paths.append(file_path)
                        else:
                            continue
    print(f"{file_paths = }")
    return file_paths

# TO-DO: GUI

# TO-DO: Search refinement (what if no query? What if query shorter than 1 character?)

if __name__ == '__main__':
    main()