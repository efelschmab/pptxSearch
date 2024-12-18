import os
import sqlite3
from customtkinter import filedialog
import customtkinter as ctk
from modules import *

# command for bulding the executable: pyinstaller --onefile --noconsole --add-data "template.pptx;./modules" main.py

# Filter list for potentially damaged files
file_filter_list: list[str] = [
    "._",
    "~$",
]

# Directories to skip
directory_filter_list: list[str] = []

ui_obj = ui.GUI()


def main() -> None:
    database_logic.create_all_presentations_table()

    window = ctk.CTk()
    window.geometry("600x230")
    window.title("pptxSearch Tool")
    window._fg_color = "#242424"
    window.columnconfigure(0, minsize=100)
    window.columnconfigure(1, weight=1)

    top_frame = ui_obj.container(master_frame=window, row=0)
    create_db_button = ui_obj.button(master_frame=top_frame, column=0,
                                     row=0, button_text="Create Database", button_command=create_database)
    db_info_label = ui_obj.db_label(master_frame=top_frame, column=1, text="")

    mid_frame = ui_obj.container(master_frame=window, row=1)
    search_button = ui_obj.button(
        master_frame=mid_frame, column=0, row=1, button_text="Search", button_command=search)
    query_entry = ui_obj.entry(master_frame=mid_frame, column=1, row=1)

    log_label = ui_obj.loglabel(master_frame=window, column=0, text="")

    # Check database content after initialization of the UI
    ui_obj.change_label_text(text=database_logic.get_database_content())

    window.mainloop()


def search() -> None:

    search_query: str = ui_obj.fetch_query_field()

    try:
        query_result = database_logic.query_database(search_query)
        ui_obj.change_log_label_text(text="")
    except sqlite3.OperationalError as e:
        ui_obj.change_log_label_text(
            text="No query result or invalid search query.")
        print(f"An error occurred: {e}")
        return None

    if not query_result:
        ui_obj.change_log_label_text(text="No query result.")
        return None

    write_pptx.create_presentation(query_result)
    os.startfile("search_results.pptx")


def create_database() -> None:

    directory_path: str = filedialog.askdirectory()
    ui_obj.change_log_label_text(text="Processing found presentations")
    found_presentations: list = find_pptx_files(directory_path)

    for presentation in found_presentations:
        presentation_name, presentation_path, slide_number, hidden_slides, presentation_created, presentation_modified, all_text = read_pptx.get_presentation(
            presentation)
        if presentation_name:
            # Processing search results -> database
            database_logic.write_to_database(word_list=all_text,
                                             presentation_name=presentation_name,
                                             presentation_path=presentation_path)
            database_logic.save_pptx_to_database(name=presentation_name,
                                                 path=presentation_path,
                                                 created=presentation_created,
                                                 modified=presentation_modified,
                                                 slides=slide_number,
                                                 hidden=hidden_slides)
    ui_obj.change_log_label_text(text="")
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
                        file_modified_time: str = read_pptx.get_time(
                            1, file_path)
                        dbcheck: int = database_logic.check_for_existing(
                            file, file_modified_time)
                        if dbcheck == 0 and file_path not in file_paths:
                            file_paths.append(file_path)
                        else:
                            continue
    return file_paths


if __name__ == '__main__':
    main()
