from customtkinter import filedialog
import os
from modules import *

# Keywords to filter specific filenames. "._" for example, as that indicates a corrupted file
file_filter_list: list[str] = [
    "._",
]

# Keywords to filter specific directory names that should always be skipped
directory_filter_list: list[str] = [
    "KFM",
    "_Stammdaten",
]

def main() -> None:
    database_logic.create_all_presentations_table()
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
    
    search_query: str = input("Enter your search query: ")
    query_result = database_logic.query_database(search_query)
    print(query_result)


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
                        if dbcheck == 0:
                            file_paths.append(file_path)
                        else:
                            continue
    return file_paths

# TO-DO: Compare all presentations from the presentations table to the ones on the systems -> update if necessary

# TO-DO: Connect write_pptx module

# TO-DO: GUI

if __name__ == '__main__':
    main()