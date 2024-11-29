import sqlite3
from datetime import datetime

# Get current time for timestamps
now = datetime.now()
current_time: str = now.strftime('%d.%m.%Y %H:%M:%S')

database: str = "pptx_search.db"


def check_for_existing(file: str, time: str) -> int:
    """Returns 0 if not found in database, otherwise 1"""
    set_database_encoding("pptx_search.db")
    conn = sqlite3.connect(database, detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    sql_request = f"SELECT * FROM all_presentations WHERE presentation_name = '{
        file}' AND last_modified = '{time}'"
    cursor.execute(sql_request)
    result = cursor.fetchone()

    if not result:
        conn.close()
        return 0
    conn.close()
    return 1


def create_all_presentations_table() -> None:
    set_database_encoding("pptx_search.db")
    conn = sqlite3.connect(database, detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    sql_build_pptx_table: str = "CREATE TABLE IF NOT EXISTS 'all_presentations' (presentation_name TEXT, presentation_path TEXT, created TEXT, last_modified TEXT, last_indexed TEXT, total_slides INTEGER, hidden_slides INTEGER)"
    cursor.execute(sql_build_pptx_table)

    conn.commit()
    conn.close()


def write_to_database(word_list, presentation_name: str, presentation_path) -> None:
    set_database_encoding("pptx_search.db")
    conn = sqlite3.connect(database, detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    for word in word_list:
        # Add a "_" in front of each word for sqlite to be able to handle numbers as first character
        current_word: str = "_" + word['word']
        current_slide_number: int = word["slide"]

        # Check if the table for the word already exists. If so then insert a new row, if not then create it first and then insert
        cursor.execute(
            f"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='{current_word}'")
        table_result = cursor.fetchone()
        table_exists = table_result[0] > 0

        sql_insert: str = f"INSERT INTO '{current_word}' (presentation_path, presentation_name, slide_number) VALUES ('{
            presentation_path}', '{presentation_name}', '{current_slide_number}')"

        if table_exists:
            cursor.execute(sql_insert)
        else:
            sql_build_table: str = f"CREATE TABLE IF NOT EXISTS '{
                current_word}' (presentation_path TEXT, presentation_name TEXT, slide_number INTEGER)"
            cursor.execute(sql_build_table)
            cursor.execute(sql_insert)

        # Remove duplicate database entries
        duplicates: str = f"DELETE FROM '{current_word}' WHERE ROWID NOT IN (SELECT MIN(ROWID) FROM '{
            current_word}' GROUP BY presentation_name, slide_number)"
        cursor.execute(duplicates)

    conn.commit()
    conn.close()


def save_pptx_to_database(name, path, created, modified, slides, hidden) -> None:
    set_database_encoding("pptx_search.db")
    conn = sqlite3.connect(database, detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    sql_insert_presentation: str = f"INSERT INTO 'all_presentations' (presentation_name, presentation_path, created, last_modified, last_indexed, total_slides, hidden_slides) VALUES ('{
        name}', '{path}', '{created}', '{modified}', '{current_time}', '{slides}', '{hidden}')"
    cursor.execute(sql_insert_presentation)

    conn.commit()
    conn.close()


def query_database(query: str) -> list:
    set_database_encoding("pptx_search.db")
    conn = sqlite3.connect(database, detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    sql_search_tables: str = "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE ?"
    cursor.execute(sql_search_tables, ('%_' + query + '%',))

    table_names = cursor.fetchall()
    fetch_list: list = []

    for table_name in table_names:
        result_word: str = table_name[0][1:]
        # Initiate a list for each word to go into the full list of search results
        word_list: list[str] = []
        word_list.append(result_word)

        sql_search_data: str = f"SELECT presentation_path, presentation_name, slide_number FROM {
            table_name[0]}"
        cursor.execute(sql_search_data)

        results = cursor.fetchall()
        for row in results:
            # Add the individual search results (path, name and slide number)
            word_list.append(row)
        fetch_list.append(word_list)

    conn.close()
    return fetch_list


def get_database_content() -> str:
    set_database_encoding("pptx_search.db")
    conn = sqlite3.connect(database, detect_types=sqlite3.PARSE_DECLTYPES)
    cursor = conn.cursor()

    presentation_number = cursor.execute(
        "SELECT COUNT(*) FROM all_presentations").fetchone()[0]

    string: str = ""

    if presentation_number > 0:
        slide_number = "{:,}".format(cursor.execute(
            "SELECT SUM(total_slides) FROM all_presentations").fetchone()[0])
        hidden_slides = "{:,}".format(cursor.execute(
            "SELECT SUM(hidden_slides) FROM all_presentations").fetchone()[0])
        string = f"Database entries found. {presentation_number} presentations with a total of {
            slide_number} slides, {hidden_slides} of which are hidden."

    else:
        string = "No database entries found."

    conn.close()
    return string


def set_database_encoding(database_file) -> None:
    conn = sqlite3.connect(database_file)
    cursor = conn.cursor()

    # Set the database encoding to UTF-8
    cursor.execute("PRAGMA encoding = 'UTF-8'")

    conn.commit()
    conn.close()
