# pptxSearch (cs50x final project)

#### Video Demo: <https://youtu.be/wivyl5saE80>

#### Short description: A desktop search for PowerPoint presentations

<hr>

### Long description

While it's easy to search for keywords within a single PowerPoint file, searching for specific keywords across multiple presentations can be more challenging. This simple tool utilizes SQLite3 to create a database of keywords and their associated presentations, providing a convenient and accessible way to search through them all at once.

##### Here is how it works

- Launch the executable (or run the program from your code editor).
- Click the "Create Database" button (or if a database already exists, the text next to the button will reflect that).
- Select a folder to process and add to the database.
- The program will then scan the selected folder and its subdirectories for .pptx files and populate the database.
- Enter a search query to retrieve files and slide numbers containing the specified word.
- A PowerPoint presentation will be generated and opened, displaying a list of search results.
- Clicking on a result will open the corresponding file at the specified slide where the result was found.

##### Why does processing all those PowerPoint files take so long?

Each discovered presentation undergoes several initial checks. The program verifies that the file isn't already in the database (based on filename and last modification time) and filters out potentially corrupted files (like those starting with "._" or "~$").

The Python code includes a placeholder for a list of directories to exclude, allowing you to specify recurring directories that should be ignored.

The content of each presentation is extracted, stripped of unnecessary elements (like words shorter than 3 characters or purely numerical), and prepended with an underscore to ensure compatibility with SQLite's naming conventions.

Several exceptions are handled to ensure smooth processing. For instance, password-protected or damaged files are skipped.

Once all directories are processed and the database button's label is updated, you can begin searching the database.
