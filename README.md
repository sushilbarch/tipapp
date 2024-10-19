1. Overview
The application allows users to:

Select and fill out templates.
Generate documents using the specified templates.
Save data into a SQLite database.
Download all submissions in Excel format.
2. Key Functionalities
Database Configuration and Initialization:

The app uses a SQLite database named tippani.db.
A table named submissions is created to store data filled out by users.
The function init_db() creates this table if it does not exist.
Routes:

/: Displays the index page where users can select templates.
/select_template: Handles the template selection, redirecting to a specific template form.
/template/<template>: Displays a form based on the selected template, pre-filling it with previous submission data if available.
/generate: Generates the "Tippani" document by taking user inputs, replacing placeholders in a Word template, and saving the output.
/download_data: Allows users to download all submissions as an Excel file.
