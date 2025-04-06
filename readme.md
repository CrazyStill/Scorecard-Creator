Scorecard Generator
Scorecard Generator is a web application designed to help you quickly generate professional scorecards. With this tool, you can upload your own templates and data to produce final, print-ready PDFs in just a few steps.

Overview
Scorecard Generator streamlines the process of creating custom scorecards by allowing users to:

Upload Custom Templates: Upload a Word document (.docx) as your front template along with a CSV data template and, optionally, a static back design (PDF).

Map CSV Headers to Placeholders: Map CSV headers to corresponding placeholders in your Word template so that your scorecards are populated with the correct data.

Generate Final Scorecards: Process the data, convert the document to PDF, merge pages if needed, and generate a print-ready PDF.

Manage Templates: Easily delete or update templates as your needs change.

How It Works
Upload a New Template:

Navigate to the "Upload New Template" page.

Enter a sport name and a unique template name.

Upload your Word template (front) in .docx format that includes placeholders for dynamic data.

Upload a CSV file that serves as a data template (must contain column headers).

Optionally, upload a PDF file for a static back design to be merged with the front scorecard.

Mapping CSV Headers to Placeholders:

After uploading your files, the system directs you to a mapping page.

The system extracts the CSV headers and displays them on the page.

For each header, enter the exact placeholder text used in your Word template.
Tip: For multiple scorecards per page, append an underscore and a number (e.g., TeamName_1, TeamName_2).

Specify the number of scorecards to be printed on each page (typically between 1 and 4).

Click "Save Mapping" to store your settings.

Generating Scorecards:

On the "Generate Scorecard" page, download the provided CSV template and fill it with your data using your preferred spreadsheet program.

Upload the completed CSV file.

The system processes your data by:

Grouping data based on the number of scorecards per page.

Replacing placeholders in the Word template with the actual data.

Converting the document to PDF.

Merging multiple pages (if necessary) into a single PDF.

Your final PDF scorecard is generated and automatically downloaded.

Deleting Templates:

To remove an old template, return to the main page and click the "Delete" button next to the template.

This action removes the entire template directory along with all associated files.

Technical Details
DOCX to PDF Conversion:
The application uses the docx2pdf library to convert Word documents to PDF. If this conversion fails, a fallback to COM automation (using pywin32) is provided.

PDF Merging:
Intermediate PDFs are merged using PyPDF2 to create the final scorecard PDF.

Front-End Styling:
The site leverages Bootstrap 5 for responsive design and a consistent user interface.

File Storage:
All uploaded files are stored in a dedicated folder structure under SCTEMP/(SPORT)/(TEMPLATENAME)/, with temporary files being purged after processing.

Frequently Asked Questions (FAQ)
Q: Do I need any special software on my computer?
A: No, all processing is performed on the server. Your computer just needs a web browser.

Q: What file formats are supported?
A: The Word template must be in .docx format. CSV files are used for data input, and optionally, a PDF can be uploaded for the back design.

Q: How do I format placeholders in my Word template?
A: Use clear, unique placeholder text. For multiple scorecards on one page, append an underscore and the scorecard number (e.g., TeamName_1, TeamName_2).

Future Enhancements
Planned improvements include:

User authentication and session management.

Real-time progress updates during scorecard generation.

Enhanced template editing and customization options.

More detailed logging and error reporting.

Contact & Support
If you have any questions or need support, please contact:
aselzer@cityofcape.org

