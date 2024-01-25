```markdown
# URL to HTML and Word Converter

This utility script provides a GUI for converting URLs listed in a CSV file into HTML and Word documents. The script processes each URL, extracts content using readability algorithms, and saves the content as both HTML and Word documents.

## Features

- GUI for easy interaction
- Deduplication of URLs to avoid processing duplicates
- Conversion of HTML content to Word documents
- Handling of images, lists, tables, and math elements in the conversion process
- Saving of additional metadata (URL) in the generated files

## Requirements

- Python 3.x
- tkinter
- readability-lxml
- BeautifulSoup4
- python-docx
- Pillow
- appscript (for macOS)
- Docker (for running singlefile)

## Installation

Before running the script, ensure all the required Python libraries are installed:

```bash
pip install tkinter readability-lxml beautifulsoup4 python-docx Pillow appscript
```

## Usage

1. Open the script in a Python environment.
2. Run the script to open the GUI.
3. Click on "Browse CSV File" to select a CSV file containing URLs.
4. Choose a directory to save the processed HTML and Word files.
5. The script will process each URL and save the corresponding files in the chosen directory.

## Functions

### `is_duplicate(url)`

Checks if the URL has already been processed to avoid duplication.

### `generate_filename(url)`

Generates a filename for the HTML and Word documents based on the URL structure.

### `extract_main_content(html_content)`

Extracts the main content from the HTML using readability algorithms.

### `extract_styles(html_content)`

Extracts CSS styles from the HTML content.

### `extract_images(html_content)`

Extracts image elements from the HTML content.

### `remove_iframes(html_content)`

Removes iframe elements from the HTML content.

### `save_html(url, save_path)`

Processes a URL, extracts content, and saves it as an HTML file.

### `convert_base64_to_rgb(base64_data)`

Decodes base64 image data and converts it to RGB.

### `save_image_as_jpeg(image_rgb, output_path)`

Saves an RGB image as a JPEG file.

### `handle_list_item(paragraph, element)`

Processes and adds list items to a Word document.

### `handle_table(doc, element)`

Processes and adds tables to a Word document.

### `handle_figure(doc, element, image_counter)`

Processes and adds figures with images to a Word document.

### `handle_math(doc, element)`

Processes and adds math elements to a Word document.

### `html_to_word(input_path, output_path)`

Converts an HTML file to a Word document, preserving styles and content structure.

### `process_csv(csv_file, save_path)`

Reads a CSV file containing URLs and processes each URL to generate HTML and Word documents.

### GUI Elements

- A button to browse and select a CSV file.
- A status label to show the current operation status.

## Notes

- The script uses Docker to run the `singlefile` command for HTML processing.
- The appscript library is used to add comments to files on macOS.
- Temporary files are created during image processing and are removed after use.
- The script assumes CSV files contain URLs in the first column.

## Contributions

Contributions to improve the script or extend its functionality are welcome. Please ensure that you follow the existing code structure and add comments where necessary.

## License

This script is released under the MIT License.
```
