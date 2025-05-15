# Website Screenshot Automation Script

This Python script automates the process of capturing scrolling screenshots of a webpage and compiles them into a Microsoft Word document. It uses `Pillow` for image comparison, and `python-docx` for creating the Word file.

---

## Features

- Automatically scrolls through a webpage
- Captures cropped screenshots of each visible viewport
- Compares screenshots to detect when the bottom of the page is reached
- Saves all screenshots into a formatted `.docx` file
- Deletes temporary screenshot images after saving

---

## Output

- A Word document named **`Website_Document_Capture.docx`** containing each captured viewport as an image, labeled by page number.

---

##  Requirements

Make sure to install the following Python libraries:

\\ bash
pip install pyautogui python-docx Pillow


##  Usage

Set your browser window and content at the correct position.

---

## Run the script using:

bash
Copy
Edit
python script.py
The script waits 5 seconds so you can focus your browser window.

It then begins taking and comparing screenshots until it detects the end of the page.

Your .docx file will be saved in the specified output directory.

---

##  Configuration

Modify these settings in the script as per your screen resolution:

python
Copy
Edit
crop_left = 120
crop_top = 150
crop_width = 1130
crop_height = 510
These control the screenshot cropping area for each page capture.

---

## Cleanup

Temporary screenshot images are automatically deleted after the Word file is saved.

---

## Credits

Developed by Steven Josh to simplify documentation of long-scrollable web pages.

---

