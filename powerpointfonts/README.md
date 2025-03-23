# PowerPoint Font Changer

This application allows you to change all fonts in a PowerPoint presentation while preserving font sizes and colors.

## Features

- Load any PowerPoint (.pptx) file
- Select any font installed on your system
- Preview the selected font
- Process the PowerPoint file and create a new version with the changed font
- Preserves all font sizes, colors and other formatting

## Installation

1. Ensure you have Python 3.8 or higher installed
2. Install the required packages:

```
pip install -r requirements.txt
```

## Usage

1. Run the application:

```
python font_changer.py
```

2. Click "Browse..." and select your PowerPoint (.pptx) file
3. Select the new font you want to use from the dropdown menu
4. Use the "Preview Font" button to see how the selected font looks
5. Click "Change Fonts" to process the file
6. A new file will be created with "\_new_font" appended to the original filename

## Note

This application only changes the font family. Font sizes, colors, and other formatting will remain unchanged. The application creates a new file, so your original presentation remains untouched.
