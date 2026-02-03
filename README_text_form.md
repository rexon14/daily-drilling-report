# Text Form App

A simple Streamlit application for saving text notes with date-based filenames.

## Features

- 📅 Date picker to select the date
- 📝 Large text area for entering content
- 💾 Save button to save text as .txt files
- 📁 Files are saved with format: `YYYY-MM-DD.txt` (e.g., `2026-01-25.txt`)

## Installation

1. Install Streamlit (if not already installed):
```bash
pip install streamlit
```

## Usage

Run the app with:
```bash
streamlit run text_form.py
```

The app will open in your default web browser.

## File Storage

All saved text files are stored in the `saved_texts` folder within the Zone 5 directory.
