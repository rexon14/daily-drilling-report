"""
Simple Text Form App
A minimalist Streamlit app to save notes as .txt files.
"""

import streamlit as st
from datetime import datetime, timedelta
from pathlib import Path
import subprocess
import sys
import re

# Page configuration
st.set_page_config(
    page_title="Daily Drilling Report",
    page_icon="📝",
    layout="centered",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': "Simple Text Form App"
    }
)

# File management setup
base_dir = Path(__file__).parent


def extract_clipboard_data(stdout):
    """
    Extract clipboard data from app.py stdout between CLIPBOARD_DATA_START and CLIPBOARD_DATA_END markers.
    """
    pattern = r'CLIPBOARD_DATA_START\n(.*?)CLIPBOARD_DATA_END'
    match = re.search(pattern, stdout, re.DOTALL)
    if match:
        return match.group(1)
    return None

# Export location options
EXPORT_LOCATIONS = {
    "Region 5": "Region 5",
    "Zone 5": "Zone 5",
    "Zone 6": "Zone 6",
    "Zone 7": "Zone 7"
}

# Sidebar
with st.sidebar:
    st.write("**Export Location**")
    selected_location = st.radio(
        "Select destination:",
        options=list(EXPORT_LOCATIONS.keys()),
        index=3,  # Default to Zone 7
        label_visibility="collapsed"
    )
    
    # Set the saved files directory based on selection
    saved_files_dir = base_dir / EXPORT_LOCATIONS[selected_location] / "daily-report"
    saved_files_dir.mkdir(parents=True, exist_ok=True)
    
    st.markdown("---")
    st.write("**Stats**")
    if saved_files_dir.exists():
        txt_files = list(saved_files_dir.glob("*.txt"))
        st.write(f"Files: {len(txt_files)}")

# Main content
st.title("Daily Drilling Report")

# Date selection
if "selected_date" not in st.session_state:
    st.session_state.selected_date = datetime.now().date()

# Date picker with navigation buttons in a clean row
col1, col2, col3, col4 = st.columns([1, 1, 3, 5])

with col1:
    if st.button("◄", help="Previous day"):
        st.session_state.selected_date -= timedelta(days=1)
        st.experimental_rerun()

with col2:
    if st.button("►", help="Next day"):
        st.session_state.selected_date += timedelta(days=1)
        st.experimental_rerun()

with col3:
    selected_date = st.date_input(
        "Date",
        value=st.session_state.selected_date,
        key="date_picker",
        label_visibility="collapsed"
    )
    st.session_state.selected_date = selected_date

selected_date = st.session_state.selected_date

# Text input
text_content = st.text_area(
    "",
    height=300,
    placeholder="Enter your text here...",
    label_visibility="collapsed"
)

# Save button (left aligned)
save_button = st.button("Save", type="secondary")

# Save functionality
if save_button:
    if not text_content.strip():
        st.warning("Please enter some text before saving.")
    else:
        try:
            date_str = selected_date.strftime("%Y-%m-%d")
            filename = f"{date_str}.txt"
            filepath = saved_files_dir / filename
            
            # Save the text file
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(text_content)
            
            st.success(f"Saved to {selected_location}\\daily-report\\{filename}")
            
            # Run app.py from the selected location folder
            location_folder = base_dir / EXPORT_LOCATIONS[selected_location]
            app_py_path = location_folder / "app.py"
            
            if app_py_path.exists():
                try:
                    # Run app.py with the date as argument
                    result = subprocess.run(
                        [sys.executable, str(app_py_path), date_str],
                        cwd=str(location_folder),
                        capture_output=True,
                        text=True,
                        timeout=60  # 60 second timeout
                    )
                    
                    # Extract and display clipboard data
                    clipboard_data = extract_clipboard_data(result.stdout)
                    if clipboard_data:
                        st.markdown("---")
                        # Create a scrollable container for the code block
                        with st.container():
                            st.markdown("""
                                <style>
                                div[data-testid="stCodeBlock"] {
                                    max-height: 400px;
                                    overflow-y: auto;
                                }
                                div[data-testid="stCodeBlock"] pre {
                                    max-height: 400px;
                                    overflow-y: auto;
                                }
                                </style>
                            """, unsafe_allow_html=True)
                            st.code(clipboard_data, language=None)
                    
                    # Only show errors if they exist
                    if result.returncode != 0:
                        if result.stderr:
                            st.error("Error occurred during processing")
                            st.code(result.stderr, language=None)
                            
                except subprocess.TimeoutExpired:
                    st.error(f"⏱️ app.py execution timed out after 60 seconds")
                except Exception as e:
                    st.error(f"Error running app.py: {str(e)}")
            else:
                st.info(f"ℹ️ app.py not found in {selected_location} folder")
            
        except Exception as e:
            st.error(f"Error: {str(e)}")

# Saved files list
st.markdown("---")
if saved_files_dir.exists():
    txt_files = sorted(saved_files_dir.glob("*.txt"), reverse=True)
    
    if txt_files:
        for txt_file in txt_files:
            file_size = txt_file.stat().st_size
            st.text(f"{txt_file.name} ({file_size:,} bytes)")
    else:
        st.text("No saved files yet.")
