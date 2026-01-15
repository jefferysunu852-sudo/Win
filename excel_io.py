from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
import streamlit as st

def load_excel(file, data_only=False) -> Workbook:
    """Loads an Excel file into an openpyxl Workbook."""
    if file is None:
        return None
    try:
        return load_workbook(file, data_only=data_only)
    except Exception as e:
        if st._is_running_with_streamlit:
            st.error(f"Error loading file: {e}")
        else:
            print(f"Error loading file: {e}")
        return None
