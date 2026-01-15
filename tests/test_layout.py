import pytest
import sys
import os

# Add parent dir to path to import modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_layout import normalize_string, parse_number, extract_week_number

def test_normalize_string():
    assert normalize_string("  Description  of   Work ") == "description of work"
    assert normalize_string("WK 8") == "wk 8"
    assert normalize_string("Man/Hour") == "manhour"
    assert normalize_string("q-ty") == "qty"

def test_parse_number():
    assert parse_number(100) == 100.0
    assert parse_number(1.5) == 1.5
    assert parse_number("1,5") == 1.5
    assert parse_number("1 234,50") == 1234.5
    assert parse_number("  500  ") == 500.0
    assert parse_number(None) is None
    assert parse_number("") is None
    assert parse_number("abc") is None

def test_extract_week_number():
    assert extract_week_number("WK8") == 8
    assert extract_week_number("Weekly Report No 2") == 2
    assert extract_week_number("Year 2024") == 2024 # unintended but strictly correct based on regex
    assert extract_week_number("No numbers here") is None
