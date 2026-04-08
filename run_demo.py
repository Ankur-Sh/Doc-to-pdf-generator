#!/usr/bin/env python3
"""Demo script to run the document generator with predefined inputs"""
import os
import sys
import builtins

# Set the current directory as the folder
folder_name = os.path.dirname(os.path.abspath(__file__))
if not folder_name:
    folder_name = "."

# Mock the input function to provide automatic responses
original_input = builtins.input

def mock_input(prompt):
    if "Generate final pdfs" in prompt:
        return "n"  # Don't generate PDFs for demo (faster)
    elif "Enter folder" in prompt:
        return folder_name
    elif "font-size" in prompt:
        return "12"
    elif "line-spacing" in prompt:
        return "1.5"
    elif "Enter link" in prompt:
        return "https://www.example.com/"
    else:
        return original_input(prompt)

# Replace input function
builtins.input = mock_input

# Now import and run the main script
if __name__ == "__main__":
    # Import the main module
    import docx_generator

