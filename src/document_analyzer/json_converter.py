import re
import openai
import time
import json
from docx import Document 
import streamlit as st



def extract_text_from_word(docx_file,doc=""):
    """
    Extracts text from a Word document without any additional cleaning or processing.
    """
    try:
        if doc=="":
            doc = Document(docx_file)
            
        else :
            doc=doc

        result = ""
        for block in doc.element.body:
            if block.tag.endswith('p'):
                # Handle paragraphs manually (like in your original logic)
                paragraph = block
                text = ''.join(node.text for node in paragraph.xpath('.//w:t') if node.text)
                if text.strip():
                    result += text.strip() + "\n"

            elif block.tag.endswith('tbl'):
                # Convert XML table to docx.table.Table
                for tbl in doc.tables:
                    if tbl._element == block:
                        table_data = []
                        for row in tbl.rows:
                            row_data = {}
                            for i, cell in enumerate(row.cells):
                                row_data[f"col_{i}"] = cell.text.strip()
                            table_data.append(row_data)

                        table_json = json.dumps(table_data, ensure_ascii=False)
                        result += table_json + "\n"
                        break  # Only match once to avoid duplicates

        return result.strip()

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

def key_stucture(json_text):
    merged_data = {}
    previous_key=""
    current_key = None
    current_value = ""
    for key, value in json_text.items():
        # Extract base key by removing digits and hyphens
        base_key = re.sub(r'-\d+', '', key)
        
        if current_key is None:  # Start with the first key
            
            previous_key=key
            current_key = base_key
            current_value = value
            
        elif base_key == current_key:  # Continue merging if the key matches
            current_value += f"\n\n{value}"
        else:  # Different key encountered, save the current group and start a new one
            merged_data[previous_key] = current_value
            current_key = base_key
            current_value = str(value)
            previous_key=key

    # Add the last merged group to the dictionary
    if current_key:
        merged_data[key] = current_value
        
    return merged_data



