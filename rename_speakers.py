#!/usr/bin/env python3
import os
import json
import re
from docx import Document

def process_files(directory='.'):
    # List all JSON files in the directory
    json_files = [f for f in os.listdir(directory) if f.endswith('.json')]
    
    for json_file in json_files:
        base_name = os.path.splitext(json_file)[0]
        docx_file = os.path.join(directory, base_name + '.docx')
        
        # Check if the corresponding DOCX file exists
        if not os.path.exists(docx_file):
            print(f"Warning: No corresponding DOCX file found for {json_file}. Skipping.")
            continue

        # Load JSON data to build speaker mapping (speaker_id -> speaker_name)
        json_path = os.path.join(directory, json_file)
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        speaker_mapping = {}
        for entry in data.get("words", []):
            speaker_id = entry.get("speaker_id")
            speaker_name = entry.get("speaker_name")
            if speaker_id and speaker_name:
                speaker_mapping[speaker_id] = speaker_name

        # Open the DOCX file
        doc = Document(docx_file)
        
        # Replace all occurrences of [speaker_x] with the corresponding [speaker_name]
        pattern = re.compile(r'\[(speaker_[^]]+)\]')
        for para in doc.paragraphs:
            # Replace using a lambda that looks up the mapping
            new_text = pattern.sub(lambda m: f'[{speaker_mapping.get(m.group(1), m.group(1))}]', para.text)
            para.text = new_text
        
        # Save the edited document with a new filename prefix
        new_docx_filename = os.path.join(directory, f"edited_{base_name}.docx")
        doc.save(new_docx_filename)
        print(f"Processed {docx_file} and saved as {new_docx_filename}")

if __name__ == '__main__':
    process_files()
