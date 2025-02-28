# ElevenLabs-Docx-Transcript-Formatter

This repository contains two Python scripts that automate the process of formatting transcript documents in `.docx` format. These scripts utilize accompanying `.json` files to map speaker IDs to speaker names and then format the transcripts for improved readability and visual appeal.

## Overview

The repository includes two main scripts:

1. **`rename_speakers.py`**  
   - Scans a folder for `.json` and corresponding `.docx` files.
   - Extracts unique `speaker_id` and their `speaker_name` values from the JSON file.
   - Replaces speaker IDs in the DOCX transcripts with the corresponding speaker names.
   - Saves the modified transcript as a new file prefixed with `edited_`.

2. **`format_transcripts.py`**  
   - Opens the edited DOCX files (with the `edited_` prefix).
   - Reformats each dialogue block to the format:  
     `[SpeakerName]: Dialogue text`
   - Assigns a unique text color to each speaker for better differentiation.
   - Saves the formatted transcript as a new file prefixed with `formatted_`.

## Prerequisites

- **Python 3.x**
- **python-docx**  
  Install using pip:
  ```bash
  pip install python-docx
