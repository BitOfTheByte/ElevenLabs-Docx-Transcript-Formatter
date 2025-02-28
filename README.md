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

## Usage

1. **Rename Speaker IDs**

Place your .docx and corresponding .json files in the same directory as the scripts. Then run:

```bash
python rename_speakers.py
```

This will generate new DOCX files with the prefix edited_ where the speaker IDs have been replaced with speaker names.

2. **Format Transcripts**

After running the first script, format the edited transcripts by running:

```bash
python format_transcripts.py
```

This will create new DOCX files with the prefix formatted_, where the dialogue blocks are reformatted and each unique speaker name is highlighted in a unique color.
