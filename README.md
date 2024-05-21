This project provides a FastAPI application that converts PowerPoint (.ppt) files to MP4 video files.

## Installation

1. Clone this repository.
2. Install the required Python packages using pip:

```bash
pip install -r requirements.txt
```

## Usage for API

Run the FastAPI application using the following command:

```bash
uvicorn convert:app --reload
```

This command starts the FastAPI application on localhost:8000.

Send a POST request to http://localhost:8000/convert_ppt_to_mp4/ with a PowerPoint file in the request body. The file should be sent as form-data with the key ppt_file.

## Usage for CLI

Run the CLI application using the following command:

```bash
python convert.py path/to/ppt_file.ppt
```

This command converts the PowerPoint file to an MP4 video file and saves it in the same directory as the working directory of the script.
