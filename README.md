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

### Postman testing

- Open Postman.

- Click on the '+' button to create a new tab for a new request.

- From the dropdown menu next to the URL bar, select 'POST'.

- In the URL bar, enter the URL of the endpoint. If you're running the FastAPI server locally, this will be http://localhost:8000/convert_ppt_to_mp4/.

- Below the URL bar, click on the 'Body' tab.

- Select 'form-data'.

- In the 'Key' column, enter ppt_file. In the corresponding 'Value' column, click on 'Select Files' and choose your PowerPoint file.

- In the next row of the 'Key' column, enter times_per_slide. In the corresponding 'Value' column, enter your list of times, formatted as a JSON array (for example, [5,2,3]).

- Click on the 'Send' button to send the request.

## Usage for CLI

Run the CLI application using the following command:

```bash
python convert.py path_to_ppt_file time_for_each_slide
```

example:

```bash
python convert.py path_to_ppt_file 5 2 3
```

This command converts the PowerPoint file to an MP4 video file and saves it in the same directory as the working directory of the script.
