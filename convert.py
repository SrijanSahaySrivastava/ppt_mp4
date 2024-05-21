import win32com.client
import argparse
import time
import os
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse

# Get the current working directory
cwd = os.getcwd()


## CLI-----------------------------------
def convert_ppt_to_mp4_c(ppt_path):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    try:
        # presentation = powerpoint.Presentations.Open(FileName='lol.pptx', WithWindow=False)
        presentation = powerpoint.Presentations.Open(
            FileName=ppt_path, WithWindow=False
        )
    except:
        print("File cannot be found")
        exit

    try:
        output_path = os.path.join(cwd, "output.wmv")
        presentation.CreateVideo(output_path)
        while presentation.CreateVideoStatus == 1:
            time.sleep(1)
        presentation.Close()
        print("Done")
    except:
        print("Unable to export to video")


## API---------------------
app = FastAPI()


@app.post("/convert_ppt_to_mp4/")
async def convert_ppt_to_mp4(ppt_file: UploadFile = File(...)):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")

    temp_file = os.path.join(os.getcwd(), ppt_file.filename)
    with open(temp_file, "wb") as buffer:
        buffer.write(await ppt_file.read())

    try:
        presentation = powerpoint.Presentations.Open(
            FileName=temp_file, WithWindow=False
        )
    except:
        return JSONResponse(
            status_code=400,
            content={"status": "error", "message": "File cannot be found"},
        )

    output_path = os.path.join(os.getcwd(), "output.wmv")

    try:
        presentation.CreateVideo(output_path)
        while presentation.CreateVideoStatus == 1:
            time.sleep(1)
        presentation.Close()
        return {"status": "success", "message": "Done"}
    except:
        return JSONResponse(
            status_code=500,
            content={"status": "error", "message": "Unable to export to video"},
        )


import asyncio

import concurrent.futures


def main():
    parser = argparse.ArgumentParser(description="Convert a PowerPoint file to MP4.")
    parser.add_argument("ppt_path", type=str, help="The path to the PowerPoint file.")

    args = parser.parse_args()

    loop = asyncio.get_event_loop()

    with concurrent.futures.ThreadPoolExecutor() as pool:
        loop.run_in_executor(pool, convert_ppt_to_mp4_c, args.ppt_path)

    loop.close()


if __name__ == "__main__":
    main()
