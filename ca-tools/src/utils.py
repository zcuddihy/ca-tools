#%%
import requests
import os


def get_pdf(url: str):
    headers = {
        "User-Agent": "Windows 10/ Edge browser: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/42.0.2311.135 Safari/537.36 Edge/12.246"
    }

    res = requests.get(url, headers=headers, stream=True)
    pdf_file = res.content
    file_name = res.headers.get("Content-Disposition").split("filename=")[-1].strip('"')

    return pdf_file, file_name


def save_to_folder(file_name: str, pdf_file: bytes, file_path: str):
    # Create new folder for the given file_path
    try:
        os.makedirs(file_path)
    except:
        print("Folder already exists")

    # Save the PDF file to the specificed folder
    try:
        with open(f"{file_path}/{file_name}.pdf", "wb") as f:
            f.write(pdf_file)
    except:
        print("File already exists")


# %%
