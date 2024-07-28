# %%
# pip install PyPDF2
from pathlib import Path
from PyPDF2 import PdfMerger, PdfReader

PathC = "C:/OTMX"
NameItem = "2024-042"

# %%
# Define input directory for the pdf files
#pdf_dir = Path(__file__).parent / "Input"
pdf_dir = Path(PathC) / "Inputs" / NameItem


# %%
# Define & create output directory
#pdf_output_dir = Path(__file__).parent / "OUPUT"
#pdf_output_dir.mkdir(parents=True, exist_ok=True)
pdf_output_dir = Path(PathC) / "Outputs" / NameItem
pdf_output_dir.mkdir(parents=True, exist_ok=True)


# %%
# List all pdf files in the input directory
pdf_files = list(pdf_dir.glob("*.pdf"))

# %%
# Use the first 10 characters as the 'key'
#00000SCS22_CERT_CABALLEROGONZALEZOMAR_GMM-20829.pdf
keys = set([file.name[:10] for file in pdf_files])
BASE_FILE_NAME_LENGTH = 10

# %%
# Determine the file name length of the base file
for key in keys:
    merger = PdfMerger()
    for file in pdf_files:
        if file.name.startswith(key):
            merger.append(PdfReader(str(file), "rb"))
            if len(file.name) >= BASE_FILE_NAME_LENGTH:
                base_file_name = file.name
    merger.write(str(pdf_output_dir / base_file_name))
    merger.close()

# %%
#pyinstaller -F -w -i c-add.ico MergePDF.py


