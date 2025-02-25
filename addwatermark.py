import os
import subprocess
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO


def add_watermark_to_pdf(input_pdf_path, output_pdf_path, watermark_text):
    """Add a watermark to a PDF document."""
    # Create a PDF with the watermark text
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont("Helvetica", 80)
    can.setFillColorRGB(0.5, 0.5, 0.5, alpha=0.5)  # Light gray with transparency
    can.saveState()
    can.translate(100, 200)  # Adjust position as needed
    can.rotate(45)  # Rotate the text diagonally
    can.drawString(0, 0, watermark_text)
    can.restoreState()
    can.save()

    # Move to the beginning of the StringIO buffer
    packet.seek(0)

    # Read the existing PDF
    existing_pdf = PdfReader(input_pdf_path)
    watermark_pdf = PdfReader(packet)
    watermark_page = watermark_pdf.pages[0]

    # Add the watermark to each page
    output = PdfWriter()
    for page in existing_pdf.pages:
        page.merge_page(watermark_page)
        output.add_page(page)

    # Write the output PDF
    with open(output_pdf_path, "wb") as output_pdf:
        output.write(output_pdf)


def convert_docx_to_pdf(input_docx_path, output_pdf_path):
    """Convert a Word document to PDF using LibreOffice."""
    output_folder = os.path.dirname(output_pdf_path)
    subprocess.run([
        "soffice", "--headless", "--convert-to", "pdf",
        input_docx_path, "--outdir", output_folder
    ])
    # Rename the converted file to the desired output path
    base_name = os.path.splitext(os.path.basename(input_docx_path))[0]
    temp_pdf_path = os.path.join(output_folder, f"{base_name}.pdf")
    os.rename(temp_pdf_path, output_pdf_path)


def process_documents_in_folder(folder_path, watermark_text):
    """Process all .docx files in the specified folder and add a watermark."""
    for filename in os.listdir(folder_path):
        # Skip temporary files (starting with ~$)
        if filename.startswith("~$"):
            continue

        if filename.endswith(".docx"):
            file_path = os.path.join(folder_path, filename)
            print(f"Processing: {file_path}")

            # Convert the Word document to PDF using LibreOffice
            pdf_path = os.path.join(folder_path, f"temp_{filename[:-5]}.pdf")
            convert_docx_to_pdf(file_path, pdf_path)

            # Add the watermark to the PDF
            watermarked_pdf_path = os.path.join(folder_path, f"{filename[:-5]}.pdf")
            add_watermark_to_pdf(pdf_path, watermarked_pdf_path, watermark_text)

            # Clean up the temporary PDF
            os.remove(pdf_path)
            print(f"Saved watermarked document as: {watermarked_pdf_path}")


if __name__ == "__main__":
    folder_path = r"F:\projects\py-scripts\COMPUTING\DIBELIT"
    watermark_text = "LIBRARY COPY"

    if os.path.exists(folder_path):
        process_documents_in_folder(folder_path, watermark_text)
    else:
        print(f"The folder {folder_path} does not exist.")