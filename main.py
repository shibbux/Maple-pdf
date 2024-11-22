import fitz  # PyMuPDF
from pdf2docx import Converter
from pypdf import PdfWriter
from gtts import gTTS  # type: ignore
from PIL import Image, ImageTk
import os
import json
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# Utility: Select a single file
def open_file_dialog():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])


# Utility: Select multiple files
def open_multiple_files_dialog():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF Files", "*.pdf")])


# Utility: Select output folder
def select_output_folder():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Output Folder")


# MP3 Conversion from PDF
def convert_pdf_to_mp3():
    file_path = open_file_dialog()
    if not file_path:
        messagebox.showwarning("Warning", "No PDF file selected!")
        return

    output_path = filedialog.asksaveasfilename(
        title="Save MP3",
        defaultextension=".mp3",
        filetypes=[("MP3 Files", "*.mp3")]
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    try:
        doc = fitz.open(file_path)
        text_content = ""
        for page in doc:
            text_content += page.get_text() + "\n\n"

        tts = gTTS(text_content, lang='en')
        tts.save(output_path)
        messagebox.showinfo("Success", "PDF converted to MP3 successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert PDF to MP3. Debug Info: {str(e)}")


# PDF Merge Functionality
def PDFmerge():
    pdfs = open_multiple_files_dialog()
    if pdfs:
        output = filedialog.asksaveasfilename(title="Save Merged PDF As", defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output:
            try:
                writer = PdfWriter()
                for pdf in pdfs:
                    with open(pdf, "rb") as f:
                        writer.append(f)
                with open(output, "wb") as f:
                    writer.write(f)
                messagebox.showinfo("Success", "PDFs merged successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to merge PDFs. Debug Info: {str(e)}")


# General PDF Conversion
def convert_pdf_to(format):
    file_path = open_file_dialog()
    if not file_path:
        messagebox.showwarning("Warning", "No PDF file selected!")
        return

    output_path = filedialog.asksaveasfilename(
        title=f"Save PDF as {format.upper()}",
        defaultextension=f".{format}",
        filetypes=[(f"{format.upper()} Files", f"*.{format}")]
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    try:
        doc = fitz.open(file_path)

        # Handle different formats
        if format == "txt":
            with open(output_path, "w", encoding="utf-8") as f:
                for page in doc:
                    f.write(page.get_text())
        elif format == "html":
            with open(output_path, "w", encoding="utf-8") as f:
                for page in doc:
                    f.write(page.get_text("html"))
        elif format == "xml":
            with open(output_path, "w", encoding="utf-8") as f:
                for page in doc:
                    f.write(page.get_text("xml"))
        elif format == "json":
            pages_data = [{"page": i + 1, "content": page.get_text()} for i, page in enumerate(doc)]
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(pages_data, f, indent=4)
        elif format == "md":
            with open(output_path, "w", encoding="utf-8") as f:
                for page in doc:
                    f.write(page.get_text())
                    f.write("\n---\n")
        elif format == "csv":
            with open(output_path, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Page", "Content"])
                for i, page in enumerate(doc):
                    writer.writerow([i + 1, page.get_text()])
        elif format in ["png", "jpg", "gif", "svg", "bmp", "tiff"]:
            output_folder = os.path.dirname(output_path)
            for page_num in range(len(doc)):
                pix = doc[page_num].get_pixmap()
                page_image_path = os.path.join(output_folder, f"page_{page_num + 1}.{format}")
                pix.save(page_image_path)
        elif format == "epub":
            with open(output_path, "w", encoding="utf-8") as f:
                f.write("<html><body>")
                for page in doc:
                    f.write(f"<p>{page.get_text()}</p><hr>")
                f.write("</body></html>")
        else:
            messagebox.showerror("Error", f"Unsupported format: {format.upper()}")

        messagebox.showinfo("Success", f"PDF converted to {format.upper()} successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert PDF to {format.upper()}. Debug Info: {str(e)}")


def convert_image_to_pdf():
    file_paths = filedialog.askopenfilenames(title="Select Image Files", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tiff")])
    if not file_paths:
        messagebox.showwarning("Warning", "No image file selected!")
        return
    
    # Ask user for the output PDF file location
    output_path = filedialog.asksaveasfilename(
        title="Save PDF As",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )
    if not output_path:
        messagebox.showwarning("Warning", "No output file selected!")
        return

    try:
        # Open the first image to start the PDF
        image = Image.open(file_paths[0])
        # Convert the rest of the images (if any) to RGB mode and append
        image_list = [Image.open(img).convert("RGB") for img in file_paths[1:]]

        # Save the images as a PDF
        image.save(output_path, save_all=True, append_images=image_list)

        messagebox.showinfo("Success", "Images successfully converted to PDF!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert images to PDF. Debug Info: {str(e)}")


# Main GUI
def main():
    root = tk.Tk()

    # Set the window title and size
    root.title("Maple pdf")
    root.geometry("1000x800")
    root.configure(bg="#BF2EF0")

    try:
        # Use PIL to open a PNG image
        icon_image = Image.open("logo.png")  # Replace with your image path
        icon_image = icon_image.resize((32, 32))  # Resize the image if needed
        photo_icon = ImageTk.PhotoImage(icon_image)
        root.iconphoto(True, photo_icon)  # Set the image as window icon
    except Exception as e:
        print(f"Failed to set window icon: {e}")

    # Title Label
    title_label = tk.Label(root, text="Maple pdf", font=("Forte", 90, "bold"), bg="#BF2EF0", fg="#FEECB3")
    title_label.pack(pady=10)

    frame = tk.Frame(root, padx=20, pady=20, bg="#BF2EF0")
    frame.pack(pady=20)

    # Convert Image to PDF Section
    image_to_pdf_label = tk.Label(frame, text="Convert Image to PDF", font=("Felix Titling", 14, "bold"), bg="#BF2EF0", fg="#FFECD1")
    image_to_pdf_label.grid(row=0, column=0, columnspan=3, pady=10)

    ttk.Button(frame, text="Convert Images to PDF", width=20, command=convert_image_to_pdf).grid(row=1, column=0, columnspan=3, pady=10)

    separator = ttk.Separator(frame, orient='horizontal')
    separator.grid(row=2, column=0, columnspan=3, sticky="ew", pady=10)

    # Merge PDFs Section
    merge_label = tk.Label(frame, text="Merge PDFs", font=("Felix Titling", 14, "bold"), bg="#BF2EF0", fg="#FFECD1")
    merge_label.grid(row=3, column=0, columnspan=3, pady=10)

    ttk.Button(frame, text="Merge PDFs", width=30, command=PDFmerge).grid(row=4, column=0, columnspan=3, pady=10)

    separator2 = ttk.Separator(frame, orient='horizontal')
    separator2.grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)

    # Convert PDF Section
    convert_label = tk.Label(frame, text="Convert PDF", font=("Felix Titling", 14, "bold"), bg="#BF2EF0", fg="#FFECD1")
    convert_label.grid(row=6, column=0, columnspan=3, pady=10)

    formats = [
        ("Convert to DOCX", lambda: convert_pdf_to("docx")),
        ("Convert to TXT", lambda: convert_pdf_to("txt")),
        ("Convert to HTML", lambda: convert_pdf_to("html")),
        ("Convert to XML", lambda: convert_pdf_to("xml")),
        ("Convert to JSON", lambda: convert_pdf_to("json")),
        ("Convert to MD", lambda: convert_pdf_to("md")),
        ("Convert to CSV", lambda: convert_pdf_to("csv")),
        ("Convert to PNG", lambda: convert_pdf_to("png")),
        ("Convert to JPG", lambda: convert_pdf_to("jpg")),
        ("Convert to EPUB", lambda: convert_pdf_to("epub")),
        ("Convert to MP3", convert_pdf_to_mp3)
    ]

    # Create buttons for each format in a horizontal manner
    row = 7  # Start from row 7
    col = 0  # Start from column 0

    for i, (text, command) in enumerate(formats):
        ttk.Button(frame, text=text, width=30, command=command).grid(row=row, column=col, padx=5, pady=5)
        col += 1
        # Move to the next row when 3 buttons are placed
        if col > 2:
            col = 0
            row += 1

    footer_label = tk.Label(root, text="Made with love by shibbux", font=("Arial", 20), bg="#BF2EF0", fg="#FFECD1")
    footer_label.pack(side="bottom", pady=1)

    git_prom = tk.Label(root, text="Follow (shibbux) on github ", font=("Arial", 30), bg="#BF2EF0", fg="#FFECD1")
    git_prom.pack(side="bottom", pady=1)

    root.mainloop()


# Run the main function
if __name__ == "__main__":
    main()
