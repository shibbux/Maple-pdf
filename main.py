import fitz  # PyMuPDF
from pdf2docx import Converter
from pypdf import PdfWriter
from gtts import gTTS # type: ignore
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


# Main GUI
def main():
    root = tk.Tk()

    # Set the window title and size
    root.title("Maple pdf")
    root.geometry("750x750")
    root.configure(bg="#FF7733")

    # Set custom image as window icon (PNG to ICO conversion if needed)
    try:
        # Use PIL to open a PNG image
        icon_image = Image.open("logo.png")  # Replace with your image path
        icon_image = icon_image.resize((32, 32))  # Resize the image if needed
        photo_icon = ImageTk.PhotoImage(icon_image)
        root.iconphoto(True, photo_icon)  # Set the image as window icon
    except Exception as e:
        print(f"Failed to set window icon: {e}")

    # Title Label
    title_label = tk.Label(root, text="Maple pdf", font=("Forte", 90, "bold"), bg="#FF7733", fg="#15616D")
    title_label.pack(pady=10)

    frame = tk.Frame(root, padx=20, pady=20, bg="#FF7733")
    frame.pack(pady=10)

    merge_label = tk.Label(frame, text="Merge PDFs", font=("Felix Titling", 14, "bold"), bg="#FF7733", fg="#FFECD1")
    merge_label.grid(row=0, column=0, columnspan=2, pady=10)

    ttk.Button(frame, text="Merge PDFs", width=30, command=PDFmerge).grid(row=1, column=0, columnspan=2, pady=10)

    separator = ttk.Separator(frame, orient='horizontal')
    separator.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)

    convert_label = tk.Label(frame, text="Convert PDF", font=("Felix Titling", 14, "bold"), bg="#FF7733", fg="#FFECD1")
    convert_label.grid(row=3, column=0, columnspan=2, pady=10)

    formats = [
        ("Convert to DOCX", lambda: convert_pdf_to("docx")),
        ("Convert to TXT", lambda: convert_pdf_to("txt")),
        ("Convert to HTML", lambda: convert_pdf_to("html")),
        ("Convert to XML", lambda: convert_pdf_to("xml")),
        ("Convert to JSON", lambda: convert_pdf_to("json")),
        ("Convert to Markdown", lambda: convert_pdf_to("md")),
        ("Convert to CSV", lambda: convert_pdf_to("csv")),
        ("Convert to PNG", lambda: convert_pdf_to("png")),
        ("Convert to JPG", lambda: convert_pdf_to("jpg")),
        ("Convert to GIF", lambda: convert_pdf_to("gif")),
        ("Convert to SVG", lambda: convert_pdf_to("svg")),
        ("Convert to MP3", convert_pdf_to_mp3),
    ]

    for i, (text, command) in enumerate(formats):
        ttk.Button(frame, text=text, width=60, command=command).grid(row=4 + i, column=0, columnspan=2, pady=5)

    footer_label = tk.Label(root, text="Made with love by shibbux", font=("Arial", 20), bg="#FF7733", fg="#FFECD1")
    footer_label.pack(side="bottom", pady=1)

    git_prom = tk.Label(root, text="Follow (shibbux) on github ", font=("Arial", 30), bg="#FF7733", fg="#FFECD1")
    git_prom.pack(side="bottom", pady=1)

    root.mainloop()


if __name__ == "__main__":
    main()
