import aspose.words as aw
import PyPDF2  # Changed library for PDF merging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Function to open file dialog for selecting PDF files
def open_file_dialog():
    root = tk.Tk()
    root.withdraw()  # Hide the small Tkinter window
    file_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )
    return file_path

# Function to open multiple files for merging
def open_multiple_files_dialog():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select PDF Files to Merge",
        filetypes=[("PDF Files", "*.pdf")]
    )
    return file_paths

# Function to open exactly two files for merging
def open_two_files_dialog():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select Two PDF Files to Merge",
        filetypes=[("PDF Files", "*.pdf")],
        multiple=True
    )
    if len(file_paths) != 2:
        messagebox.showwarning("Warning", "Please select exactly two PDF files!")
        return None
    return file_paths

# Function for merging two PDFs
def merge_two_pdfs():
    pdfs = open_two_files_dialog()  # Select exactly two PDFs to merge
    if pdfs:
        output = filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if output:
            try:
                pdfMerger = PyPDF2.PdfMerger()  # Changed to PyPDF2's PdfMerger
                for pdf in pdfs:
                    with open(pdf, 'rb') as f:
                        pdfMerger.append(f)
                with open(output, 'wb') as f:
                    pdfMerger.write(f)
                messagebox.showinfo("Success", "Two PDFs merged successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to merge PDFs. {str(e)}")
        else:
            messagebox.showwarning("Warning", "No output file selected!")
    else:
        messagebox.showwarning("Warning", "No PDF files selected for merging!")

# Function for merging multiple PDFs
def PDFmerge():
    pdfs = open_multiple_files_dialog()  # Select multiple PDFs to merge
    if pdfs:
        output = filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if output:
            try:
                pdfMerger = PyPDF2.PdfMerger()  # Changed to PyPDF2's PdfMerger
                for pdf in pdfs:
                    with open(pdf, 'rb') as f:
                        pdfMerger.append(f)
                with open(output, 'wb') as f:
                    pdfMerger.write(f)
                messagebox.showinfo("Success", "PDFs merged successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to merge PDFs. {str(e)}")
        else:
            messagebox.showwarning("Warning", "No output file selected!")
    else:
        messagebox.showwarning("Warning", "No PDF files selected for merging!")

# Conversion functions
def convert_pdf_to(format):
    file_path = open_file_dialog()  # Select a single PDF file
    if file_path:
        doc = aw.Document(file_path)
        output_path = filedialog.asksaveasfilename(
            title=f"Save PDF as {format.upper()}",
            defaultextension=f".{format}",
            filetypes=[(f"{format.upper()} Files", f"*.{format}")]
        )
        if output_path:
            doc.save(output_path)
            messagebox.showinfo("Success", f"PDF converted to {format.upper()} successfully!")
        else:
            messagebox.showwarning("Warning", "No output file selected!")
    else:
        messagebox.showwarning("Warning", "No PDF file selected!")

# Image conversion function (PNG, SVG, JPG, GIF)
def convert_pdf_to_images(format):
    file_path = open_file_dialog()
    if file_path:
        doc = aw.Document(file_path)
        for page in range(0, doc.page_count):
            output_path = filedialog.asksaveasfilename(
                title=f"Save Page {page + 1} as {format.upper()}",
                defaultextension=f".{format}",
                filetypes=[(f"{format.upper()} Files", f"*.{format}")]
            )
            if output_path:
                extractedPage = doc.extract_pages(page, 1)
                extractedPage.save(output_path)
        messagebox.showinfo("Success", f"PDF converted to {format.upper()} successfully!")
    else:
        messagebox.showwarning("Warning", "No PDF file selected!")

# Main GUI setup
def main():
    root = tk.Tk()
    root.title("PDF Tool")
    root.geometry("450x550")
    root.resizable(False, False)
    root.configure(bg="#f2f2f2")  # Light grey background

    # Title Label
    title_label = tk.Label(root, text="PDF Conversion Tool", font=("Arial", 20, "bold"), bg="#f2f2f2", fg="#333")
    title_label.pack(pady=10)

    # Create a style for buttons
    style = ttk.Style()
    style.configure("TButton", font=("Arial", 12, "bold"), padding=6, background="#3498db", foreground="#6C2B9C")
    style.map("TButton", foreground=[('pressed', 'white'), ('active', 'white')],
              background=[('pressed', '!disabled', '#2980b9'), ('active', '#2980b9')])

    # Frame for better layout with section dividers
    frame = tk.Frame(root, padx=20, pady=20, bg="#f2f2f2")
    frame.pack(pady=10)

    # Section Title for Merging
    merge_label = tk.Label(frame, text="Merge PDFs", font=("Arial", 14, "bold"), bg="#f2f2f2", fg="#2980b9")
    merge_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))

    # Merge Button for merging two PDFs
    btn_merge_two = ttk.Button(frame, text="Merge Two PDFs", width=30, command=merge_two_pdfs)
    btn_merge_two.grid(row=1, column=0, columnspan=2, pady=10)

    # Merge Button for merging multiple PDFs
    btn_merge_multiple = ttk.Button(frame, text="Merge Multiple PDFs", width=30, command=PDFmerge)
    btn_merge_multiple.grid(row=2, column=0, columnspan=2, pady=10)

    # Section Divider
    separator = ttk.Separator(frame, orient='horizontal')
    separator.grid(row=3, column=0, columnspan=2, sticky="ew", pady=10)

    # Section Title for Conversion
    convert_label = tk.Label(frame, text="Convert PDF", font=("Arial", 14, "bold"), bg="#f2f2f2", fg="#2980b9")
    convert_label.grid(row=4, column=0, columnspan=2, pady=(10, 10))

    # Conversion buttons
    conversion_options = [
        ("Convert PDF to DOCX", lambda: convert_pdf_to("docx")),
        ("Convert PDF to HTML", lambda: convert_pdf_to("html")),
        ("Convert PDF to TXT", lambda: convert_pdf_to("txt")),
        ("Convert PDF to DOC", lambda: convert_pdf_to("doc")),
        ("Convert PDF to EPUB", lambda: convert_pdf_to("epub")),
        ("Convert PDF to PNG", lambda: convert_pdf_to_images("png")),
        ("Convert PDF to SVG", lambda: convert_pdf_to_images("svg")),
        ("Convert PDF to JPG", lambda: convert_pdf_to_images("jpg")),
        ("Convert PDF to GIF", lambda: convert_pdf_to_images("gif")),
    ]

    for i, (text, command) in enumerate(conversion_options):
        btn = ttk.Button(frame, text=text, width=30, command=command)
        btn.grid(row=5 + i, column=0, columnspan=2, pady=5)

    # Footer Section
    footer_label = tk.Label(root, text="Powered by Aspose.Words & Tkinter", font=("Arial", 10), bg="#f2f2f2", fg="#666")
    footer_label.pack(side="bottom", pady=10)

    # Start the main loop
    root.mainloop()

if __name__ == "__main__":
    main()
