import tkinter as tk
from tkinter import filedialog, messagebox
import os
import platform

try:
    from docx2pdf import convert
except ImportError:
    messagebox.showerror(
        "Dependency Missing",
        "The 'docx2pdf' library is not installed.\n\n"
        "Please install it by running:\n"
        "pip install docx2pdf\n\n"
        "Note: On Windows, 'docx2pdf' often requires Microsoft Word to be installed. "
        "On macOS/Linux, it may use LibreOffice if Word is not available."
    )
    exit()

class WordToPdfConverterApp:
    """GUI application to convert Word (.docx) files to PDF."""

    def __init__(self, master_root: tk.Tk) -> None:
        self.master = master_root
        self.master.title("Word to PDF Converter")
        self.master.geometry("500x300")

        self.selected_files: list[str] = []
        self.output_directory = ""

        selection_frame = tk.Frame(self.master, pady=10)
        selection_frame.pack(fill=tk.X, padx=10)

        self.browse_button = tk.Button(
            selection_frame,
            text="1. Browse for .docx Files",
            command=self.browse_files,
            width=25,
        )
        self.browse_button.pack(side=tk.LEFT, padx=(0, 10))

        self.status_label = tk.Label(selection_frame, text="No files selected.", anchor="w")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        output_dir_frame = tk.Frame(self.master, pady=5)
        output_dir_frame.pack(fill=tk.X, padx=10)

        self.output_dir_button = tk.Button(
            output_dir_frame,
            text="2. Select Output Folder (Optional)",
            command=self.select_output_directory,
            width=25,
        )
        self.output_dir_button.pack(side=tk.LEFT, padx=(0, 10))

        self.output_dir_label = tk.Label(
            output_dir_frame,
            text="Default: Same as input file's folder",
            anchor="w",
        )
        self.output_dir_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        action_frame = tk.Frame(self.master, pady=10)
        action_frame.pack(fill=tk.X, padx=10)

        self.convert_button = tk.Button(
            action_frame,
            text="3. Convert Selected Files to PDF",
            command=self.convert_selected_files,
            state=tk.DISABLED,
            width=25,
        )
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))

        self.complete_button = tk.Button(
            action_frame,
            text="4. Finish & Exit",
            command=self.master.quit,
            state=tk.DISABLED,
        )
        self.complete_button.pack(side=tk.LEFT)

        info_text = "This tool uses 'docx2pdf'."
        if platform.system() == "Windows":
            info_text += " MS Word may be required."
        elif platform.system() in ["Linux", "Darwin"]:
            info_text += " LibreOffice might be used if Word is not found."

        info_label = tk.Label(self.master, text=info_text, fg="gray", pady=10)
        info_label.pack(fill=tk.X, padx=10)

    def browse_files(self) -> None:
        self.selected_files = filedialog.askopenfilenames(
            title="Select Word Documents (.docx)",
            filetypes=(("Word documents", "*.docx"), ("All files", "*.*")),
        )

        if self.selected_files:
            file_count = len(self.selected_files)
            self.status_label.config(text=f"{file_count} file(s) selected.")
            self.convert_button.config(state=tk.NORMAL)
            self.complete_button.config(state=tk.DISABLED)
        else:
            self.status_label.config(text="No files selected.")
            self.convert_button.config(state=tk.DISABLED)
            self.selected_files = []

    def select_output_directory(self) -> None:
        chosen_directory = filedialog.askdirectory(title="Select Output Folder for PDFs")
        if chosen_directory:
            self.output_directory = chosen_directory
            self.output_dir_label.config(text=f"Output to: {self.output_directory}")
        else:
            self.output_directory = ""
            self.output_dir_label.config(text="Default: Same as input file's folder")

    def convert_selected_files(self) -> None:
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select .docx files to convert.")
            return

        converted_files_paths: list[str] = []
        error_messages: list[str] = []
        processed_count = 0
        success_count = 0

        self.convert_button.config(state=tk.DISABLED)

        for docx_file_path in self.selected_files:
            processed_count += 1
            self.status_label.config(
                text=f"Processing {processed_count}/{len(self.selected_files)}: "
                f"{os.path.basename(docx_file_path)}..."
            )
            self.master.update_idletasks()

            try:
                file_name_without_ext = os.path.splitext(os.path.basename(docx_file_path))[0]
                pdf_file_name = f"{file_name_without_ext}.pdf"

                if self.output_directory:
                    output_path = os.path.join(self.output_directory, pdf_file_name)
                    convert(docx_file_path, output_path)
                else:
                    convert(docx_file_path)
                    output_path = os.path.join(os.path.dirname(docx_file_path), pdf_file_name)

                converted_files_paths.append(output_path)
                success_count += 1
            except Exception as e:
                error_msg = f"Error converting '{os.path.basename(docx_file_path)}': {e}"
                error_messages.append(error_msg)
                print(f"Error: {error_msg}")

        message_title = "Conversion Report"
        final_message = (
            f"Conversion process finished.\n\nSuccessfully converted: {success_count} file(s)."
        )

        if converted_files_paths:
            final_message += "\n\nCreated PDF files:\n" + "\n".join(converted_files_paths)

        if error_messages:
            final_message += "\n\nErrors encountered:\n" + "\n".join(error_messages)
            messagebox.showerror(message_title, final_message)
        elif success_count > 0:
            messagebox.showinfo(message_title, final_message)
        else:
            messagebox.showwarning(
                message_title,
                "No files were converted. Please check selection or errors.",
            )

        self.selected_files = []
        self.status_label.config(text="No files selected. Browse again or finish.")
        self.convert_button.config(state=tk.DISABLED)
        self.complete_button.config(state=tk.NORMAL)


def main() -> None:
    root = tk.Tk()
    app = WordToPdfConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

