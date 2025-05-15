import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import PyPDF2
from docx import Document

class StrictResumeChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("Strict Resume Qualification Checker")
        self.root.geometry("1050x800")
        
        self.uploaded_files = []
        self.required_qualifications = []
        
        self.create_widgets()
        self.create_directories()
    
    def create_directories(self):
        os.makedirs("Accepted", exist_ok=True)
        os.makedirs("Rejected", exist_ok=True)
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # File Upload Section
        upload_frame = ttk.LabelFrame(main_frame, text="Upload Resumes", padding=10)
        upload_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(upload_frame, text="Select Files", command=self.upload_files).pack(side=tk.LEFT)
        self.file_count = ttk.Label(upload_frame, text="0 files selected")
        self.file_count.pack(side=tk.LEFT, padx=10)

        ttk.Button(upload_frame, text="Check Eligibility", command=self.check_eligibility).pack(side=tk.LEFT, padx=10)
        
        # Qualifications Input
        qual_frame = ttk.LabelFrame(main_frame, text="Required Qualifications (one per line)", padding=10)
        qual_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.qual_text = scrolledtext.ScrolledText(qual_frame, height=10, wrap=tk.WORD)
        self.qual_text.pack(fill=tk.BOTH, expand=True)
        self.qual_text.insert(tk.END, "Bachelor's Degree\nPython Programming\n3+ years experience\nProject Management")
        
        # Results Display
        result_frame = ttk.LabelFrame(main_frame, text="Results", padding=10)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("filename", "status", "missing")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        self.result_tree.heading("filename", text="File Name")
        self.result_tree.heading("status", text="Status")
        self.result_tree.heading("missing", text="Missing Qualifications")
        self.result_tree.column("filename", width=250)
        self.result_tree.column("status", width=100)
        self.result_tree.column("missing", width=600)
        self.result_tree.pack(fill=tk.BOTH, expand=True)

        self.result_tree.bind("<Double-1>", self.preview_content)

        # Control Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="Clear Results", command=self.clear_results).pack(side=tk.LEFT)

    def upload_files(self):
        files = filedialog.askopenfilenames(filetypes=(("Resume Files", "*.pdf;*.docx;*.txt"),))
        if files:
            self.uploaded_files = files
            self.file_count.config(text=f"{len(files)} files selected")
    
    def clear_results(self):
        self.result_tree.delete(*self.result_tree.get_children())

    def check_eligibility(self):
        if not self.uploaded_files:
            messagebox.showwarning("Warning", "Please select resume files first")
            return
        
        self.required_qualifications = [
            q.strip() for q in self.qual_text.get("1.0", tk.END).splitlines() if q.strip()
        ]
        if not self.required_qualifications:
            messagebox.showwarning("Warning", "Please define required qualifications")
            return

        threading.Thread(target=self.process_resumes, daemon=True).start()

    def process_resumes(self):
        self.clear_results()
        for file_path in self.uploaded_files:
            filename = os.path.basename(file_path)
            status, missing = self.analyze_resume(file_path)

            self.result_tree.insert("", "end", values=(
                filename,
                "Approved" if status else "Rejected",
                ", ".join(missing) if missing else "All requirements met"
            ))

            self.move_file(file_path, status)
        
        messagebox.showinfo("Done", "Eligibility check completed.")

    def analyze_resume(self, file_path):
        try:
            content = self.extract_content(file_path).lower()
            required_lower = [q.lower() for q in self.required_qualifications]

            missing = [
                qual for qual, req in zip(self.required_qualifications, required_lower)
                if req not in content
            ]
            return (len(missing) == 0, missing)

        except Exception as e:
            return (False, [f"Error reading file: {str(e)}"])

    def extract_content(self, file_path):
        if file_path.endswith(".pdf"):
            with open(file_path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                return " ".join([page.extract_text() or "" for page in reader.pages])
        elif file_path.endswith(".docx"):
            doc = Document(file_path)
            return " ".join([p.text for p in doc.paragraphs])
        else:
            with open(file_path, "r", encoding="utf-8") as f:
                return f.read()

    def move_file(self, src, approved):
        dest_dir = "Accepted" if approved else "Rejected"
        filename = os.path.basename(src)
        dest_path = os.path.join(dest_dir, filename)
        shutil.copy(src, dest_path)

    def preview_content(self, event):
        item_id = self.result_tree.focus()
        if not item_id:
            return
        
        index = self.result_tree.index(item_id)
        file_path = self.uploaded_files[index]

        try:
            content = self.extract_content(file_path)
            preview_win = tk.Toplevel(self.root)
            preview_win.title("Resume Preview")
            preview_win.geometry("800x600")
            text_area = scrolledtext.ScrolledText(preview_win, wrap=tk.WORD)
            text_area.insert(tk.END, content)
            text_area.pack(fill=tk.BOTH, expand=True)
        except Exception as e:
            messagebox.showerror("Error", f"Could not preview file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = StrictResumeChecker(root)
    root.mainloop()
