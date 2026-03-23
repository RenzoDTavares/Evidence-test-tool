import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel
from PIL import Image, ImageTk
from pathlib import Path
import base64
from core.config import logger, CONFIG, IMAGE_EXTENSIONS, ICON_FILE, KEY_ICON_FILE
from controllers.qa_controller import QAController

class TestAssistantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QA Test Evidence Assistant")
        self.root.geometry("720x520")
        try:
            self.root.iconbitmap(ICON_FILE)
        except tk.TclError:
            pass 

        self.entries = {}
        self.vars = {
            "devops_integration": tk.BooleanVar(value=False),
            "env": tk.StringVar(value='STG')
        }
        
        self._setup_styles()
        self._build_ui()
        logger.info("Application UI started.")

    def _setup_styles(self):
        style = ttk.Style()
        style.configure("TEntry", padding=(5, 4), relief="flat")
        style.configure("TCombobox", padding=(5, 4), relief="flat")
        style.configure("TButton", padding=(10, 5), relief="flat")
        style.map('TCombobox',
            fieldbackground=[('readonly', 'white'), ('focus', 'white')],
            selectbackground=[('readonly', 'white')],
            selectforeground=[('readonly', 'black')]
        )

    def _build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        main_tab = ttk.Frame(notebook)
        notebook.add(main_tab, text="Document Generator")

        ui_labels = {
            "template_file": "Template File", "img_dir": "Image Directory",
            "tester": "Tester Name", "test_id": "Test Case ID/Name",
            "env": "Environment", "product": "Product",
            "profile": "Profile", "bugs": "Bugs"
        }

        for i, (key, text) in enumerate(ui_labels.items()):
            ttk.Label(main_tab, text=text).grid(row=i, column=0, padx=10, pady=8, sticky='w')
            if key in ["env", "product"]: continue
            
            entry = ttk.Entry(main_tab, width=45)
            entry.grid(row=i, column=1, padx=5, pady=8, sticky='we')
            self.entries[key] = entry

        ttk.Button(main_tab, text="Browse", command=lambda: self._browse_file(self.entries['template_file'])).grid(row=0, column=2, padx=5)
        ttk.Button(main_tab, text="Browse", command=lambda: self._select_directory(self.entries['img_dir'])).grid(row=1, column=2, padx=5)
        ttk.Button(main_tab, text="Clear Images", command=self._clear_images).grid(row=1, column=3, padx=5)

        env_dropdown = ttk.Combobox(main_tab, textvariable=self.vars['env'], values=["DEV", "STG", "PRD"], state="readonly", width=42)
        env_dropdown.grid(row=4, column=1, padx=5, pady=8, sticky='we')

        devops_frame = ttk.Frame(main_tab)
        devops_frame.grid(row=2, column=2, columnspan=3, sticky='w', padx=5)
        
        ttk.Checkbutton(devops_frame, text="DevOps Integration", variable=self.vars['devops_integration'], command=self._toggle_devops_integration).pack(side='left')
        
        try:
            key_img = ImageTk.PhotoImage(Image.open(KEY_ICON_FILE).resize((16, 16)))
            self.key_button = ttk.Button(devops_frame, image=key_img, command=self._show_key_dialog)
            self.key_button.image = key_img
            self.key_button.pack(side='left', padx=10)
        except Exception:
            ttk.Button(devops_frame, text="🔑", command=self._show_key_dialog).pack(side='left', padx=10)

        self.project_dropdown = ttk.Combobox(main_tab, state="disabled", width=42)
        self.project_dropdown.grid(row=5, column=1, padx=5, pady=8, sticky='we')
        
        self.generate_button = ttk.Button(main_tab, text="Generate Evidence", command=self._start_generation)
        self.generate_button.grid(row=8, column=1, columnspan=1, padx=5, pady=20, sticky='e')

    def _safe_gui_update(self, func, *args, **kwargs):
        self.root.after(0, lambda: func(*args, **kwargs))

    def _show_msg(self, title, message, msg_type="info"):
        funcs = {"info": messagebox.showinfo, "warning": messagebox.showwarning, "error": messagebox.showerror}
        funcs[msg_type](title, message)

    def _browse_file(self, entry_widget):
        file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if file_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, file_path)

    def _select_directory(self, entry_widget):
        dir_path = filedialog.askdirectory()
        if dir_path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, dir_path)

    def _clear_images(self):
        dir_str = self.entries['img_dir'].get()
        if not dir_str:
            self._show_msg("Warning", "Please select an image directory first.", "warning")
            return

        img_dir = Path(dir_str)
        if not img_dir.is_dir():
            self._show_msg("Error", f"Directory '{img_dir}' does not exist.", "error")
            return
            
        removed_count = 0
        for file in img_dir.iterdir():
            if file.suffix.lower() in IMAGE_EXTENSIONS:
                try:
                    file.unlink()
                    removed_count += 1
                except Exception as e:
                    logger.error(f"Failed to delete {file.name}: {e}")
        
        self._show_msg("Cleanup Complete", f"{removed_count} image(s) removed.")
        logger.info(f"Cleared {removed_count} images from {img_dir}")

    def _toggle_devops_integration(self):
        if self.vars['devops_integration'].get():
            self.project_dropdown.set('Loading projects...')
            self.project_dropdown.config(state='disabled')
            
            QAController.fetch_projects_async(
                callback_success=lambda projs: self._safe_gui_update(self._apply_projects, projs),
                callback_error=lambda title, msg: self._safe_gui_update(self._fail_projects, title, msg)
            )
        else:
            self.project_dropdown.set('')
            self.project_dropdown.config(state='disabled', values=[])

    def _apply_projects(self, projects):
        if projects:
            self.project_dropdown.config(state='readonly', values=projects)
            self.project_dropdown.set('')
        else:
            self._fail_projects("Error", "No projects found in the organization.")

    def _fail_projects(self, title, msg):
        self._show_msg(title, msg, "error" if title != "Missing Token" else "warning")
        self.vars['devops_integration'].set(False)
        self.project_dropdown.config(state='disabled', values=[])

    def _start_generation(self):
        data = {key: entry.get().strip() for key, entry in self.entries.items()}
        data['env'] = self.vars['env'].get()
        data['is_devops'] = self.vars['devops_integration'].get()
        data['project'] = self.project_dropdown.get()

        if not data['template_file'] or not data['img_dir']:
            self._show_msg("Error", "Template File and Image Directory are required.", "error")
            return

        if data['is_devops'] and not data['project']:
            self._show_msg("Error", "Please select a DevOps Project.", "error")
            return

        self.generate_button.config(state='disabled', text="Generating...")
        
        QAController.generate_evidence_async(
            data=data,
            callback_success=lambda t, m, is_w=False: self._safe_gui_update(self._show_msg, t, m, "warning" if is_w else "info"),
            callback_error=lambda t, m: self._safe_gui_update(self._show_msg, t, m, "error"),
            callback_finally=lambda: self._safe_gui_update(self.generate_button.config, state='normal', text="Generate Evidence")
        )

    def _show_key_dialog(self):
        dialog = Toplevel(self.root)
        dialog.title("DevOps Configuration")
        dialog.geometry("380x160")
        
        ttk.Label(dialog, text="Enter your Personal Access Token (PAT):").pack(pady=10)
        key_entry = ttk.Entry(dialog, width=50, show="*")
        key_entry.pack(pady=5, padx=10)

        def save_key():
            pat = key_entry.get().strip()
            if not pat: return
            try:
                encoded_key = base64.b64encode(pat.encode('utf-8')).decode('utf-8')
                with open(CONFIG["KEY_FILE"], "w") as f:
                    f.write(encoded_key)
                messagebox.showinfo("Success", "Token saved successfully.", parent=dialog)
                logger.info("New PAT token saved by user.")
                dialog.destroy()
            except Exception as e:
                logger.error(f"Failed to save PAT token: {e}")
                messagebox.showerror("Error", f"Failed to save token: {e}", parent=dialog)
        
        ttk.Button(dialog, text="Save Configuration", command=save_key).pack(pady=10)