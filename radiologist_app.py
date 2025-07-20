import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os
import sys
import subprocess
import json
import random
import string
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

class RadiologistApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Radiologist Report Generator")
        self.root.geometry("900x800")
        self.root.configure(bg='#f8fafc')
        self.radiologist_name = "Dr. Noor Ahmad"
        
        # Get the application directory (works for both script and executable)
        if getattr(sys, 'frozen', False):
            # Running as executable
            self.app_dir = os.path.dirname(sys.executable)
        else:
            # Running as script
            self.app_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Create Reports directory in the same location as the executable/script
        self.reports_dir = os.path.join(self.app_dir, "Reports")
        if not os.path.exists(self.reports_dir):
            os.makedirs(self.reports_dir)
        
        # Load admin data first
        self.load_admin_data()
        
        # Initialize patient ID counter
        self.patient_id_counter = self.load_patient_id_counter()
        
        self.setup_styles()
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#f8fafc')
        style.configure('TLabel', background='#f8fafc', font=("Segoe UI", 11))
        style.configure('Header.TLabel', font=("Segoe UI", 16, "bold"), foreground="#1e293b", background='#f8fafc')
        style.configure('Subheader.TLabel', font=("Segoe UI", 14, "bold"), foreground="#475569", background='#f8fafc')
        style.configure('Section.TLabelframe', background='#f8fafc', foreground="#1e293b", font=("Segoe UI", 12, "bold"))
        style.configure('Section.TLabelframe.Label', font=("Segoe UI", 12, "bold"), foreground="#1e293b", background='#f8fafc')
        style.configure('Accent.TButton', font=("Segoe UI", 11, "bold"), background="#3b82f6", foreground="white", borderwidth=0, padding=(12, 10), relief="flat")
        style.map('Accent.TButton', background=[('active', '#2563eb')])
        style.configure('Success.TButton', font=("Segoe UI", 11, "bold"), background="#10b981", foreground="white", borderwidth=0, padding=(12, 10), relief="flat")
        style.map('Success.TButton', background=[('active', '#059669')])
        style.configure('Danger.TButton', font=("Segoe UI", 11), background="#ef4444", foreground="white", borderwidth=0, padding=(12, 10), relief="flat")
        style.map('Danger.TButton', background=[('active', '#dc2626')])
        style.configure('TButton', font=("Segoe UI", 11), padding=(10, 8), relief="flat")
        style.configure('TEntry', font=("Segoe UI", 11))
        style.configure('TCombobox', font=("Segoe UI", 11))
        style.configure('TNotebook', background='#f8fafc')
        style.configure('TNotebook.Tab', font=("Segoe UI", 10, "bold"), padding=[20, 10])

    def create_widgets(self):
        # Header with logo and title
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", padx=20, pady=(10, 0))
        
        # Logo
        logo_frame = ttk.Frame(header_frame)
        logo_frame.pack(side="left")
        logo = tk.Canvas(logo_frame, width=32, height=32, bg='#f8fafc', highlightthickness=0)
        logo.create_oval(2, 2, 30, 30, fill="#3b82f6", outline="")
        logo.create_text(16, 16, text="Rx", fill="white", font=("Segoe UI", 12, "bold"))
        logo.pack(side="left", padx=(0, 8))
        
        # Title and info
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side="left", fill="x", expand=True)
        ttk.Label(title_frame, text="Radiologist Report Generator", style='Header.TLabel').pack(anchor='w')
        info_frame = ttk.Frame(title_frame)
        info_frame.pack(anchor='w', pady=(1, 0))
        ttk.Label(info_frame, text=f"Radiologist: {self.radiologist_name}", font=("Segoe UI", 9, "bold"), foreground="#475569").pack(side="left", padx=(0, 12))
        ttk.Label(info_frame, text=f"Date: {datetime.now().strftime('%B %d, %Y')}", font=("Segoe UI", 8), foreground="#64748b").pack(side="left")

        # Main content area with tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=20, pady=20)

        # Tab 1: Patient Information
        patient_tab = ttk.Frame(notebook)
        notebook.add(patient_tab, text="Patient Information")
        self.create_patient_tab(patient_tab)

        # Tab 2: Report Details
        report_tab = ttk.Frame(notebook)
        notebook.add(report_tab, text="Report Details")
        self.create_report_tab(report_tab)

        # Tab 3: Preview & Save
        preview_tab = ttk.Frame(notebook)
        notebook.add(preview_tab, text="Preview & Save")
        self.create_preview_tab(preview_tab)

        # Tab 4: View All Reports
        reports_tab = ttk.Frame(notebook)
        notebook.add(reports_tab, text="View All Reports")
        self.create_reports_tab(reports_tab)

        # Tab 5: Admin Settings
        admin_tab = ttk.Frame(notebook)
        notebook.add(admin_tab, text="Admin Settings")
        self.create_admin_tab(admin_tab)

        # Status bar
        self.status_label = ttk.Label(self.root, text="Ready to create reports", font=("Segoe UI", 10), foreground="#64748b", background='#f8fafc')
        self.status_label.pack(side="bottom", fill="x", padx=20, pady=(0, 10))

        # Auto-generate first Patient ID
        self.generate_new_patient_id()

    def load_patient_id_counter(self):
        """Load the current patient ID counter from file"""
        counter_file = "patient_id_counter.json"
        try:
            if os.path.exists(counter_file):
                with open(counter_file, 'r') as f:
                    data = json.load(f)
                    return data.get('counter', 1)
            return 1
        except:
            return 1

    def save_patient_id_counter(self):
        """Save the current patient ID counter to file"""
        counter_file = "patient_id_counter.json"
        try:
            with open(counter_file, 'w') as f:
                json.dump({'counter': self.patient_id_counter}, f)
        except Exception as e:
            print(f"Error saving patient ID counter: {e}")

    def generate_patient_id(self):
        """Generate a unique patient ID"""
        # Format: P + 6-digit number (e.g., P000001, P000002, etc.)
        patient_id = f"P{self.patient_id_counter:06d}"
        self.patient_id_counter += 1
        self.save_patient_id_counter()
        return patient_id

    def generate_new_patient_id(self):
        """Generate and display a new patient ID"""
        new_id = self.generate_patient_id()
        self.patient_id_entry.config(state='normal')
        self.patient_id_entry.delete(0, tk.END)
        self.patient_id_entry.insert(0, new_id)
        self.patient_id_entry.config(state='readonly')
        self.status_label.config(text=f"Generated Patient ID: {new_id}", foreground="#10b981")

    def create_patient_tab(self, parent):
        # Patient Information Section
        patient_frame = ttk.Labelframe(parent, text="Patient Information", style='Section.TLabelframe', padding=(20, 15))
        patient_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Grid layout for better organization
        # Row 1
        ttk.Label(patient_frame, text="Patient ID:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").grid(row=0, column=0, sticky='w', pady=(0, 10), padx=(0, 20))
        
        self.patient_id_entry = ttk.Entry(patient_frame, width=15, font=("Segoe UI", 11), state='readonly')
        self.patient_id_entry.grid(row=0, column=1, sticky='w', pady=(0, 10), padx=(0, 20))
        
        # Generate Patient ID button
        self.generate_id_btn = ttk.Button(patient_frame, text="🔄 Generate ID", style='Accent.TButton', 
                                         command=self.generate_new_patient_id)
        self.generate_id_btn.grid(row=0, column=2, sticky='w', pady=(0, 10), padx=(10, 0))

        ttk.Label(patient_frame, text="Patient Name:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").grid(row=0, column=3, sticky='w', pady=(0, 10), padx=(0, 20))
        self.patient_name_entry = ttk.Entry(patient_frame, width=25, font=("Segoe UI", 11))
        self.patient_name_entry.grid(row=0, column=4, sticky='ew', pady=(0, 10))

        # Row 2
        ttk.Label(patient_frame, text="Age:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").grid(row=1, column=0, sticky='w', pady=(0, 10), padx=(0, 20))
        self.age_entry = ttk.Entry(patient_frame, width=10, font=("Segoe UI", 11))
        self.age_entry.grid(row=1, column=1, sticky='w', pady=(0, 10), padx=(0, 20))

        # Row 3
        ttk.Label(patient_frame, text="Test Type:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").grid(row=2, column=0, sticky='w', pady=(0, 10), padx=(0, 20))
        self.test_type_var = tk.StringVar()
        self.test_type_combo = ttk.Combobox(patient_frame, textvariable=self.test_type_var, values=self.admin_data["test_types"], state="readonly", width=27, font=("Segoe UI", 11))
        self.test_type_combo.grid(row=2, column=1, sticky='ew', pady=(0, 10), padx=(0, 20))
        self.test_type_combo.set("Select Test Type")
        self.test_type_combo.bind('<<ComboboxSelected>>', self.on_test_type_change)

        ttk.Label(patient_frame, text="Body Part:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").grid(row=2, column=2, sticky='w', pady=(0, 10), padx=(0, 20))
        self.body_part_var = tk.StringVar()
        body_parts = ["Chest", "Abdomen", "Head", "Spine", "Extremities", "Pelvis", "Other"]
        self.body_part_combo = ttk.Combobox(patient_frame, textvariable=self.body_part_var, values=body_parts, state="readonly", width=15, font=("Segoe UI", 11))
        self.body_part_combo.grid(row=2, column=3, sticky='w', pady=(0, 10))
        self.body_part_combo.set("Select Body Part")
        self.body_part_combo.bind('<<ComboboxSelected>>', self.on_body_part_change)

        # Configure grid weights
        patient_frame.columnconfigure(1, weight=1)
        
        # Add Clear All button for patient information
        clear_patient_btn = ttk.Button(patient_frame, text="🗑️ Clear Patient Info", style='Danger.TButton', 
                                      command=self.clear_patient_info)
        clear_patient_btn.grid(row=3, column=0, columnspan=5, pady=(10, 0), sticky='ew')

    def create_report_tab(self, parent):
        # Report Template Section
        template_frame = ttk.Labelframe(parent, text="Report Template", style='Section.TLabelframe', padding=(15, 10))
        template_frame.pack(fill="x", padx=20, pady=(15, 8))

        # Template selection
        template_btn_frame = ttk.Frame(template_frame)
        template_btn_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(template_btn_frame, text="Load Template:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(side="left", padx=(0, 10))
        self.template_button = ttk.Button(template_btn_frame, text="📋 Load Template", style='Accent.TButton', command=self.load_template)
        self.template_button.pack(side="left", ipadx=10, ipady=5)

        # Report content
        content_frame = ttk.Labelframe(parent, text="Report Content", style='Section.TLabelframe', padding=(15, 12))
        content_frame.pack(fill="both", expand=True, padx=20, pady=(8, 15))

        # Technique with rich text editing
        ttk.Label(content_frame, text="Technique:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Technique formatting toolbar
        technique_toolbar = ttk.Frame(content_frame)
        technique_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(technique_toolbar, text="B", font=("Segoe UI", 10, "bold"), width=3, 
                 command=lambda: self.toggle_bold(self.technique_text)).pack(side="left", padx=(0, 2))
        tk.Button(technique_toolbar, text="I", font=("Segoe UI", 10, "italic"), width=3,
                 command=lambda: self.toggle_italic(self.technique_text)).pack(side="left", padx=(0, 2))
        tk.Button(technique_toolbar, text="•", font=("Segoe UI", 10), width=3,
                 command=lambda: self.add_bullet(self.technique_text)).pack(side="left", padx=(0, 2))
        tk.Button(technique_toolbar, text="📋", font=("Segoe UI", 10), width=3,
                 command=lambda: self.paste_table(self.technique_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown for technique
        tk.Label(technique_toolbar, text="Size:", font=("Segoe UI", 9)).pack(side="left", padx=(10, 2))
        technique_size_var = tk.StringVar(value="11")
        technique_size_combo = ttk.Combobox(technique_toolbar, textvariable=technique_size_var, 
                                           values=["10", "11", "12", "14", "16"], width=5, state="readonly")
        technique_size_combo.pack(side="left", padx=(0, 10))
        technique_size_combo.bind('<<ComboboxSelected>>', 
                                 lambda e: self.change_font_size(self.technique_text, technique_size_var.get()))
        
        self.technique_text = tk.Text(content_frame, height=3, font=("Segoe UI", 11), relief='solid', bd=1, wrap='word', 
                                     highlightbackground="#cbd5e1", highlightcolor="#3b82f6", bg='#ffffff')
        self.technique_text.pack(fill="x", pady=(0, 15))
        technique_scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=self.technique_text.yview)
        technique_scrollbar.pack(side="right", fill="y")
        self.technique_text.configure(yscrollcommand=technique_scrollbar.set)

        # Findings with rich text editing
        ttk.Label(content_frame, text="Findings:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Findings formatting toolbar
        findings_toolbar = ttk.Frame(content_frame)
        findings_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(findings_toolbar, text="B", font=("Segoe UI", 10, "bold"), width=3, 
                 command=lambda: self.toggle_bold(self.findings_text)).pack(side="left", padx=(0, 2))
        tk.Button(findings_toolbar, text="I", font=("Segoe UI", 10, "italic"), width=3,
                 command=lambda: self.toggle_italic(self.findings_text)).pack(side="left", padx=(0, 2))
        tk.Button(findings_toolbar, text="•", font=("Segoe UI", 10), width=3,
                 command=lambda: self.add_bullet(self.findings_text)).pack(side="left", padx=(0, 2))
        tk.Button(findings_toolbar, text="📋", font=("Segoe UI", 10), width=3,
                 command=lambda: self.paste_table(self.findings_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown for findings
        tk.Label(findings_toolbar, text="Size:", font=("Segoe UI", 9)).pack(side="left", padx=(10, 2))
        findings_size_var = tk.StringVar(value="11")
        findings_size_combo = ttk.Combobox(findings_toolbar, textvariable=findings_size_var, 
                                          values=["10", "11", "12", "14", "16"], width=5, state="readonly")
        findings_size_combo.pack(side="left", padx=(0, 10))
        findings_size_combo.bind('<<ComboboxSelected>>', 
                                lambda e: self.change_font_size(self.findings_text, findings_size_var.get()))
        
        self.findings_text = tk.Text(content_frame, height=6, font=("Segoe UI", 11), relief='solid', bd=1, wrap='word', 
                                   highlightbackground="#cbd5e1", highlightcolor="#3b82f6", bg='#ffffff')
        self.findings_text.pack(fill="both", expand=True, pady=(0, 12))
        findings_scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=self.findings_text.yview)
        findings_scrollbar.pack(side="right", fill="y")
        self.findings_text.configure(yscrollcommand=findings_scrollbar.set)

        # Impression with rich text editing
        ttk.Label(content_frame, text="Impression:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Impression formatting toolbar
        impression_toolbar = ttk.Frame(content_frame)
        impression_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(impression_toolbar, text="B", font=("Segoe UI", 10, "bold"), width=3, 
                 command=lambda: self.toggle_bold(self.impression_text)).pack(side="left", padx=(0, 2))
        tk.Button(impression_toolbar, text="I", font=("Segoe UI", 10, "italic"), width=3,
                 command=lambda: self.toggle_italic(self.impression_text)).pack(side="left", padx=(0, 2))
        tk.Button(impression_toolbar, text="•", font=("Segoe UI", 10), width=3,
                 command=lambda: self.add_bullet(self.impression_text)).pack(side="left", padx=(0, 2))
        tk.Button(impression_toolbar, text="📋", font=("Segoe UI", 10), width=3,
                 command=lambda: self.paste_table(self.impression_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown for impression
        tk.Label(impression_toolbar, text="Size:", font=("Segoe UI", 9)).pack(side="left", padx=(10, 2))
        impression_size_var = tk.StringVar(value="11")
        impression_size_combo = ttk.Combobox(impression_toolbar, textvariable=impression_size_var, 
                                            values=["10", "11", "12", "14", "16"], width=5, state="readonly")
        impression_size_combo.pack(side="left", padx=(0, 10))
        impression_size_combo.bind('<<ComboboxSelected>>', 
                                  lambda e: self.change_font_size(self.impression_text, impression_size_var.get()))
        
        self.impression_text = tk.Text(content_frame, height=3, font=("Segoe UI", 11), relief='solid', bd=1, wrap='word',
                                     highlightbackground="#cbd5e1", highlightcolor="#3b82f6", bg='#ffffff')
        self.impression_text.pack(fill="both", expand=True, pady=(0, 12))
        impression_scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=self.impression_text.yview)
        impression_scrollbar.pack(side="right", fill="y")
        self.impression_text.configure(yscrollcommand=impression_scrollbar.set)

        # Recommendations with rich text editing
        ttk.Label(content_frame, text="Recommendations:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Recommendations formatting toolbar
        recommendations_toolbar = ttk.Frame(content_frame)
        recommendations_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(recommendations_toolbar, text="B", font=("Segoe UI", 10, "bold"), width=3, 
                 command=lambda: self.toggle_bold(self.recommendations_text)).pack(side="left", padx=(0, 2))
        tk.Button(recommendations_toolbar, text="I", font=("Segoe UI", 10, "italic"), width=3,
                 command=lambda: self.toggle_italic(self.recommendations_text)).pack(side="left", padx=(0, 2))
        tk.Button(recommendations_toolbar, text="•", font=("Segoe UI", 10), width=3,
                 command=lambda: self.add_bullet(self.recommendations_text)).pack(side="left", padx=(0, 2))
        tk.Button(recommendations_toolbar, text="📋", font=("Segoe UI", 10), width=3,
                 command=lambda: self.paste_table(self.recommendations_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown for recommendations
        tk.Label(recommendations_toolbar, text="Size:", font=("Segoe UI", 9)).pack(side="left", padx=(10, 2))
        recommendations_size_var = tk.StringVar(value="11")
        recommendations_size_combo = ttk.Combobox(recommendations_toolbar, textvariable=recommendations_size_var, 
                                                 values=["10", "11", "12", "14", "16"], width=5, state="readonly")
        recommendations_size_combo.pack(side="left", padx=(0, 10))
        recommendations_size_combo.bind('<<ComboboxSelected>>', 
                                       lambda e: self.change_font_size(self.recommendations_text, recommendations_size_var.get()))
        
        self.recommendations_text = tk.Text(content_frame, height=2, font=("Segoe UI", 11), relief='solid', bd=1, wrap='word',
                                          highlightbackground="#cbd5e1", highlightcolor="#3b82f6", bg='#ffffff')
        self.recommendations_text.pack(fill="both", expand=True, pady=(0, 10))
        recommendations_scrollbar = ttk.Scrollbar(content_frame, orient="vertical", command=self.recommendations_text.yview)
        recommendations_scrollbar.pack(side="right", fill="y")
        self.recommendations_text.configure(yscrollcommand=recommendations_scrollbar.set)
        
        # Add Clear All button for report content
        clear_report_btn = ttk.Button(content_frame, text="🗑️ Clear Report Content", style='Danger.TButton', 
                                     command=self.clear_report_content)
        clear_report_btn.pack(pady=(10, 0))

    def create_preview_tab(self, parent):
        # Preview Section
        preview_frame = ttk.Labelframe(parent, text="Report Preview", style='Section.TLabelframe', padding=(20, 15))
        preview_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Preview text area with reduced height
        preview_text_frame = ttk.Frame(preview_frame)
        preview_text_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.preview_text = tk.Text(preview_text_frame, font=("Segoe UI", 11), relief='solid', bd=1, wrap='word',
                                  highlightbackground="#cbd5e1", highlightcolor="#3b82f6", bg='#ffffff', state='disabled', height=15)
        self.preview_text.pack(fill="both", expand=True, side="left")
        
        preview_scrollbar = ttk.Scrollbar(preview_text_frame, orient="vertical", command=self.preview_text.yview)
        preview_scrollbar.pack(side="right", fill="y")
        self.preview_text.configure(yscrollcommand=preview_scrollbar.set)

        # Simple responsive button layout
        button_container = ttk.Frame(preview_frame)
        button_container.pack(fill="x", pady=10, padx=20)
        
        # Create compact but visible buttons
        update_btn = tk.Button(button_container, 
                             text="🔄 Update Preview", 
                             font=("Segoe UI", 12, "bold"),
                             bg="#3b82f6", 
                             fg="white",
                             relief="solid",
                             bd=2,
                             padx=15,
                             pady=8,
                             cursor="hand2",
                             height=1,
                             command=self.update_preview)
        update_btn.pack(side="left", fill="both", expand=True, padx=(0, 10), pady=5)
        
        print_btn = tk.Button(button_container, 
                            text="🖨️ Print Report", 
                            font=("Segoe UI", 12, "bold"),
                            bg="#3b82f6", 
                            fg="white",
                            relief="solid",
                            bd=2,
                            padx=15,
                            pady=8,
                            cursor="hand2",
                            height=1,
                            command=self.print_report)
        print_btn.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        save_btn = tk.Button(button_container, 
                           text="💾 Save Report", 
                           font=("Segoe UI", 12, "bold"),
                           bg="#10b981", 
                           fg="white",
                           relief="solid",
                           bd=2,
                           padx=15,
                           pady=8,
                           cursor="hand2",
                           height=1,
                           command=self.save_report)
        save_btn.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=5)
        
        # Instructions
        instructions_label = ttk.Label(preview_frame, text="💡 Tip: Click 'Update Preview' to see your report before saving", 
                                     font=("Segoe UI", 10), foreground="#64748b", background='#f8fafc')
        instructions_label.pack(pady=(10, 0))

    def create_reports_tab(self, parent):
        """Create the View All Reports tab"""
        # Search and filter frame
        search_frame = ttk.Labelframe(parent, text="Search & Filter", style='Section.TLabelframe', padding=(15, 10))
        search_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        # Search row
        search_row = ttk.Frame(search_frame)
        search_row.pack(fill="x", pady=(0, 10))
        
        ttk.Label(search_row, text="Search:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(side="left", padx=(0, 10))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var, width=30, font=("Segoe UI", 11))
        self.search_entry.pack(side="left", padx=(0, 20))
        self.search_entry.bind('<KeyRelease>', self.filter_reports_list)
        
        # Filter by test type
        ttk.Label(search_row, text="Test Type:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(side="left", padx=(0, 10))
        self.filter_test_type = tk.StringVar()
        filter_values = ["All"] + self.admin_data["test_types"]
        self.filter_combo = ttk.Combobox(search_row, textvariable=self.filter_test_type, 
                                        values=filter_values, 
                                        state="readonly", width=15, font=("Segoe UI", 11))
        self.filter_combo.pack(side="left", padx=(0, 10))
        self.filter_combo.set("All")
        self.filter_combo.bind('<<ComboboxSelected>>', self.filter_reports_list)
        
        # Refresh button
        refresh_btn = ttk.Button(search_row, text="🔄 Refresh", style='Accent.TButton', command=self.refresh_reports_list)
        refresh_btn.pack(side="right", ipadx=10, ipady=5)
        
        # Main content area
        content_frame = ttk.Frame(parent)
        content_frame.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        
        # Left side - Reports list
        list_frame = ttk.Labelframe(content_frame, text="Reports", style='Section.TLabelframe', padding=(15, 10))
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Reports list with scrollbar
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill="both", expand=True)
        
        self.reports_list = tk.Listbox(list_container, font=("Segoe UI", 11), selectmode=tk.SINGLE, activestyle='dotbox', 
                                      bg='#ffffff', bd=1, highlightbackground="#cbd5e1", highlightcolor="#3b82f6")
        self.reports_list.pack(fill="both", expand=True, side="left")
        
        list_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.reports_list.yview)
        list_scrollbar.pack(side="right", fill="y")
        self.reports_list.configure(yscrollcommand=list_scrollbar.set)
        
        # Right side - Report details
        details_frame = ttk.Labelframe(content_frame, text="Report Details", style='Section.TLabelframe', padding=(8, 6))
        details_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # Report details text with minimal height
        self.details_text = tk.Text(details_frame, font=("Segoe UI", 8), relief='solid', bd=1, wrap='word',
                                   highlightbackground="#cbd5e1", highlightcolor="#3b82f6", bg='#ffffff', state='disabled', height=4)
        self.details_text.pack(fill="both", expand=True, pady=(0, 2))
        
        details_scrollbar = ttk.Scrollbar(details_frame, orient="vertical", command=self.details_text.yview)
        details_scrollbar.pack(side="right", fill="y")
        self.details_text.configure(yscrollcommand=details_scrollbar.set)
        
        # Action buttons with better layout
        action_frame = ttk.Frame(details_frame)
        action_frame.pack(fill="x", pady=(0, 3))
        
        self.open_btn = tk.Button(action_frame, text="📄 Open Report", 
                                font=("Segoe UI", 10, "bold"),
                                bg="#3b82f6", fg="white", relief="solid", bd=2,
                                padx=8, pady=4, cursor="hand2", height=1,
                                command=self.open_selected_report)
        self.open_btn.pack(side="left", fill="both", expand=True, padx=(0, 2), pady=2)
        
        self.print_btn = tk.Button(action_frame, text="🖨️ Print Report", 
                                 font=("Segoe UI", 10, "bold"),
                                 bg="#3b82f6", fg="white", relief="solid", bd=2,
                                 padx=8, pady=4, cursor="hand2", height=1,
                                 command=self.print_selected_report)
        self.print_btn.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        
        self.delete_btn = tk.Button(action_frame, text="🗑️ Delete Report", 
                                  font=("Segoe UI", 10, "bold"),
                                  bg="#ef4444", fg="white", relief="solid", bd=2,
                                  padx=8, pady=4, cursor="hand2", height=1,
                                  command=self.delete_selected_report)
        self.delete_btn.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=2)
        
        # Status label
        self.reports_status_label = ttk.Label(details_frame, text="No reports found", font=("Segoe UI", 8), 
                                            foreground="#64748b", background='#f8fafc')
        self.reports_status_label.pack(pady=(3, 0))
        
        # Bind selection event
        self.reports_list.bind('<<ListboxSelect>>', self.on_report_select)
        
        # Initialize variables
        self.filtered_reports = []
        self.all_reports = []
        
        # Load initial reports
        self.load_reports_list()

    def load_reports_list(self):
        """Load all reports into the list"""
        self.reports_list.delete(0, tk.END)
        self.filtered_reports = []
        
        try:
            report_files = [f for f in os.listdir(self.reports_dir) if f.endswith(".docx")]
            report_files.sort(reverse=True)  # Most recent first
            self.all_reports = report_files
            self.filtered_reports = report_files
            
            for f in report_files:
                self.reports_list.insert(tk.END, f)
            
            # Update status
            self.reports_status_label.config(text=f"Total Reports: {len(report_files)}")
            
        except Exception as e:
            self.reports_status_label.config(text=f"Error loading reports: {str(e)}")

    def filter_reports_list(self, event=None):
        """Filter reports based on search text and test type"""
        search_text = self.search_var.get().lower()
        test_type_filter = self.filter_test_type.get()
        
        self.reports_list.delete(0, tk.END)
        self.filtered_reports = []
        
        for filename in self.all_reports:
            # Check search text
            if search_text and search_text not in filename.lower():
                continue
                
            # Check test type filter
            if test_type_filter != "All":
                if test_type_filter not in filename:
                    continue
            
            self.reports_list.insert(tk.END, filename)
            self.filtered_reports.append(filename)
        
        # Update status
        self.reports_status_label.config(text=f"Showing {len(self.filtered_reports)} of {len(self.all_reports)} reports")

    def refresh_reports_list(self):
        """Refresh the reports list"""
        self.load_reports_list()
        self.filter_reports_list()

    def on_report_select(self, event=None):
        """Show details when a report is selected"""
        sel = self.reports_list.curselection()
        if not sel:
            return
            
        filename = self.filtered_reports[sel[0]]
        filepath = os.path.join(self.reports_dir, filename)
        
        try:
            # Get file info
            stat_info = os.stat(filepath)
            file_size = stat_info.st_size
            created_time = datetime.fromtimestamp(stat_info.st_ctime)
            modified_time = datetime.fromtimestamp(stat_info.st_mtime)
            
            # Parse filename for details
            name_parts = filename.replace('.docx', '').split('_')
            
            details = f"""📄 Report Details
{'='*30}

📁 Filename: {filename}
📏 Size: {file_size/1024:.1f} KB
📅 Created: {created_time.strftime('%B %d, %Y')}
👤 Patient Info:
"""
            
            # Try to extract patient info from filename
            if len(name_parts) >= 5:
                patient_id = name_parts[0].replace('_', ' ')
                patient_name = name_parts[1].replace('_', ' ')
                test_type = name_parts[2].replace('_', ' ')
                body_part = name_parts[3].replace('_', ' ')
                
                details += f"   • Patient ID: {patient_id}\n"
                details += f"   • Patient: {patient_name}\n"
                details += f"   • Test: {test_type}\n"
                details += f"   • Part: {body_part}\n"
            elif len(name_parts) >= 4:
                # Handle old format files (without Patient ID)
                patient_name = name_parts[0].replace('_', ' ')
                test_type = name_parts[1].replace('_', ' ')
                body_part = name_parts[2].replace('_', ' ')
                
                details += f"   • Patient: {patient_name}\n"
                details += f"   • Test: {test_type}\n"
                details += f"   • Part: {body_part}\n"
            else:
                details += "   • Info not available\n"
            
            details += f"""
💡 Use buttons below to open/delete
"""
            
            self.details_text.config(state='normal')
            self.details_text.delete("1.0", tk.END)
            self.details_text.insert("1.0", details)
            self.details_text.config(state='disabled')
            
        except Exception as e:
            self.details_text.config(state='normal')
            self.details_text.delete("1.0", tk.END)
            self.details_text.insert("1.0", f"Error loading details: {str(e)}")
            self.details_text.config(state='disabled')

    def open_selected_report(self):
        """Open the selected report"""
        sel = self.reports_list.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a report to open.")
            return
        filename = self.filtered_reports[sel[0]]
        filepath = os.path.abspath(os.path.join(self.reports_dir, filename))
        try:
            if sys.platform.startswith('darwin'):
                subprocess.call(('open', filepath))
            elif os.name == 'nt':
                os.startfile(filepath)
            elif os.name == 'posix':
                subprocess.call(('xdg-open', filepath))
            else:
                messagebox.showinfo("Info", f"File path: {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")

    def delete_selected_report(self):
        """Delete the selected report"""
        sel = self.reports_list.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Please select a report to delete.")
            return
        filename = self.filtered_reports[sel[0]]
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{filename}'?"):
            try:
                os.remove(os.path.join(self.reports_dir, filename))
                self.refresh_reports_list()  # Refresh the list
                self.details_text.config(state='normal')
                self.details_text.delete("1.0", tk.END)
                self.details_text.config(state='disabled')
                messagebox.showinfo("Success", f"Report '{filename}' deleted successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Could not delete file: {e}")

    def print_selected_report(self):
        """Print the selected report"""
        sel = self.reports_list.curselection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a report to print.")
            return
            
        filename = self.filtered_reports[sel[0]]
        filepath = os.path.join(self.reports_dir, filename)
        
        try:
            # Open the Word document with the default application for printing
            import subprocess
            import platform
            
            system = platform.system()
            
            if system == "Darwin":  # macOS
                # Use 'open' command to open with default application
                subprocess.run(['open', filepath])
                messagebox.showinfo("Print Report", f"Report '{filename}' opened in Word. Please use File > Print to print the document.")
                
            elif system == "Windows":
                # Use 'start' command to open with default application
                subprocess.run(['start', filepath], shell=True)
                messagebox.showinfo("Print Report", f"Report '{filename}' opened in Word. Please use File > Print to print the document.")
                
            elif system == "Linux":
                # Use 'xdg-open' to open with default application
                subprocess.run(['xdg-open', filepath])
                messagebox.showinfo("Print Report", f"Report '{filename}' opened in Word. Please use File > Print to print the document.")
                
        except Exception as e:
            messagebox.showerror("Print Error", f"Could not open report for printing: {str(e)}")

    def on_test_type_change(self, event=None):
        """Update technique field based on test type"""
        test_type = self.test_type_var.get()
        body_part = self.body_part_var.get()
        
        technique_templates = {
            "X-Ray": f"Standard radiographic examination of {body_part} with AP and lateral views",
            "CT Scan": f"Computed tomography examination of {body_part} with contrast enhancement",
            "MRI": f"Magnetic resonance imaging of {body_part} with T1, T2, and FLAIR sequences",
            "Ultrasound": f"Ultrasonographic examination of {body_part} using high-frequency transducer",
            "Mammography": "Digital mammographic examination with CC and MLO views",
            "Bone Density": "Dual-energy X-ray absorptiometry (DEXA) scan",
            "Other": f"Specialized radiological examination of {body_part}"
        }
        
        self.technique_text.delete(1.0, tk.END)
        self.technique_text.insert(1.0, technique_templates.get(test_type, ""))

    def on_body_part_change(self, event=None):
        """Update technique when body part changes"""
        self.on_test_type_change()

    def load_template(self):
        """Load predefined template based on test type and body part"""
        test_type = self.test_type_var.get()
        body_part = self.body_part_var.get()
        
        if test_type == "Select Test Type":
            messagebox.showwarning("Warning", "Please select a test type first.")
            return
            
        if body_part == "Select Body Part":
            messagebox.showwarning("Warning", "Please select a body part first.")
            return

        # Check if admin-defined requirements exist for this test type
        if test_type in self.admin_data["requirements"]:
            requirements = self.admin_data["requirements"][test_type]
            
            # Load admin-defined templates
            self.technique_text.delete(1.0, tk.END)
            self.technique_text.insert(1.0, requirements.get("technique", ""))
            
            self.findings_text.delete("1.0", tk.END)
            self.findings_text.insert("1.0", requirements.get("findings", ""))
            
            self.impression_text.delete("1.0", tk.END)
            self.impression_text.insert("1.0", requirements.get("impression", ""))
            
            self.recommendations_text.delete("1.0", tk.END)
            self.recommendations_text.insert("1.0", requirements.get("recommendations", ""))
            
            messagebox.showinfo("Template Loaded", f"Admin-defined template for {test_type} has been loaded. Please customize as needed.")
        else:
            # Fallback to generic template
            template = {
                "technique": f"Standard {test_type} examination of {body_part}",
                "findings": f"EXAMINATION: {test_type} of {body_part}\n\nFINDINGS:\n- Normal examination\n- No significant abnormality detected\n- Normal anatomical relationships\n- No evidence of acute pathology",
                "impression": f"Normal {test_type} examination of {body_part} with no significant abnormality identified.",
                "recommendations": "Clinical correlation recommended. Follow-up as clinically indicated."
            }
            
            # Load generic template
            self.technique_text.delete(1.0, tk.END)
            self.technique_text.insert(1.0, template["technique"])
            
            self.findings_text.delete("1.0", tk.END)
            self.findings_text.insert("1.0", template["findings"])
            
            self.impression_text.delete("1.0", tk.END)
            self.impression_text.insert("1.0", template["impression"])
            
            self.recommendations_text.delete("1.0", tk.END)
            self.recommendations_text.insert("1.0", template["recommendations"])
            
            messagebox.showinfo("Template Loaded", f"Generic template for {test_type} - {body_part} has been loaded. Please customize as needed.")

    def update_preview(self):
        """Update the preview tab with current report content"""
        self.preview_text.config(state='normal')
        self.preview_text.delete("1.0", tk.END)
        
        preview_content = f"""RADIOLOGY REPORT
{'='*50}

Radiologist: {self.radiologist_name}
Date: {datetime.now().strftime('%B %d, %Y')}
Time: {datetime.now().strftime('%I:%M %p')}

PATIENT INFORMATION
{'-'*20}
Patient ID: {self.patient_id_entry.get().strip()}
Name: {self.patient_name_entry.get().strip()}
Age: {self.age_entry.get().strip()}
Test Type: {self.test_type_var.get()}
Body Part: {self.body_part_var.get()}

TECHNIQUE
{'-'*10}
{self.technique_text.get("1.0", tk.END).strip()}

FINDINGS
{'-'*9}
{self.findings_text.get("1.0", tk.END).strip()}

IMPRESSION
{'-'*11}
{self.impression_text.get("1.0", tk.END).strip()}

RECOMMENDATIONS
{'-'*16}
{self.recommendations_text.get("1.0", tk.END).strip()}
"""
        
        self.preview_text.insert("1.0", preview_content)
        self.preview_text.config(state='disabled')

    def print_report(self):
        """Print the current report"""
        # Validate inputs first
        if not self.validate_inputs():
            return
        
        # Update preview to get current content
        self.update_preview()
        
        # Get the preview content
        preview_content = self.preview_text.get("1.0", tk.END)
        
        if not preview_content.strip():
            messagebox.showwarning("Warning", "No report content to print. Please fill in the report details first.")
            return
        
        try:
            # Create a temporary file for printing
            import tempfile
            import subprocess
            import platform
            
            # Create a temporary text file with the report content
            with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False, encoding='utf-8') as temp_file:
                temp_file.write(preview_content)
                temp_file_path = temp_file.name
            
            # Print based on operating system
            system = platform.system()
            
            if system == "Darwin":  # macOS
                # Use the 'lpr' command for printing
                try:
                    subprocess.run(['lpr', temp_file_path], check=True)
                    messagebox.showinfo("Print Success", "Report sent to printer successfully!")
                except subprocess.CalledProcessError:
                    # If lpr fails, try opening with default text editor for manual printing
                    subprocess.run(['open', temp_file_path])
                    messagebox.showinfo("Print Alternative", "Report opened in text editor. Please use File > Print to print manually.")
                except FileNotFoundError:
                    # If lpr not available, open in text editor
                    subprocess.run(['open', temp_file_path])
                    messagebox.showinfo("Print Alternative", "Report opened in text editor. Please use File > Print to print manually.")
                    
            elif system == "Windows":
                # For Windows, try using notepad with print command
                try:
                    subprocess.run(['notepad', '/p', temp_file_path], check=True)
                    messagebox.showinfo("Print Success", "Report sent to printer successfully!")
                except:
                    # Fallback: open in notepad for manual printing
                    subprocess.run(['notepad', temp_file_path])
                    messagebox.showinfo("Print Alternative", "Report opened in Notepad. Please use File > Print to print manually.")
                    
            elif system == "Linux":
                # For Linux, try using lpr
                try:
                    subprocess.run(['lpr', temp_file_path], check=True)
                    messagebox.showinfo("Print Success", "Report sent to printer successfully!")
                except:
                    # Fallback: open in default text editor
                    subprocess.run(['xdg-open', temp_file_path])
                    messagebox.showinfo("Print Alternative", "Report opened in text editor. Please use File > Print to print manually.")
            
            # Clean up temporary file after a delay
            import threading
            import time
            
            def cleanup_temp_file():
                time.sleep(5)  # Wait 5 seconds before cleanup
                try:
                    os.unlink(temp_file_path)
                except:
                    pass  # Ignore cleanup errors
            
            cleanup_thread = threading.Thread(target=cleanup_temp_file)
            cleanup_thread.daemon = True
            cleanup_thread.start()
            
        except Exception as e:
            messagebox.showerror("Print Error", f"Could not print report: {str(e)}\n\nPlease try saving the report and printing from the saved file.")

    def open_view_reports(self):
        view_win = tk.Toplevel(self.root)
        view_win.title("View & Search Reports")
        view_win.geometry("800x600")
        view_win.configure(bg='#f8fafc')
        
        # Header
        header_frame = ttk.Frame(view_win)
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        ttk.Label(header_frame, text="📁 Report Manager", font=("Segoe UI", 18, "bold"), background='#f8fafc', foreground="#1e293b").pack(side="left")
        
        # Search frame
        search_frame = ttk.Frame(view_win)
        search_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        ttk.Label(search_frame, text="Search:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(side="left", padx=(0, 10))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30, font=("Segoe UI", 11))
        search_entry.pack(side="left", padx=(0, 10))
        
        # Filter by test type
        ttk.Label(search_frame, text="Test Type:", font=("Segoe UI", 11, "bold"), foreground="#1e293b").pack(side="left", padx=(20, 10))
        filter_test_type = tk.StringVar()
        filter_values = ["All"] + self.admin_data["test_types"]
        filter_combo = ttk.Combobox(search_frame, textvariable=filter_test_type, 
                                   values=filter_values, 
                                   state="readonly", width=15, font=("Segoe UI", 11))
        filter_combo.pack(side="left", padx=(0, 10))
        filter_combo.set("All")
        
        # Main content area
        content_frame = ttk.Frame(view_win)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Left side - Reports list
        list_frame = ttk.Labelframe(content_frame, text="Reports", style='Section.TLabelframe', padding=(15, 10))
        list_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Reports list with scrollbar
        list_container = ttk.Frame(list_frame)
        list_container.pack(fill="both", expand=True)
        
        reports_list = tk.Listbox(list_container, font=("Segoe UI", 11), selectmode=tk.SINGLE, activestyle='dotbox', 
                                 bg='#ffffff', bd=1, highlightbackground="#cbd5e1", highlightcolor="#3b82f6")
        reports_list.pack(fill="both", expand=True, side="left")
        
        list_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=reports_list.yview)
        list_scrollbar.pack(side="right", fill="y")
        reports_list.configure(yscrollcommand=list_scrollbar.set)
        
        # Right side - Report details
        details_frame = ttk.Labelframe(content_frame, text="Report Details", style='Section.TLabelframe', padding=(15, 10))
        details_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # Report details text
        details_text = tk.Text(details_frame, font=("Segoe UI", 10), relief='solid', bd=1, wrap='word',
                              highlightbackground="#cbd5e1", highlightcolor="#3b82f6", bg='#ffffff', state='disabled')
        details_text.pack(fill="both", expand=True, pady=(0, 10))
        
        details_scrollbar = ttk.Scrollbar(details_frame, orient="vertical", command=details_text.yview)
        details_scrollbar.pack(side="right", fill="y")
        details_text.configure(yscrollcommand=details_scrollbar.set)
        
        # Initialize variables
        filtered_reports = []
        all_reports = []
        
        def load_reports():
            """Load all reports into the list"""
            reports_list.delete(0, tk.END)
            filtered_reports.clear()
            
            try:
                report_files = [f for f in os.listdir("Reports") if f.endswith(".docx")]
                report_files.sort(reverse=True)  # Most recent first
                all_reports.clear()
                all_reports.extend(report_files)
                filtered_reports.extend(report_files)
                
                for f in report_files:
                    reports_list.insert(tk.END, f)
                    
            except Exception as e:
                messagebox.showerror("Error", f"Could not load reports: {e}")

        def filter_reports(event=None):
            """Filter reports based on search text and test type"""
            search_text = search_var.get().lower()
            test_type_filter = filter_test_type.get()
            
            reports_list.delete(0, tk.END)
            filtered_reports.clear()
            
            for filename in all_reports:
                # Check search text
                if search_text and search_text not in filename.lower():
                    continue
                    
                # Check test type filter
                if test_type_filter != "All":
                    if test_type_filter not in filename:
                        continue
                
                reports_list.insert(tk.END, filename)
                filtered_reports.append(filename)

        def on_report_select(event=None):
            """Show details when a report is selected"""
            sel = reports_list.curselection()
            if not sel:
                return
                
            filename = filtered_reports[sel[0]]
            filepath = os.path.join("Reports", filename)
            
            try:
                # Get file info
                stat_info = os.stat(filepath)
                file_size = stat_info.st_size
                created_time = datetime.fromtimestamp(stat_info.st_ctime)
                modified_time = datetime.fromtimestamp(stat_info.st_mtime)
                
                # Parse filename for details
                name_parts = filename.replace('.docx', '').split('_')
                
                details = f"""📄 Report Details
{'='*40}

📁 Filename: {filename}
📏 Size: {file_size:,} bytes ({file_size/1024:.1f} KB)
📅 Created: {created_time.strftime('%B %d, %Y at %I:%M %p')}
🔄 Modified: {modified_time.strftime('%B %d, %Y at %I:%M %p')}

👤 Patient Information:
"""
                
                # Try to extract patient info from filename
                if len(name_parts) >= 5:
                    patient_id = name_parts[0].replace('_', ' ')
                    patient_name = name_parts[1].replace('_', ' ')
                    test_type = name_parts[2].replace('_', ' ')
                    body_part = name_parts[3].replace('_', ' ')
                    date_time = name_parts[4] if len(name_parts) > 4 else "Unknown"
                    
                    details += f"   • Patient ID: {patient_id}\n"
                    details += f"   • Patient: {patient_name}\n"
                    details += f"   • Test Type: {test_type}\n"
                    details += f"   • Body Part: {body_part}\n"
                    details += f"   • Date/Time: {date_time}\n"
                elif len(name_parts) >= 4:
                    # Handle old format files (without Patient ID)
                    patient_name = name_parts[0].replace('_', ' ')
                    test_type = name_parts[1].replace('_', ' ')
                    body_part = name_parts[2].replace('_', ' ')
                    date_time = name_parts[3] if len(name_parts) > 3 else "Unknown"
                    
                    details += f"   • Patient: {patient_name}\n"
                    details += f"   • Test Type: {test_type}\n"
                    details += f"   • Body Part: {body_part}\n"
                    details += f"   • Date/Time: {date_time}\n"
                else:
                    details += "   • Information not available\n"
                
                details += f"""
📂 File Path: {os.path.abspath(filepath)}

💡 Actions:
   • Double-click to open in Word
   • Use buttons below for more options
"""
                
                details_text.config(state='normal')
                details_text.delete("1.0", tk.END)
                details_text.insert("1.0", details)
                details_text.config(state='disabled')
                
            except Exception as e:
                details_text.config(state='normal')
                details_text.delete("1.0", tk.END)
                details_text.insert("1.0", f"Error loading details: {str(e)}")
                details_text.config(state='disabled')
        
        # Bind events
        search_entry.bind('<KeyRelease>', filter_reports)
        filter_combo.bind('<<ComboboxSelected>>', filter_reports)
        reports_list.bind('<<ListboxSelect>>', on_report_select)
        
        # Buttons frame
        button_frame = ttk.Frame(view_win)
        button_frame.pack(fill="x", padx=20, pady=(10, 20))
        
        def open_selected():
            sel = reports_list.curselection()
            if not sel:
                messagebox.showwarning("No selection", "Please select a report to open.", parent=view_win)
                return
            filename = filtered_reports[sel[0]]
            filepath = os.path.abspath(os.path.join("Reports", filename))
            try:
                if sys.platform.startswith('darwin'):
                    subprocess.call(('open', filepath))
                elif os.name == 'nt':
                    os.startfile(filepath)
                elif os.name == 'posix':
                    subprocess.call(('xdg-open', filepath))
                else:
                    messagebox.showinfo("Info", f"File path: {filepath}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open file: {e}", parent=view_win)
        
        def delete_selected():
            sel = reports_list.curselection()
            if not sel:
                messagebox.showwarning("No selection", "Please select a report to delete.", parent=view_win)
                return
            filename = filtered_reports[sel[0]]
            if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{filename}'?", parent=view_win):
                try:
                    os.remove(os.path.join("Reports", filename))
                    load_reports()  # Refresh the list
                    details_text.config(state='normal')
                    details_text.delete("1.0", tk.END)
                    details_text.config(state='disabled')
                    messagebox.showinfo("Success", f"Report '{filename}' deleted successfully.", parent=view_win)
                except Exception as e:
                    messagebox.showerror("Error", f"Could not delete file: {e}", parent=view_win)
        
        def refresh_reports():
            load_reports()
            messagebox.showinfo("Refresh", "Reports list refreshed.", parent=view_win)
        
        # Buttons
        open_btn = ttk.Button(button_frame, text="📄 Open Report", style='Accent.TButton', command=open_selected)
        open_btn.pack(side="left", ipadx=10, ipady=5)
        
        delete_btn = ttk.Button(button_frame, text="🗑️ Delete Report", style='Danger.TButton', command=delete_selected)
        delete_btn.pack(side="left", padx=(10, 0), ipadx=10, ipady=5)
        
        refresh_btn = ttk.Button(button_frame, text="🔄 Refresh", style='Accent.TButton', command=refresh_reports)
        refresh_btn.pack(side="right", ipadx=10, ipady=5)
        
        # Load initial reports
        load_reports()

    def validate_inputs(self):
        if not self.patient_id_entry.get().strip():
            messagebox.showerror("Error", "Please generate a patient ID first")
            return False
        if not self.patient_name_entry.get().strip():
            messagebox.showerror("Error", "Please enter patient name")
            return False
        if not self.age_entry.get().strip():
            messagebox.showerror("Error", "Please enter patient age")
            return False
        try:
            age = int(self.age_entry.get())
            if age <= 0 or age > 150:
                messagebox.showerror("Error", "Please enter a valid age (1-150)")
                return False
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid age (numbers only)")
            return False
        if self.test_type_var.get() == "Select Test Type":
            messagebox.showerror("Error", "Please select a test type")
            return False
        if self.body_part_var.get() == "Select Body Part":
            messagebox.showerror("Error", "Please select a body part")
            return False
        if not self.findings_text.get("1.0", tk.END).strip():
            messagebox.showerror("Error", "Please enter findings")
            return False
        if not self.impression_text.get("1.0", tk.END).strip():
            messagebox.showerror("Error", "Please enter impression")
            return False
        return True

    def save_report(self):
        if not self.validate_inputs():
            return
        try:
            # Create document without external templates to avoid PyInstaller issues
            doc = Document()
            
            # Set up document margins
            sections = doc.sections
            for section in sections:
                section.top_margin = Inches(1)
                section.bottom_margin = Inches(1)
                section.left_margin = Inches(1)
                section.right_margin = Inches(1)
            
            # Add simple header as first paragraph (avoiding template issues)
            header_para = doc.add_paragraph("NOOR CLINIC - RADIOLOGY DEPARTMENT")
            header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            header_run = header_para.runs[0]
            header_run.font.name = 'Arial'
            header_run.font.size = Pt(12)
            header_run.font.bold = True
            header_run.font.color.rgb = RGBColor(0, 51, 102)
            
            # Add clinic logo/name at top
            clinic_para = doc.add_paragraph()
            clinic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clinic_run = clinic_para.add_run("NOOR CLINIC")
            clinic_run.font.name = 'Arial'
            clinic_run.font.size = Pt(24)
            clinic_run.font.bold = True
            clinic_run.font.color.rgb = RGBColor(0, 51, 102)
            
            # Add main title
            title_para = doc.add_paragraph()
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            title_run = title_para.add_run("RADIOLOGY REPORT")
            title_run.font.name = 'Arial'
            title_run.font.size = Pt(20)
            title_run.font.bold = True
            title_run.font.color.rgb = RGBColor(0, 0, 0)
            
            # Add decorative line
            doc.add_paragraph("_" * 80).alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add report information in a table
            info_table = doc.add_table(rows=1, cols=2)
            info_table.style = 'Table Grid'
            info_table.autofit = True
            
            # Left column - Report details
            left_cell = info_table.cell(0, 0)
            left_cell.text = f"Report Date: {datetime.now().strftime('%B %d, %Y')}\nReport Time: {datetime.now().strftime('%I:%M %p')}\nRadiologist: {self.radiologist_name}"
            
            # Right column - Patient details
            right_cell = info_table.cell(0, 1)
            right_cell.text = f"Patient ID: {self.patient_id_entry.get().strip()}\nPatient Name: {self.patient_name_entry.get().strip()}\nAge: {self.age_entry.get().strip()}\nTest Type: {self.test_type_var.get()}\nBody Part: {self.body_part_var.get()}"
            
            # Style the table
            for row in info_table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        paragraph.style.font.name = 'Arial'
                        paragraph.style.font.size = Pt(11)
                        paragraph.style.font.color.rgb = RGBColor(0, 0, 0)
            
            doc.add_paragraph()  # Add spacing
            
            # Add Technique section
            technique_heading = doc.add_heading('TECHNIQUE', level=1)
            technique_heading.style.font.name = 'Arial'
            technique_heading.style.font.size = Pt(14)
            technique_heading.style.font.bold = True
            technique_heading.style.font.color.rgb = RGBColor(0, 51, 102)
            
            technique_para = doc.add_paragraph(self.technique_text.get("1.0", tk.END).strip())
            technique_para.style.font.name = 'Arial'
            technique_para.style.font.size = Pt(11)
            technique_para.style.font.color.rgb = RGBColor(0, 0, 0)
            
            doc.add_paragraph()  # Add spacing
            
            # Add Findings section
            findings_heading = doc.add_heading('FINDINGS', level=1)
            findings_heading.style.font.name = 'Arial'
            findings_heading.style.font.size = Pt(14)
            findings_heading.style.font.bold = True
            findings_heading.style.font.color.rgb = RGBColor(0, 51, 102)
            
            findings_content = self.findings_text.get("1.0", tk.END).strip()
            findings_para = doc.add_paragraph(findings_content)
            findings_para.style.font.name = 'Arial'
            findings_para.style.font.size = Pt(11)
            findings_para.style.font.color.rgb = RGBColor(0, 0, 0)
            
            doc.add_paragraph()  # Add spacing
            
            # Add Impression section
            impression_heading = doc.add_heading('IMPRESSION', level=1)
            impression_heading.style.font.name = 'Arial'
            impression_heading.style.font.size = Pt(14)
            impression_heading.style.font.bold = True
            impression_heading.style.font.color.rgb = RGBColor(0, 51, 102)
            
            impression_content = self.impression_text.get("1.0", tk.END).strip()
            impression_para = doc.add_paragraph(impression_content)
            impression_para.style.font.name = 'Arial'
            impression_para.style.font.size = Pt(11)
            impression_para.style.font.color.rgb = RGBColor(0, 0, 0)
            
            doc.add_paragraph()  # Add spacing
            
            # Add Recommendations section
            recommendations_heading = doc.add_heading('RECOMMENDATIONS', level=1)
            recommendations_heading.style.font.name = 'Arial'
            recommendations_heading.style.font.size = Pt(14)
            recommendations_heading.style.font.bold = True
            recommendations_heading.style.font.color.rgb = RGBColor(0, 51, 102)
            
            recommendations_content = self.recommendations_text.get("1.0", tk.END).strip()
            recommendations_para = doc.add_paragraph(recommendations_content)
            recommendations_para.style.font.name = 'Arial'
            recommendations_para.style.font.size = Pt(11)
            recommendations_para.style.font.color.rgb = RGBColor(0, 0, 0)
            
            doc.add_paragraph()  # Add spacing
            
            # Add signature section
            signature_para = doc.add_paragraph()
            signature_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            signature_run = signature_para.add_run(f"\n\n{self.radiologist_name}\nRadiologist\nNoor Clinic")
            signature_run.font.name = 'Arial'
            signature_run.font.size = Pt(11)
            signature_run.font.bold = True
            signature_run.font.color.rgb = RGBColor(0, 0, 0)
            
            # Add decorative line at bottom
            doc.add_paragraph("_" * 80).alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Add simple footer (avoiding template issues)
            footer_para = doc.add_paragraph(f"Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
            footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            footer_run = footer_para.runs[0]
            footer_run.font.name = 'Arial'
            footer_run.font.size = Pt(9)
            footer_run.font.color.rgb = RGBColor(128, 128, 128)
            
            # Generate filename
            patient_id_clean = self.patient_id_entry.get().strip().replace(" ", "_")
            patient_name_clean = self.patient_name_entry.get().strip().replace(" ", "_")
            test_type_clean = self.test_type_var.get().replace(" ", "_")
            body_part_clean = self.body_part_var.get().replace(" ", "_")
            date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(self.reports_dir, f"{patient_id_clean}_{patient_name_clean}_{test_type_clean}_{body_part_clean}_{date_str}.docx")
            
            doc.save(filename)
            messagebox.showinfo("Success", f"Report saved successfully!\nLocation: {filename}")
            self.status_label.config(text=f"Report saved: {os.path.basename(filename)}", foreground="#10b981")
            self.clear_form()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save report: {str(e)}")
            self.status_label.config(text="Error saving report", foreground="#ef4444")

    def clear_form(self):
        self.patient_id_entry.config(state='normal')
        self.patient_id_entry.delete(0, tk.END)
        self.patient_id_entry.config(state='readonly')
        self.patient_name_entry.delete(0, tk.END)
        self.age_entry.delete(0, tk.END)
        self.test_type_combo.set("Select Test Type")
        self.body_part_combo.set("Select Body Part")
        self.technique_text.delete(1.0, tk.END)
        self.findings_text.delete("1.0", tk.END)
        self.impression_text.delete("1.0", tk.END)
        self.recommendations_text.delete("1.0", tk.END)
        self.preview_text.config(state='normal')
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.config(state='disabled')
        self.status_label.config(text="Ready to create reports")

    def create_admin_tab(self, parent):
        """Create the Admin Settings tab"""
        # Admin header
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill="x", padx=20, pady=(20, 10))
        ttk.Label(header_frame, text="⚙️ Admin Settings", font=("Segoe UI", 16, "bold"), 
                 background='#f8fafc', foreground="#1e293b").pack(side="left")
        
        # Main content area
        content_frame = ttk.Frame(parent)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Left side - Test Types Management
        test_types_frame = ttk.Labelframe(content_frame, text="Test Types Management", style='Section.TLabelframe', padding=(15, 10))
        test_types_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Test types list
        list_frame = ttk.Frame(test_types_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        ttk.Label(list_frame, text="Current Test Types:", font=("Segoe UI", 11, "bold"), 
                 foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        self.test_types_listbox = tk.Listbox(list_frame, font=("Segoe UI", 11), selectmode=tk.SINGLE, 
                                            bg='#ffffff', bd=1, highlightbackground="#cbd5e1", highlightcolor="#3b82f6")
        self.test_types_listbox.pack(fill="both", expand=True, side="left")
        
        list_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.test_types_listbox.yview)
        list_scrollbar.pack(side="right", fill="y")
        self.test_types_listbox.configure(yscrollcommand=list_scrollbar.set)
        
        # Test type buttons
        test_buttons_frame = ttk.Frame(test_types_frame)
        test_buttons_frame.pack(fill="x", pady=(0, 10))
        
        add_test_btn = ttk.Button(test_buttons_frame, text="➕ Add Test Type", style='Success.TButton', 
                                 command=self.add_test_type)
        add_test_btn.pack(side="left", ipadx=10, ipady=5)
        
        edit_test_btn = ttk.Button(test_buttons_frame, text="✏️ Edit Test Type", style='Accent.TButton', 
                                  command=self.edit_test_type)
        edit_test_btn.pack(side="left", padx=(10, 0), ipadx=10, ipady=5)
        
        delete_test_btn = ttk.Button(test_buttons_frame, text="🗑️ Delete Test Type", style='Danger.TButton', 
                                    command=self.delete_test_type)
        delete_test_btn.pack(side="left", padx=(10, 0), ipadx=10, ipady=5)
        
        # Right side - Requirements Management
        requirements_frame = ttk.Labelframe(content_frame, text="Test Requirements", style='Section.TLabelframe', padding=(15, 10))
        requirements_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # Create scrollable frame for requirements
        canvas = tk.Canvas(requirements_frame, bg='#f8fafc')
        scrollbar = ttk.Scrollbar(requirements_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Requirements form
        form_frame = ttk.Frame(scrollable_frame)
        form_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Test type selection
        ttk.Label(form_frame, text="Selected Test Type:", font=("Segoe UI", 11, "bold"), 
                 foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        self.selected_test_type = tk.StringVar()
        self.test_type_combo_admin = ttk.Combobox(form_frame, textvariable=self.selected_test_type, 
                                                 state="readonly", font=("Segoe UI", 11))
        self.test_type_combo_admin.pack(fill="x", pady=(0, 15))
        self.test_type_combo_admin.bind('<<ComboboxSelected>>', self.load_test_requirements)
        
        # Requirements fields with rich text formatting
        ttk.Label(form_frame, text="Default Technique:", font=("Segoe UI", 11, "bold"), 
                 foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Technique field with formatting toolbar
        technique_frame = ttk.Frame(form_frame)
        technique_frame.pack(fill="x", pady=(0, 15))
        
        # Formatting toolbar for technique
        technique_toolbar = ttk.Frame(technique_frame)
        technique_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(technique_toolbar, text="B", font=("Segoe UI", 10, "bold"), 
                 command=lambda: self.toggle_bold(self.default_technique_text)).pack(side="left", padx=(0, 2))
        tk.Button(technique_toolbar, text="I", font=("Segoe UI", 10, "italic"), 
                 command=lambda: self.toggle_italic(self.default_technique_text)).pack(side="left", padx=(0, 2))
        tk.Button(technique_toolbar, text="•", font=("Segoe UI", 10), 
                 command=lambda: self.add_bullet(self.default_technique_text)).pack(side="left", padx=(0, 2))
        tk.Button(technique_toolbar, text="📋", font=("Segoe UI", 10), 
                 command=lambda: self.paste_table(self.default_technique_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown
        font_size_var = tk.StringVar(value="11")
        font_size_combo = ttk.Combobox(technique_toolbar, textvariable=font_size_var, 
                                      values=["8", "9", "10", "11", "12", "14", "16", "18"], 
                                      width=5, state="readonly")
        font_size_combo.pack(side="left", padx=(10, 0))
        font_size_combo.bind('<<ComboboxSelected>>', 
                           lambda e: self.change_font_size(self.default_technique_text, font_size_var.get()))
        
        self.default_technique_text = tk.Text(technique_frame, height=3, font=("Segoe UI", 11), 
                                            relief='solid', bd=1, wrap='word')
        self.default_technique_text.pack(fill="x")
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        ttk.Label(form_frame, text="Default Findings Template:", font=("Segoe UI", 11, "bold"), 
                 foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Findings field with formatting toolbar
        findings_frame = ttk.Frame(form_frame)
        findings_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # Formatting toolbar for findings
        findings_toolbar = ttk.Frame(findings_frame)
        findings_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(findings_toolbar, text="B", font=("Segoe UI", 10, "bold"), 
                 command=lambda: self.toggle_bold(self.default_findings_text)).pack(side="left", padx=(0, 2))
        tk.Button(findings_toolbar, text="I", font=("Segoe UI", 10, "italic"), 
                 command=lambda: self.toggle_italic(self.default_findings_text)).pack(side="left", padx=(0, 2))
        tk.Button(findings_toolbar, text="•", font=("Segoe UI", 10), 
                 command=lambda: self.add_bullet(self.default_findings_text)).pack(side="left", padx=(0, 2))
        tk.Button(findings_toolbar, text="📋", font=("Segoe UI", 10), 
                 command=lambda: self.paste_table(self.default_findings_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown for findings
        findings_font_size_var = tk.StringVar(value="11")
        findings_font_size_combo = ttk.Combobox(findings_toolbar, textvariable=findings_font_size_var, 
                                               values=["8", "9", "10", "11", "12", "14", "16", "18"], 
                                               width=5, state="readonly")
        findings_font_size_combo.pack(side="left", padx=(10, 0))
        findings_font_size_combo.bind('<<ComboboxSelected>>', 
                                    lambda e: self.change_font_size(self.default_findings_text, findings_font_size_var.get()))
        
        self.default_findings_text = tk.Text(findings_frame, height=6, font=("Segoe UI", 11), 
                                           relief='solid', bd=1, wrap='word')
        self.default_findings_text.pack(fill="both", expand=True)
        
        ttk.Label(form_frame, text="Default Impression Template:", font=("Segoe UI", 11, "bold"), 
                 foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Impression field with formatting toolbar
        impression_frame = ttk.Frame(form_frame)
        impression_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # Formatting toolbar for impression
        impression_toolbar = ttk.Frame(impression_frame)
        impression_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(impression_toolbar, text="B", font=("Segoe UI", 10, "bold"), 
                 command=lambda: self.toggle_bold(self.default_impression_text)).pack(side="left", padx=(0, 2))
        tk.Button(impression_toolbar, text="I", font=("Segoe UI", 10, "italic"), 
                 command=lambda: self.toggle_italic(self.default_impression_text)).pack(side="left", padx=(0, 2))
        tk.Button(impression_toolbar, text="•", font=("Segoe UI", 10), 
                 command=lambda: self.add_bullet(self.default_impression_text)).pack(side="left", padx=(0, 2))
        tk.Button(impression_toolbar, text="📋", font=("Segoe UI", 10), 
                 command=lambda: self.paste_table(self.default_impression_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown for impression
        impression_font_size_var = tk.StringVar(value="11")
        impression_font_size_combo = ttk.Combobox(impression_toolbar, textvariable=impression_font_size_var, 
                                                 values=["8", "9", "10", "11", "12", "14", "16", "18"], 
                                                 width=5, state="readonly")
        impression_font_size_combo.pack(side="left", padx=(10, 0))
        impression_font_size_combo.bind('<<ComboboxSelected>>', 
                                      lambda e: self.change_font_size(self.default_impression_text, impression_font_size_var.get()))
        
        self.default_impression_text = tk.Text(impression_frame, height=4, font=("Segoe UI", 11), 
                                             relief='solid', bd=1, wrap='word')
        self.default_impression_text.pack(fill="both", expand=True)
        
        ttk.Label(form_frame, text="Default Recommendations Template:", font=("Segoe UI", 11, "bold"), 
                 foreground="#1e293b").pack(anchor='w', pady=(0, 5))
        
        # Recommendations field with formatting toolbar
        recommendations_frame = ttk.Frame(form_frame)
        recommendations_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # Formatting toolbar for recommendations
        recommendations_toolbar = ttk.Frame(recommendations_frame)
        recommendations_toolbar.pack(fill="x", pady=(0, 5))
        
        tk.Button(recommendations_toolbar, text="B", font=("Segoe UI", 10, "bold"), 
                  command=lambda: self.toggle_bold(self.default_recommendations_text)).pack(side="left", padx=(0, 2))
        tk.Button(recommendations_toolbar, text="I", font=("Segoe UI", 10, "italic"), 
                  command=lambda: self.toggle_italic(self.default_recommendations_text)).pack(side="left", padx=(0, 2))
        tk.Button(recommendations_toolbar, text="•", font=("Segoe UI", 10), 
                  command=lambda: self.add_bullet(self.default_recommendations_text)).pack(side="left", padx=(0, 2))
        tk.Button(recommendations_toolbar, text="📋", font=("Segoe UI", 10), 
                  command=lambda: self.paste_table(self.default_recommendations_text)).pack(side="left", padx=(0, 2))
        
        # Font size dropdown for recommendations
        recommendations_font_size_var = tk.StringVar(value="11")
        recommendations_font_size_combo = ttk.Combobox(recommendations_toolbar, textvariable=recommendations_font_size_var, 
                                                      values=["8", "9", "10", "11", "12", "14", "16", "18"], 
                                                      width=5, state="readonly")
        recommendations_font_size_combo.pack(side="left", padx=(10, 0))
        recommendations_font_size_combo.bind('<<ComboboxSelected>>', 
                                           lambda e: self.change_font_size(self.default_recommendations_text, recommendations_font_size_var.get()))
        
        self.default_recommendations_text = tk.Text(recommendations_frame, height=3, font=("Segoe UI", 11), 
                                                  relief='solid', bd=1, wrap='word')
        self.default_recommendations_text.pack(fill="both", expand=True)
        
        # Save requirements button
        save_req_btn = ttk.Button(form_frame, text="💾 Save Requirements", style='Success.TButton', 
                                 command=self.save_test_requirements)
        save_req_btn.pack(ipadx=10, ipady=5)
        
        # Initialize admin data
        self.load_admin_data()
        self.load_test_types_list()

    def load_admin_data(self):
        """Load admin data from JSON file"""
        self.admin_file = "admin_settings.json"
        if os.path.exists(self.admin_file):
            try:
                with open(self.admin_file, 'r') as f:
                    self.admin_data = json.load(f)
            except:
                self.admin_data = {"test_types": [], "requirements": {}}
        else:
            self.admin_data = {"test_types": [], "requirements": {}}
            # Add default test types
            default_types = ["X-Ray", "CT Scan", "MRI", "Ultrasound", "Mammography", "Bone Density", "Other"]
            for test_type in default_types:
                self.admin_data["test_types"].append(test_type)
                self.admin_data["requirements"][test_type] = {
                    "technique": f"Standard {test_type} examination",
                    "findings": f"EXAMINATION: {test_type}\n\nFINDINGS:\n- Normal examination\n- No significant abnormality detected",
                    "impression": f"Normal {test_type} examination with no significant abnormality identified.",
                    "recommendations": "Clinical correlation recommended. Follow-up as clinically indicated."
                }
            self.save_admin_data()

    def save_admin_data(self):
        """Save admin data to JSON file"""
        try:
            with open(self.admin_file, 'w') as f:
                json.dump(self.admin_data, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save admin data: {e}")

    def load_test_types_list(self):
        """Load test types into the listbox"""
        self.test_types_listbox.delete(0, tk.END)
        for test_type in self.admin_data["test_types"]:
            self.test_types_listbox.insert(tk.END, test_type)
        
        # Update combobox in main app
        if hasattr(self, 'test_type_combo'):
            self.test_type_combo['values'] = self.admin_data["test_types"]
        
        # Update admin combobox
        self.test_type_combo_admin['values'] = self.admin_data["test_types"]

    def add_test_type(self):
        """Add a new test type"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Test Type")
        dialog.geometry("400x200")
        dialog.configure(bg='#f8fafc')
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Enter Test Type Name:", font=("Segoe UI", 12, "bold"), 
                 background='#f8fafc', foreground="#1e293b").pack(pady=20)
        
        test_type_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=test_type_var, font=("Segoe UI", 11), width=30)
        entry.pack(pady=(0, 20))
        entry.focus()
        
        def save_test_type():
            test_type = test_type_var.get().strip()
            if not test_type:
                messagebox.showwarning("Warning", "Please enter a test type name.", parent=dialog)
                return
            if test_type in self.admin_data["test_types"]:
                messagebox.showwarning("Warning", "Test type already exists.", parent=dialog)
                return
            
            self.admin_data["test_types"].append(test_type)
            self.admin_data["requirements"][test_type] = {
                "technique": f"Standard {test_type} examination",
                "findings": f"EXAMINATION: {test_type}\n\nFINDINGS:\n- Normal examination\n- No significant abnormality detected",
                "impression": f"Normal {test_type} examination with no significant abnormality identified.",
                "recommendations": "Clinical correlation recommended. Follow-up as clinically indicated."
            }
            self.save_admin_data()
            self.load_test_types_list()
            dialog.destroy()
            messagebox.showinfo("Success", f"Test type '{test_type}' added successfully.")
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        save_btn = ttk.Button(button_frame, text="Save", style='Success.TButton', command=save_test_type)
        save_btn.pack(side="left", padx=(0, 10), ipadx=15, ipady=5)
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", style='Danger.TButton', command=dialog.destroy)
        cancel_btn.pack(side="left", ipadx=15, ipady=5)

    def edit_test_type(self):
        """Edit selected test type"""
        selection = self.test_types_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a test type to edit.")
            return
        
        old_name = self.test_types_listbox.get(selection[0])
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Test Type")
        dialog.geometry("400x200")
        dialog.configure(bg='#f8fafc')
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Edit Test Type Name:", font=("Segoe UI", 12, "bold"), 
                 background='#f8fafc', foreground="#1e293b").pack(pady=20)
        
        test_type_var = tk.StringVar(value=old_name)
        entry = ttk.Entry(dialog, textvariable=test_type_var, font=("Segoe UI", 11), width=30)
        entry.pack(pady=(0, 20))
        entry.focus()
        
        def save_changes():
            new_name = test_type_var.get().strip()
            if not new_name:
                messagebox.showwarning("Warning", "Please enter a test type name.", parent=dialog)
                return
            if new_name != old_name and new_name in self.admin_data["test_types"]:
                messagebox.showwarning("Warning", "Test type already exists.", parent=dialog)
                return
            
            # Update test type name
            index = self.admin_data["test_types"].index(old_name)
            self.admin_data["test_types"][index] = new_name
            
            # Update requirements key
            if old_name in self.admin_data["requirements"]:
                self.admin_data["requirements"][new_name] = self.admin_data["requirements"].pop(old_name)
            
            self.save_admin_data()
            self.load_test_types_list()
            dialog.destroy()
            messagebox.showinfo("Success", f"Test type updated successfully.")
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        save_btn = ttk.Button(button_frame, text="Save", style='Success.TButton', command=save_changes)
        save_btn.pack(side="left", padx=(0, 10), ipadx=15, ipady=5)
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", style='Danger.TButton', command=dialog.destroy)
        cancel_btn.pack(side="left", ipadx=15, ipady=5)

    def delete_test_type(self):
        """Delete selected test type"""
        selection = self.test_types_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a test type to delete.")
            return
        
        test_type = self.test_types_listbox.get(selection[0])
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{test_type}'?\n\nThis will also delete all associated requirements."):
            self.admin_data["test_types"].remove(test_type)
            if test_type in self.admin_data["requirements"]:
                del self.admin_data["requirements"][test_type]
            self.save_admin_data()
            self.load_test_types_list()
            messagebox.showinfo("Success", f"Test type '{test_type}' deleted successfully.")

    def load_test_requirements(self, event=None):
        """Load requirements for selected test type"""
        test_type = self.selected_test_type.get()
        if not test_type or test_type not in self.admin_data["requirements"]:
            return
        
        requirements = self.admin_data["requirements"][test_type]
        
        self.default_technique_text.delete("1.0", tk.END)
        self.default_technique_text.insert("1.0", requirements.get("technique", ""))
        
        self.default_findings_text.delete("1.0", tk.END)
        self.default_findings_text.insert("1.0", requirements.get("findings", ""))
        
        self.default_impression_text.delete("1.0", tk.END)
        self.default_impression_text.insert("1.0", requirements.get("impression", ""))
        
        self.default_recommendations_text.delete("1.0", tk.END)
        self.default_recommendations_text.insert("1.0", requirements.get("recommendations", ""))

    def save_test_requirements(self):
        """Save requirements for selected test type"""
        test_type = self.selected_test_type.get()
        if not test_type:
            messagebox.showwarning("Warning", "Please select a test type first.")
            return
        
        requirements = {
            "technique": self.default_technique_text.get("1.0", tk.END).strip(),
            "findings": self.default_findings_text.get("1.0", tk.END).strip(),
            "impression": self.default_impression_text.get("1.0", tk.END).strip(),
            "recommendations": self.default_recommendations_text.get("1.0", tk.END).strip()
        }
        
        self.admin_data["requirements"][test_type] = requirements
        self.save_admin_data()
        messagebox.showinfo("Success", f"Requirements for '{test_type}' saved successfully.")

    # Rich Text Formatting Functions
    def toggle_bold(self, text_widget):
        """Toggle bold formatting for selected text"""
        try:
            # Get current selection
            sel_start = text_widget.index("sel.first")
            sel_end = text_widget.index("sel.last")
            
            # Get current text
            selected_text = text_widget.get(sel_start, sel_end)
            
            # Toggle bold (add/remove ** markers)
            if selected_text.startswith("**") and selected_text.endswith("**"):
                # Remove bold
                new_text = selected_text[2:-2]
            else:
                # Add bold
                new_text = f"**{selected_text}**"
            
            # Replace the text
            text_widget.delete(sel_start, sel_end)
            text_widget.insert(sel_start, new_text)
            
        except tk.TclError:
            # No selection, insert bold markers at cursor
            text_widget.insert("insert", "**")

    def toggle_italic(self, text_widget):
        """Toggle italic formatting for selected text"""
        try:
            # Get current selection
            sel_start = text_widget.index("sel.first")
            sel_end = text_widget.index("sel.last")
            
            # Get current text
            selected_text = text_widget.get(sel_start, sel_end)
            
            # Toggle italic (add/remove * markers)
            if selected_text.startswith("*") and selected_text.endswith("*") and not selected_text.startswith("**"):
                # Remove italic
                new_text = selected_text[1:-1]
            else:
                # Add italic
                new_text = f"*{selected_text}*"
            
            # Replace the text
            text_widget.delete(sel_start, sel_end)
            text_widget.insert(sel_start, new_text)
            
        except tk.TclError:
            # No selection, insert italic markers at cursor
            text_widget.insert("insert", "*")

    def add_bullet(self, text_widget):
        """Add bullet point at cursor position"""
        current_line = text_widget.index("insert linestart")
        text_widget.insert(current_line, "• ")

    def paste_table(self, text_widget):
        """Insert a table template"""
        table_template = """
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
| Data 4   | Data 5   | Data 6   |
"""
        text_widget.insert("insert", table_template)

    def change_font_size(self, text_widget, size):
        """Change font size for the text widget"""
        try:
            size = int(size)
            text_widget.configure(font=("Segoe UI", size))
        except ValueError:
            pass

    def clear_report_content(self):
        self.technique_text.delete(1.0, tk.END)
        self.findings_text.delete("1.0", tk.END)
        self.impression_text.delete("1.0", tk.END)
        self.recommendations_text.delete("1.0", tk.END)
        self.preview_text.config(state='normal')
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.config(state='disabled')
        self.status_label.config(text="Ready to create reports")

    def clear_patient_info(self):
        """Clear patient information fields"""
        self.patient_name_entry.delete(0, tk.END)
        self.age_entry.delete(0, tk.END)
        self.test_type_combo.set("Select Test Type")
        self.body_part_combo.set("Select Body Part")
        self.generate_new_patient_id()  # Generate new patient ID
        self.status_label.config(text="Patient information cleared", foreground="#10b981")


def main():
    root = tk.Tk()
    app = RadiologistApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 