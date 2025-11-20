import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import sqlite3
from datetime import datetime
import os
import csv
from PIL import Image, ImageTk  # Now Pillow is properly installed

# Set appearance mode and color theme
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ContactManagementSystem:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("Modern Contact Management System")
        self.root.geometry("1200x700")
        self.root.resizable(True, True)
        
        # Center the window on screen
        self.center_window()
        
        # Initialize database
        self.init_database()
        
        # Setup GUI
        self.setup_gui()
        
        # Load contacts
        self.load_contacts()
        
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def init_database(self):
        """Initialize SQLite database"""
        self.conn = sqlite3.connect('contacts.db', check_same_thread=False)
        self.cursor = self.conn.cursor()
        
        # Create contacts table if not exists
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                first_name TEXT NOT NULL,
                last_name TEXT NOT NULL,
                phone TEXT,
                email TEXT,
                address TEXT,
                company TEXT,
                notes TEXT,
                category TEXT,
                created_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_modified TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        self.conn.commit()
    
    def setup_gui(self):
        """Setup the main GUI components"""
        # Configure grid layout
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        
        # Create sidebar frame
        self.sidebar_frame = ctk.CTkFrame(self.root, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)
        
        # Sidebar widgets
        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Contact Manager", 
                                      font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Navigation buttons
        self.dashboard_btn = ctk.CTkButton(self.sidebar_frame, text="📊 Dashboard", 
                                          command=self.show_dashboard)
        self.dashboard_btn.grid(row=1, column=0, padx=20, pady=10)
        
        self.contacts_btn = ctk.CTkButton(self.sidebar_frame, text="👥 All Contacts", 
                                         command=self.show_contacts)
        self.contacts_btn.grid(row=2, column=0, padx=20, pady=10)
        
        self.add_contact_btn = ctk.CTkButton(self.sidebar_frame, text="➕ Add Contact", 
                                            command=self.show_add_contact)
        self.add_contact_btn.grid(row=3, column=0, padx=20, pady=10)
        
        self.import_export_btn = ctk.CTkButton(self.sidebar_frame, text="📁 Import/Export", 
                                              command=self.show_import_export)
        self.import_export_btn.grid(row=4, column=0, padx=20, pady=10)
        
        # Appearance mode
        self.appearance_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_label.grid(row=7, column=0, padx=20, pady=(10, 0))
        self.appearance_mode = ctk.CTkOptionMenu(self.sidebar_frame, 
                                                values=["Light", "Dark", "System"],
                                                command=self.change_appearance_mode)
        self.appearance_mode.grid(row=8, column=0, padx=20, pady=(10, 20))
        
        # Main content area
        self.main_frame = ctk.CTkFrame(self.root, corner_radius=0)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Create different pages
        self.create_dashboard_page()
        self.create_contacts_page()
        self.create_add_contact_page()
        self.create_import_export_page()
        
        # Show dashboard by default
        self.show_dashboard()
    
    def create_dashboard_page(self):
        """Create dashboard page"""
        self.dashboard_page = ctk.CTkFrame(self.main_frame)
        
        # Dashboard content
        title_label = ctk.CTkLabel(self.dashboard_page, text="📊 Dashboard", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=20)
        
        # Stats frame
        stats_frame = ctk.CTkFrame(self.dashboard_page)
        stats_frame.pack(pady=20, padx=20, fill="x")
        
        # Get stats
        total_contacts = self.get_total_contacts()
        recent_contacts = self.get_recent_contacts()
        categories_count = self.get_categories_count()
        
        stats_text = f"""
        Welcome to Contact Management System!
        
        📈 Statistics:
        • Total Contacts: {total_contacts}
        • Recent Contacts (Last 7 days): {recent_contacts}
        • Categories: {categories_count}
        
        🚀 Quick Actions:
        • Add new contact
        • View all contacts  
        • Search contacts
        • Import/Export data
        """
        
        stats_label = ctk.CTkLabel(stats_frame, text=stats_text, 
                                  font=ctk.CTkFont(size=14), justify="left")
        stats_label.pack(pady=20, padx=20)
        
        # Quick action buttons
        action_frame = ctk.CTkFrame(self.dashboard_page)
        action_frame.pack(pady=20, padx=20, fill="x")
        
        quick_actions_label = ctk.CTkLabel(action_frame, text="Quick Actions:", 
                                          font=ctk.CTkFont(size=16, weight="bold"))
        quick_actions_label.pack(pady=10)
        
        button_frame = ctk.CTkFrame(action_frame)
        button_frame.pack(pady=10)
        
        ctk.CTkButton(button_frame, text="Add New Contact", 
                     command=self.show_add_contact).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="View All Contacts", 
                     command=self.show_contacts).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="Export Contacts", 
                     command=self.export_contacts).pack(side="left", padx=10)
    
    def create_contacts_page(self):
        """Create contacts list page"""
        self.contacts_page = ctk.CTkFrame(self.main_frame)
        
        # Header
        header_frame = ctk.CTkFrame(self.contacts_page)
        header_frame.pack(fill="x", padx=20, pady=10)
        
        title_label = ctk.CTkLabel(header_frame, text="👥 All Contacts", 
                                  font=ctk.CTkFont(size=20, weight="bold"))
        title_label.pack(side="left", padx=10, pady=10)
        
        # Search and filter frame
        search_frame = ctk.CTkFrame(self.contacts_page)
        search_frame.pack(fill="x", padx=20, pady=10)
        
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="🔍 Search contacts by name, phone, email...")
        self.search_entry.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        self.search_entry.bind("<KeyRelease>", self.search_contacts)
        
        self.refresh_btn = ctk.CTkButton(search_frame, text="🔄 Refresh", 
                                        command=self.load_contacts)
        self.refresh_btn.pack(side="right", padx=10, pady=10)
        
        # Contacts table frame
        table_frame = ctk.CTkFrame(self.contacts_page)
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Create treeview with style
        style = ttk.Style()
        style.theme_use("clam")
        
        columns = ("ID", "Name", "Phone", "Email", "Company", "Category")
        self.contacts_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        
        # Configure columns
        column_widths = {"ID": 50, "Name": 150, "Phone": 120, "Email": 200, "Company": 150, "Category": 100}
        for col in columns:
            self.contacts_tree.heading(col, text=col)
            self.contacts_tree.column(col, width=column_widths[col])
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.contacts_tree.yview)
        self.contacts_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack treeview and scrollbar
        self.contacts_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind double click event
        self.contacts_tree.bind("<Double-1>", self.on_contact_double_click)
        
        # Action buttons frame
        action_frame = ctk.CTkFrame(self.contacts_page)
        action_frame.pack(fill="x", padx=20, pady=10)
        
        self.edit_btn = ctk.CTkButton(action_frame, text="✏️ Edit Contact", 
                                     command=self.edit_contact)
        self.edit_btn.pack(side="left", padx=10)
        
        self.delete_btn = ctk.CTkButton(action_frame, text="🗑️ Delete Contact", 
                                       command=self.delete_contact, fg_color="#d13438")
        self.delete_btn.pack(side="left", padx=10)
        
        self.export_btn = ctk.CTkButton(action_frame, text="📤 Export Selected", 
                                       command=self.export_selected_contacts)
        self.export_btn.pack(side="left", padx=10)
        
        # Status label
        self.status_label = ctk.CTkLabel(self.contacts_page, text="")
        self.status_label.pack(pady=5)
    
    def create_add_contact_page(self):
        """Create add/edit contact page"""
        self.add_contact_page = ctk.CTkFrame(self.main_frame)
        
        # Form frame with scrollbar
        main_form_frame = ctk.CTkFrame(self.add_contact_page)
        main_form_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        self.contact_form_title = ctk.CTkLabel(main_form_frame, text="➕ Add New Contact", 
                                              font=ctk.CTkFont(size=20, weight="bold"))
        self.contact_form_title.pack(pady=20)
        
        # Form container
        form_container = ctk.CTkFrame(main_form_frame)
        form_container.pack(fill="both", expand=True, padx=20)
        
        # Form fields
        fields = [
            ("First Name*", "first_name"),
            ("Last Name*", "last_name"),
            ("Phone", "phone"),
            ("Email", "email"),
            ("Company", "company"),
            ("Category", "category")
        ]
        
        self.entry_widgets = {}
        
        for i, (label, field) in enumerate(fields):
            row_frame = ctk.CTkFrame(form_container)
            row_frame.pack(fill="x", padx=10, pady=8)
            
            lbl = ctk.CTkLabel(row_frame, text=label, width=120, anchor="e")
            lbl.pack(side="left", padx=10, pady=5)
            
            entry = ctk.CTkEntry(row_frame, width=300)
            entry.pack(side="left", padx=10, pady=5, fill="x", expand=True)
            self.entry_widgets[field] = entry
        
        # Address (text box)
        address_frame = ctk.CTkFrame(form_container)
        address_frame.pack(fill="x", padx=10, pady=8)
        
        address_lbl = ctk.CTkLabel(address_frame, text="Address:", width=120, anchor="e")
        address_lbl.pack(side="left", padx=10, pady=5)
        
        self.address_text = ctk.CTkTextbox(address_frame, width=300, height=60)
        self.address_text.pack(side="left", padx=10, pady=5, fill="x", expand=True)
        
        # Notes (text box)
        notes_frame = ctk.CTkFrame(form_container)
        notes_frame.pack(fill="x", padx=10, pady=8)
        
        notes_lbl = ctk.CTkLabel(notes_frame, text="Notes:", width=120, anchor="e")
        notes_lbl.pack(side="left", padx=10, pady=5)
        
        self.notes_text = ctk.CTkTextbox(notes_frame, width=300, height=80)
        self.notes_text.pack(side="left", padx=10, pady=5, fill="x", expand=True)
        
        # Buttons
        button_frame = ctk.CTkFrame(form_container)
        button_frame.pack(fill="x", padx=10, pady=20)
        
        self.save_btn = ctk.CTkButton(button_frame, text="💾 Save Contact", 
                                     command=self.save_contact)
        self.save_btn.pack(side="left", padx=10)
        
        self.clear_btn = ctk.CTkButton(button_frame, text="🗑️ Clear Form", 
                                      command=self.clear_form)
        self.clear_btn.pack(side="left", padx=10)
        
        self.cancel_btn = ctk.CTkButton(button_frame, text="❌ Cancel", 
                                       command=self.show_contacts)
        self.cancel_btn.pack(side="left", padx=10)
        
        self.editing_id = None
    
    def create_import_export_page(self):
        """Create import/export page"""
        self.import_export_page = ctk.CTkFrame(self.main_frame)
        
        # Header
        header_frame = ctk.CTkFrame(self.import_export_page)
        header_frame.pack(fill="x", padx=20, pady=20)
        
        title_label = ctk.CTkLabel(header_frame, text="📁 Import/Export Data", 
                                  font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=10)
        
        # Export section
        export_frame = ctk.CTkFrame(self.import_export_page)
        export_frame.pack(fill="x", padx=20, pady=10)
        
        export_label = ctk.CTkLabel(export_frame, text="Export Contacts", 
                                   font=ctk.CTkFont(size=16, weight="bold"))
        export_label.pack(pady=10)
        
        export_desc = ctk.CTkLabel(export_frame, 
                                  text="Export your contacts to CSV format for backup or use in other applications.")
        export_desc.pack(pady=5)
        
        export_btn_frame = ctk.CTkFrame(export_frame)
        export_btn_frame.pack(pady=10)
        
        ctk.CTkButton(export_btn_frame, text="📤 Export All Contacts", 
                     command=self.export_contacts).pack(side="left", padx=10)
        ctk.CTkButton(export_btn_frame, text="📊 Export with Details", 
                     command=self.export_detailed_contacts).pack(side="left", padx=10)
        
        # Import section
        import_frame = ctk.CTkFrame(self.import_export_page)
        import_frame.pack(fill="x", padx=20, pady=10)
        
        import_label = ctk.CTkLabel(import_frame, text="Import Contacts", 
                                   font=ctk.CTkFont(size=16, weight="bold"))
        import_label.pack(pady=10)
        
        import_desc = ctk.CTkLabel(import_frame, 
                                  text="Import contacts from CSV file. File should have columns: first_name, last_name, phone, email, company, category, address, notes")
        import_desc.pack(pady=5)
        
        import_btn_frame = ctk.CTkFrame(import_frame)
        import_btn_frame.pack(pady=10)
        
        ctk.CTkButton(import_btn_frame, text="📥 Import from CSV", 
                     command=self.import_contacts).pack(side="left", padx=10)
        
        # Database section
        db_frame = ctk.CTkFrame(self.import_export_page)
        db_frame.pack(fill="x", padx=20, pady=10)
        
        db_label = ctk.CTkLabel(db_frame, text="Database Management", 
                               font=ctk.CTkFont(size=16, weight="bold"))
        db_label.pack(pady=10)
        
        db_btn_frame = ctk.CTkFrame(db_frame)
        db_btn_frame.pack(pady=10)
        
        ctk.CTkButton(db_btn_frame, text="🗃️ Backup Database", 
                     command=self.backup_database).pack(side="left", padx=10)
        ctk.CTkButton(db_btn_frame, text="🔄 Reset Database", 
                     command=self.reset_database, fg_color="#d13438").pack(side="left", padx=10)
    
    def show_dashboard(self):
        """Show dashboard page"""
        self.hide_all_pages()
        self.dashboard_page.pack(fill="both", expand=True)
        self.update_dashboard_stats()
    
    def show_contacts(self):
        """Show contacts page"""
        self.hide_all_pages()
        self.contacts_page.pack(fill="both", expand=True)
        self.load_contacts()
    
    def show_add_contact(self):
        """Show add contact page"""
        self.hide_all_pages()
        self.add_contact_page.pack(fill="both", expand=True)
        self.clear_form()
        self.contact_form_title.configure(text="➕ Add New Contact")
        self.editing_id = None
    
    def show_import_export(self):
        """Show import/export page"""
        self.hide_all_pages()
        self.import_export_page.pack(fill="both", expand=True)
    
    def hide_all_pages(self):
        """Hide all pages"""
        for page in [self.dashboard_page, self.contacts_page, 
                    self.add_contact_page, self.import_export_page]:
            page.pack_forget()
    
    def update_dashboard_stats(self):
        """Update dashboard statistics"""
        total_contacts = self.get_total_contacts()
        recent_contacts = self.get_recent_contacts()
        categories_count = self.get_categories_count()
        
        # Update would be done when we refresh the dashboard
        # For now, we'll just reload the dashboard
        pass
    
    def load_contacts(self):
        """Load contacts into treeview"""
        # Clear existing items
        for item in self.contacts_tree.get_children():
            self.contacts_tree.delete(item)
        
        # Fetch contacts from database
        self.cursor.execute('''
            SELECT id, first_name, last_name, phone, email, company, category 
            FROM contacts 
            ORDER BY first_name, last_name
        ''')
        contacts = self.cursor.fetchall()
        
        # Insert into treeview
        for contact in contacts:
            full_name = f"{contact[1]} {contact[2]}"
            self.contacts_tree.insert("", "end", values=(
                contact[0], full_name, contact[3] or "-", contact[4] or "-", 
                contact[5] or "-", contact[6] or "-"
            ))
        
        # Update status
        self.status_label.configure(text=f"Loaded {len(contacts)} contacts")
    
    def search_contacts(self, event=None):
        """Search contacts based on search term"""
        search_term = self.search_entry.get().lower()
        
        # Clear existing items
        for item in self.contacts_tree.get_children():
            self.contacts_tree.delete(item)
        
        if not search_term:
            self.load_contacts()
            return
        
        # Fetch and filter contacts
        self.cursor.execute('''
            SELECT id, first_name, last_name, phone, email, company, category 
            FROM contacts 
            WHERE LOWER(first_name) LIKE ? OR LOWER(last_name) LIKE ? OR phone LIKE ? OR LOWER(email) LIKE ? OR LOWER(company) LIKE ?
            ORDER BY first_name, last_name
        ''', (f'%{search_term}%', f'%{search_term}%', f'%{search_term}%', 
              f'%{search_term}%', f'%{search_term}%'))
        
        contacts = self.cursor.fetchall()
        
        for contact in contacts:
            full_name = f"{contact[1]} {contact[2]}"
            self.contacts_tree.insert("", "end", values=(
                contact[0], full_name, contact[3] or "-", contact[4] or "-", 
                contact[5] or "-", contact[6] or "-"
            ))
        
        # Update status
        self.status_label.configure(text=f"Found {len(contacts)} contacts matching '{search_term}'")
    
    def save_contact(self):
        """Save contact to database"""
        # Get form data
        data = {}
        for field, widget in self.entry_widgets.items():
            data[field] = widget.get().strip()
        
        address = self.address_text.get("1.0", "end-1c").strip()
        notes = self.notes_text.get("1.0", "end-1c").strip()
        
        # Validate required fields
        if not data['first_name'] or not data['last_name']:
            messagebox.showerror("Error", "First Name and Last Name are required!")
            return
        
        try:
            if self.editing_id:
                # Update existing contact
                self.cursor.execute('''
                    UPDATE contacts 
                    SET first_name=?, last_name=?, phone=?, email=?, address=?, 
                        company=?, notes=?, category=?, last_modified=?
                    WHERE id=?
                ''', (data['first_name'], data['last_name'], data['phone'], 
                      data['email'], address, data['company'], notes, 
                      data['category'], datetime.now(), self.editing_id))
                messagebox.showinfo("Success", "✅ Contact updated successfully!")
            else:
                # Insert new contact
                self.cursor.execute('''
                    INSERT INTO contacts 
                    (first_name, last_name, phone, email, address, company, notes, category)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (data['first_name'], data['last_name'], data['phone'], 
                      data['email'], address, data['company'], notes, data['category']))
                messagebox.showinfo("Success", "✅ Contact added successfully!")
            
            self.conn.commit()
            self.clear_form()
            self.show_contacts()
            
        except sqlite3.Error as e:
            messagebox.showerror("Database Error", f"❌ Failed to save contact: {str(e)}")
    
    def clear_form(self):
        """Clear the contact form"""
        for widget in self.entry_widgets.values():
            widget.delete(0, "end")
        self.address_text.delete("1.0", "end")
        self.notes_text.delete("1.0", "end")
        self.editing_id = None
        self.contact_form_title.configure(text="➕ Add New Contact")
    
    def on_contact_double_click(self, event):
        """Handle double click on contact"""
        self.edit_contact()
    
    def edit_contact(self):
        """Edit selected contact"""
        selected_item = self.contacts_tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "⚠️ Please select a contact to edit!")
            return
        
        contact_id = self.contacts_tree.item(selected_item[0])['values'][0]
        
        # Fetch contact details
        self.cursor.execute('SELECT * FROM contacts WHERE id=?', (contact_id,))
        contact = self.cursor.fetchone()
        
        if contact:
            self.show_add_contact()
            self.contact_form_title.configure(text="✏️ Edit Contact")
            self.editing_id = contact_id
            
            # Fill form with contact data
            self.entry_widgets['first_name'].insert(0, contact[1])
            self.entry_widgets['last_name'].insert(0, contact[2])
            self.entry_widgets['phone'].insert(0, contact[3] or "")
            self.entry_widgets['email'].insert(0, contact[4] or "")
            self.entry_widgets['company'].insert(0, contact[6] or "")
            self.entry_widgets['category'].insert(0, contact[8] or "")
            
            if contact[5]:
                self.address_text.insert("1.0", contact[5])
            if contact[7]:
                self.notes_text.insert("1.0", contact[7])
    
    def delete_contact(self):
        """Delete selected contact"""
        selected_item = self.contacts_tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "⚠️ Please select a contact to delete!")
            return
        
        contact_name = self.contacts_tree.item(selected_item[0])['values'][1]
        contact_id = self.contacts_tree.item(selected_item[0])['values'][0]
        
        if messagebox.askyesno("Confirm Delete", 
                             f"Are you sure you want to delete '{contact_name}'?"):
            try:
                self.cursor.execute('DELETE FROM contacts WHERE id=?', (contact_id,))
                self.conn.commit()
                messagebox.showinfo("Success", "✅ Contact deleted successfully!")
                self.load_contacts()
            except sqlite3.Error as e:
                messagebox.showerror("Database Error", f"❌ Failed to delete contact: {str(e)}")
    
    def export_contacts(self):
        """Export all contacts to CSV"""
        try:
            self.cursor.execute('SELECT * FROM contacts')
            contacts = self.cursor.fetchall()
            
            if not contacts:
                messagebox.showinfo("Info", "ℹ️ No contacts to export!")
                return
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export contacts to CSV"
            )
            
            if filename:
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    # Write header
                    writer.writerow(['ID', 'First Name', 'Last Name', 'Phone', 'Email', 
                                   'Address', 'Company', 'Notes', 'Category', 'Created Date'])
                    
                    # Write data
                    for contact in contacts:
                        writer.writerow(contact)
                
                messagebox.showinfo("Success", f"✅ Contacts exported successfully to:\n{filename}")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"❌ Failed to export contacts: {str(e)}")
    
    def export_selected_contacts(self):
        """Export selected contacts to CSV"""
        selected_items = self.contacts_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "⚠️ Please select contacts to export!")
            return
        
        try:
            contact_ids = [self.contacts_tree.item(item)['values'][0] for item in selected_items]
            placeholders = ','.join('?' for _ in contact_ids)
            
            self.cursor.execute(f'SELECT * FROM contacts WHERE id IN ({placeholders})', contact_ids)
            contacts = self.cursor.fetchall()
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Export selected contacts to CSV"
            )
            
            if filename:
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    # Write header
                    writer.writerow(['ID', 'First Name', 'Last Name', 'Phone', 'Email', 
                                   'Address', 'Company', 'Notes', 'Category', 'Created Date'])
                    
                    # Write data
                    for contact in contacts:
                        writer.writerow(contact)
                
                messagebox.showinfo("Success", f"✅ {len(contacts)} contacts exported successfully to:\n{filename}")
                
        except Exception as e:
            messagebox.showerror("Export Error", f"❌ Failed to export contacts: {str(e)}")
    
    def export_detailed_contacts(self):
        """Export contacts with detailed information"""
        self.export_contacts()  # Same as regular export for now
    
    def import_contacts(self):
        """Import contacts from CSV"""
        try:
            filename = filedialog.askopenfilename(
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                title="Import contacts from CSV"
            )
            
            if not filename:
                return
            
            with open(filename, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                imported_count = 0
                
                for row in reader:
                    try:
                        # Insert contact
                        self.cursor.execute('''
                            INSERT INTO contacts 
                            (first_name, last_name, phone, email, address, company, notes, category)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            row.get('first_name', '').strip(),
                            row.get('last_name', '').strip(),
                            row.get('phone', '').strip(),
                            row.get('email', '').strip(),
                            row.get('address', '').strip(),
                            row.get('company', '').strip(),
                            row.get('notes', '').strip(),
                            row.get('category', '').strip()
                        ))
                        imported_count += 1
                    except Exception as e:
                        print(f"Error importing row {row}: {e}")
                        continue
                
                self.conn.commit()
                messagebox.showinfo("Success", f"✅ Successfully imported {imported_count} contacts!")
                self.load_contacts()
                
        except Exception as e:
            messagebox.showerror("Import Error", f"❌ Failed to import contacts: {str(e)}")
    
    def backup_database(self):
        """Create a backup of the database"""
        try:
            backup_filename = f"contacts_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            
            # Simple file copy for backup
            import shutil
            shutil.copy2('contacts.db', backup_filename)
            
            messagebox.showinfo("Success", f"✅ Database backed up successfully as:\n{backup_filename}")
            
        except Exception as e:
            messagebox.showerror("Backup Error", f"❌ Failed to backup database: {str(e)}")
    
    def reset_database(self):
        """Reset the database (delete all contacts)"""
        if messagebox.askyesno("Confirm Reset", 
                             "⚠️ This will delete ALL contacts! Are you sure?"):
            try:
                self.cursor.execute('DELETE FROM contacts')
                self.conn.commit()
                messagebox.showinfo("Success", "✅ Database reset successfully!")
                self.load_contacts()
            except Exception as e:
                messagebox.showerror("Reset Error", f"❌ Failed to reset database: {str(e)}")
    
    def get_total_contacts(self):
        """Get total number of contacts"""
        self.cursor.execute('SELECT COUNT(*) FROM contacts')
        return self.cursor.fetchone()[0]
    
    def get_recent_contacts(self):
        """Get number of contacts added in last 7 days"""
        self.cursor.execute('''
            SELECT COUNT(*) FROM contacts 
            WHERE created_date >= datetime('now', '-7 days')
        ''')
        return self.cursor.fetchone()[0]
    
    def get_categories_count(self):
        """Get number of unique categories"""
        self.cursor.execute('SELECT COUNT(DISTINCT category) FROM contacts WHERE category IS NOT NULL AND category != ""')
        return self.cursor.fetchone()[0]
    
    def change_appearance_mode(self, new_appearance_mode):
        """Change appearance mode"""
        ctk.set_appearance_mode(new_appearance_mode)
    
    def run(self):
        """Run the application"""
        self.root.mainloop()
    
    def __del__(self):
        """Close database connection"""
        if hasattr(self, 'conn'):
            self.conn.close()

if __name__ == "__main__":
    app = ContactManagementSystem()
    app.run()