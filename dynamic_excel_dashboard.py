import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import datetime
import win32com.client as win32
import os

class DynamicExcelDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Dynamic Universal Excel Dashboard & Mailer")
        self.root.geometry("1100x700")
        
        # Style
        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use('clam')
            
        self.bg_color = "#f4f6f9"
        self.header_bg = "#2c3e50"
        self.root.configure(bg=self.bg_color)
        
        self.df = None
        self.excel_path = "contracts.xlsx"
        
        self.create_widgets()
        
        # Auto-load on startup
        if os.path.exists(self.excel_path):
            self.upload_file()
            
    def create_widgets(self):
        # Top Header
        header_frame = tk.Frame(self.root, bg=self.header_bg, height=70)
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text="📊 Dynamic Excel Viewer & Mailer", bg=self.header_bg, fg="white", font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT, padx=20, pady=15)
        
        # Action Bar (Load File)
        action_bar = tk.Frame(self.root, bg=self.bg_color)
        action_bar.pack(fill=tk.X, padx=20, pady=10)
        
        btn_load = tk.Button(action_bar, text="🔄 Reload File Data", command=self.upload_file, bg="#007bff", fg="white", font=("Segoe UI", 11, "bold"), padx=15, pady=5, relief=tk.FLAT)
        btn_load.pack(side=tk.LEFT)
        
        self.lbl_file_info = tk.Label(action_bar, text="No file loaded.", bg=self.bg_color, font=("Segoe UI", 10, "italic"), fg="#6c757d")
        self.lbl_file_info.pack(side=tk.LEFT, padx=15)
        
        # Stats summary
        self.stats_frame = tk.Frame(self.root, bg=self.bg_color)
        self.stats_frame.pack(fill=tk.X, padx=20, pady=5)
        
        self.lbl_rows = tk.Label(self.stats_frame, text="Rows: 0", font=("Segoe UI", 12, "bold"), bg=self.bg_color, fg="#17a2b8")
        self.lbl_rows.pack(side=tk.LEFT, padx=(0, 20))
        self.lbl_cols = tk.Label(self.stats_frame, text="Columns: 0", font=("Segoe UI", 12, "bold"), bg=self.bg_color, fg="#f39c12")
        self.lbl_cols.pack(side=tk.LEFT, padx=20)
        
        # Dynamic Table View Frame
        self.table_frame = tk.Frame(self.root, bg=self.bg_color)
        self.table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # We will mount the treeview inside this frame dynamically
        self.tree_container = tk.Frame(self.table_frame)
        self.tree_container.pack(fill=tk.BOTH, expand=True)
        
        self.tree = ttk.Treeview(self.tree_container)
        
        # Bottom controls for Mailer
        bottom_frame = tk.Frame(self.root, bg=self.bg_color, highlightbackground="#ddd", highlightthickness=1)
        bottom_frame.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(bottom_frame, text="Mailer Controls:", bg=self.bg_color, font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=10, pady=10)
        
        tk.Label(bottom_frame, text="Days Notice:", bg=self.bg_color).pack(side=tk.LEFT, padx=5)
        self.days_var = tk.IntVar(value=180)
        tk.Entry(bottom_frame, textvariable=self.days_var, width=5).pack(side=tk.LEFT)
        
        tk.Label(bottom_frame, text="Specific Sender Email (Optional):", bg=self.bg_color).pack(side=tk.LEFT, padx=(15, 5))
        self.sender_var = tk.StringVar(value="emdnkstppcontracts@outlook.com")
        tk.Entry(bottom_frame, textvariable=self.sender_var, width=35).pack(side=tk.LEFT)
        
        self.force_var = tk.BooleanVar(value=False)
        tk.Checkbutton(bottom_frame, text="Force Send (Ignore 7-Day Rule)", variable=self.force_var, bg=self.bg_color).pack(side=tk.LEFT, padx=10)
        
        btn_send = tk.Button(bottom_frame, text="📧 Attempt to Send Expiry Emails", command=self.send_emails, bg="#28a745", fg="white", font=("Segoe UI", 10, "bold"), padx=10, pady=5, relief=tk.FLAT)
        btn_send.pack(side=tk.RIGHT, padx=10)

    def upload_file(self):
        file_path = "contracts.xlsx"
        if not os.path.exists(file_path):
            messagebox.showwarning("File Missing", f"Could not find '{file_path}'. Please place it in the same folder as the script.")
            return
            
        self.excel_path = file_path
        self.lbl_file_info.config(text=f"Loading {os.path.basename(file_path)}...")
        self.root.update()
        
        try:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)
                
            # Cleanup column names
            self.df.columns = self.df.columns.astype(str).str.strip()
            
            # --- Calculate Days to Expiry ---
            expiry_col = None
            for col in self.df.columns:
                if col.lower() in ['expiry date', 'expiry_date', 'expiry']:
                    expiry_col = col
                    break
                    
            if expiry_col:
                today = datetime.datetime.now()
                try:
                    # Convert to datetime objects, coercing errors to NaT
                    temp_dates = pd.to_datetime(self.df[expiry_col], dayfirst=True, errors='coerce')
                    # Calculate difference in days and handle missing/invalid dates gracefully
                    self.df['Days to Expiry'] = temp_dates.apply(
                        lambda x: "" if pd.isna(x) else (x - today).days
                    )
                except Exception as e:
                    print(f"Could not calculate 'Days to Expiry': {e}")
            # --------------------------------
            
            # Rebuild treeview completely to accommodate new dynamic columns
            self.rebuild_treeview()
            
            self.lbl_file_info.config(text=f"Successfully loaded: {file_path}")
            self.lbl_rows.config(text=f"Rows: {len(self.df)}")
            self.lbl_cols.config(text=f"Columns: {len(self.df.columns)}")
            
        except Exception as e:
            messagebox.showerror("File Error", f"Could not read the file entirely.\nError:\n{e}")
            self.lbl_file_info.config(text="File load failed.")

    def rebuild_treeview(self):
        # Destroy current tree container contents completely
        for widget in self.tree_container.winfo_children():
            widget.destroy()
            
        # Create a new style for tree
        self.style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        
        columns = list(self.df.columns)
        
        self.tree = ttk.Treeview(self.tree_container, columns=columns, show="headings")
        
        # Configure scrollbars
        vsb = ttk.Scrollbar(self.tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Position everything
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Set headings
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor=tk.W)
            
        # Display data (limit to first 10,000 for gui stability if it's large)
        for _, row in self.df.head(10000).iterrows():
            # Convert NaN to empty string for display
            display_row = ["" if pd.isna(val) else str(val) for val in row]
            self.tree.insert("", tk.END, values=display_row)

    def send_emails(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Warning", "Please upload an Excel file first.")
            return
            
        # Soft-check for mapping columns required for emails
        required = ["Contract Name", "Expiry Date", "Email Address", "Person Name"]
        missing = [col for col in required if col not in self.df.columns]
        
        if missing:
            msg = f"Cannot send emails safely. Your data is missing the following column names required for the email template:\n\n{', '.join(missing)}\n\nPlease ensure your Excel headers match exactly, or rename them."
            messagebox.showerror("Missing Columns", msg)
            return
            
        # Convert dates
        try:
            self.df['Expiry Date'] = pd.to_datetime(self.df['Expiry Date'], errors='coerce')
        except Exception:
            pass
            
        if 'Last Emailed Date' not in self.df.columns:
            self.df['Last Emailed Date'] = pd.NaT
            
        self.df['Last Emailed Date'] = pd.to_datetime(self.df['Last Emailed Date'], errors='coerce')
        
        if messagebox.askyesno("Confirm", "Do you want to run the email process using Outlook?"):
            self._execute_email_sender()
            
    def _execute_email_sender(self):
        today = datetime.datetime.now()
        days_notice = self.days_var.get()
        sender_email = self.sender_var.get().strip()
        force_send = self.force_var.get()
        
        try:
            outlook = win32.Dispatch('outlook.application')
            
            # Find the specific Outlook account if provided
            account_to_use = None
            if sender_email:
                for acc in outlook.Session.Accounts:
                    if acc.SmtpAddress.lower() == sender_email.lower():
                        account_to_use = acc
                        break
                
                if not account_to_use:
                    messagebox.showerror(
                        "Sender Not Found", 
                        f"Could not find an Outlook account linked to:\n{sender_email}\n\nPlease ensure this email is added under your Outlook app accounts."
                    )
                    return
                    
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Outlook:\n{e}")
            return
            
        updated = False
        
        # Identify contracts to send
        emails_to_send = {}
        for index, row in self.df.iterrows():
            if pd.isna(row.get('Email Address')) or pd.isna(row.get('Expiry Date')):
                continue
                
            days_left = (row['Expiry Date'] - today).days
            
            if 0 <= days_left <= days_notice:
                last_emailed = row.get('Last Emailed Date', pd.NaT)
                needs_email = False
                
                if force_send:
                    needs_email = True
                elif pd.isna(last_emailed):
                    needs_email = True
                else:
                    try:
                        if (today - last_emailed).days >= 7:
                            needs_email = True
                    except Exception:
                        needs_email = True
                        
                if needs_email:
                    email = str(row['Email Address']).strip()
                    if email not in emails_to_send:
                        emails_to_send[email] = {
                            'person_name': str(row.get('Person Name', 'Colleague')),
                            'contracts': [],
                            'indices': []
                        }
                    emails_to_send[email]['contracts'].append({
                        'name': str(row.get('Contract Name', 'Unknown')),
                        'expiry': row['Expiry Date'],
                        'days_left': days_left
                    })
                    emails_to_send[email]['indices'].append(index)
                    
        sent = 0
        
        for email, data in emails_to_send.items():
            try:
                mail = outlook.CreateItem(0)
                if account_to_use:
                    try:
                        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account_to_use))
                    except Exception:
                        pass # Fallback below
                        
                # Alternative fallback for pywin32 Outlook send as
                if sender_email and not account_to_use:
                    mail.SentOnBehalfOfName = sender_email
                
                mail.To = email
                mail.Subject = f"Reminder: {len(data['contracts'])} Contract(s) Expiring Soon"
                
                rows_html = ""
                for c in data['contracts']:
                    rows_html += f"<tr><td style='border: 1px solid #ddd; padding: 8px;'>{c['name']}</td>"
                    rows_html += f"<td style='border: 1px solid #ddd; padding: 8px; color: #d35400;'>{c['expiry'].strftime('%Y-%m-%d')}</td>"
                    rows_html += f"<td style='border: 1px solid #ddd; padding: 8px;'>{c['days_left']}</td></tr>"
                    
                table_html = f"<table style='border-collapse: collapse; width: 100%;'><tr><th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f4f6f9;'>Contract Name</th><th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f4f6f9;'>Expiry Date</th><th style='border: 1px solid #ddd; padding: 8px; text-align: left; background-color: #f4f6f9;'>Days Left</th></tr>{rows_html}</table>"
                
                mail.HTMLBody = f'''
                <html>
                <body style="font-family: Arial, sans-serif;">
                    <p>Dear {data['person_name']},</p>
                    <p>This is an automated reminder that the following contract(s) assigned to you are set to expire soon:</p>
                    {table_html}
                    <br>
                    <p>Please take the necessary actions to renew or close these contracts.</p>
                    <br>
                    <p>Best regards,<br>Contract Management System Dashboard</p>
                </body>
                </html>
                '''
                mail.Send()
                
                # Update last emailed date
                for idx in data['indices']:
                    self.df.at[idx, 'Last Emailed Date'] = today
                
                sent += 1
                updated = True
            except Exception as e:
                print(e)
                
        if updated:
            try:
                self.df.to_excel(self.excel_path, index=False)
                self.rebuild_treeview() # refresh showing new dates
            except Exception as e:
                messagebox.showerror("Save Error", f"Could not save to {self.excel_path}. Close the file if it is open in Excel.\nError: {e}")
                
        messagebox.showinfo("Done", f"Scan complete. {sent} batched emails sent.")

if __name__ == "__main__":
    root = tk.Tk()
    app = DynamicExcelDashboard(root)
    root.mainloop()
