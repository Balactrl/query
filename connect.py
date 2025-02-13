import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import pyodbc
from tkcalendar import DateEntry
import os
import pandas as pd

def connect_to_database(site_id, username, password, database, custom_ip=None):
    """
    Connect to the SQL Server using the provided credentials.
    If a custom IP is provided, it is used directly.
    Otherwise, the site_id is formatted (e.g. '13100' becomes '131.00')
    and appended to a default IP series ("10.16.").
    """
    if custom_ip:
        host = custom_ip
    else:
        try:
            formatted_site_id = f"{site_id[:3]}.{int(site_id[3:])}"
        except Exception as e:
            raise ValueError(f"Error formatting Site ID: {e}")
        host = f"10.16.{formatted_site_id}"
    try:
        connection = pyodbc.connect(
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={host};"
            f"UID={username};"
            f"PWD={password};"
            f"DATABASE={database};"
        )
        return connection
    except pyodbc.Error as e:
        raise Exception(f"Connection error for site {site_id}: {e}")

class SQLQueryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SQL Multi-Site Query Runner")
        self.geometry("1200x750")
        self.file_path = None   # For file-based site IDs
        # Instead of results_data, we'll store results grouped by query number.
        self.results_by_query = {}  
        self.create_widgets()
    
    def create_widgets(self):
        # Main frame with two columns: left (controls) and right (output)
        main_frame = ttk.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0,10))
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew")
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # --- Credentials Frame in left_frame (two rows) ---
        creds_frame = ttk.LabelFrame(left_frame, text="DB Credentials", padding=10)
        creds_frame.pack(fill="x", padx=5, pady=5)
        
        # Row 0: Username, Password, Database
        ttk.Label(creds_frame, text="Username:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.username_entry = ttk.Entry(creds_frame, width=15)
        self.username_entry.grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(creds_frame, text="Password:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.password_entry = ttk.Entry(creds_frame, width=15, show="*")
        self.password_entry.grid(row=0, column=3, padx=5, pady=5)
        ttk.Label(creds_frame, text="Database:").grid(row=0, column=4, padx=5, pady=5, sticky='w')
        self.database_entry = ttk.Entry(creds_frame, width=15)
        self.database_entry.grid(row=0, column=5, padx=5, pady=5)
        
        # Row 1: Site ID, Custom IP, Date, Test Connection
        ttk.Label(creds_frame, text="Site ID:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.siteid_entry = ttk.Entry(creds_frame, width=15)
        self.siteid_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(creds_frame, text="Custom IP (optional):").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        self.custom_ip_entry = ttk.Entry(creds_frame, width=15)
        self.custom_ip_entry.grid(row=1, column=3, padx=5, pady=5)
        ttk.Label(creds_frame, text="Date (optional):").grid(row=1, column=4, padx=5, pady=5, sticky='w')
        self.date_entry = DateEntry(creds_frame, date_pattern='yyyy-mm-dd', width=12)
        self.date_entry.grid(row=1, column=5, padx=5, pady=5)
        self.connect_button = ttk.Button(creds_frame, text="Test Connection", command=self.test_connection)
        self.connect_button.grid(row=1, column=6, padx=5, pady=5)
        
        # --- File Upload Frame in left_frame ---
        file_frame = ttk.LabelFrame(left_frame, text="Site IDs File (Optional)", padding=10)
        file_frame.pack(fill="x", padx=5, pady=5)
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side="left", padx=5, pady=5)
        ttk.Button(file_frame, text="Select File", command=self.select_file).pack(side="left", padx=5, pady=5)
        ttk.Button(file_frame, text="Clear File", command=self.clear_file).pack(side="left", padx=5, pady=5)
        
        # --- Query Input Frame in left_frame ---
        query_frame = ttk.LabelFrame(left_frame, text="SQL Query", padding=10)
        query_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.query_text = ScrolledText(query_frame, wrap="word", height=10)
        self.query_text.pack(fill="both", expand=True)
        query_buttons_frame = ttk.Frame(query_frame)
        query_buttons_frame.pack(pady=5)
        self.run_button = ttk.Button(query_buttons_frame, text="Run Query", command=self.run_query)
        self.run_button.pack(side="left", padx=5)
        self.clear_query_button = ttk.Button(query_buttons_frame, text="Clear Query", command=self.clear_query)
        self.clear_query_button.pack(side="left", padx=5)
        
        # --- Output Frame in right_frame ---
        output_frame = ttk.LabelFrame(right_frame, text="Output", padding=10)
        output_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.output_text = ScrolledText(output_frame, wrap="word", height=20)
        self.output_text.pack(fill="both", expand=True)
        buttons_frame = ttk.Frame(right_frame)
        buttons_frame.pack(pady=5)
        self.clear_output_button = ttk.Button(buttons_frame, text="Clear Output", command=self.clear_output)
        self.clear_output_button.pack(side="left", padx=5)
        self.download_button = ttk.Button(buttons_frame, text="Download Report", command=self.download_to_excel, state="disabled")
        self.download_button.pack(side="left", padx=5)
    
    def select_file(self):
        filename = filedialog.askopenfilename(
            title="Select File",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("Text files", "*.txt")]
        )
        if filename:
            self.file_path = filename
            self.file_label.config(text=os.path.basename(filename))
    
    def clear_file(self):
        self.file_path = None
        self.file_label.config(text="No file selected")
    
    def test_connection(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        database = self.database_entry.get().strip()
        site_id = self.siteid_entry.get().strip()
        custom_ip = self.custom_ip_entry.get().strip()
        
        if not (username and password and database and site_id):
            messagebox.showerror("Error", "Please fill in Username, Password, Database, and Site ID.")
            return
        
        try:
            conn = connect_to_database(site_id, username, password, database, custom_ip if custom_ip else None)
            conn.close()
            messagebox.showinfo("Success", f"Connection successful for Site ID: {site_id}")
        except Exception as e:
            messagebox.showerror("Connection Error", str(e))
    
    def run_query(self):
        # Split queries by semicolon.
        query_text = self.query_text.get("1.0", tk.END).strip()
        if not query_text:
            messagebox.showerror("Error", "Please enter a SQL query.")
            return
        queries = [q.strip() for q in query_text.split(";") if q.strip()]
        
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        database = self.database_entry.get().strip()
        custom_ip = self.custom_ip_entry.get().strip()
        optional_date = self.date_entry.get_date()
        
        if not (username and password and database):
            messagebox.showerror("Error", "Please fill in Username, Password, and Database.")
            return
        
        self.output_text.delete("1.0", tk.END)
        # We'll group results by query number.
        self.results_by_query = {}
        
        # Determine site IDs.
        if self.file_path:
            try:
                ext = os.path.splitext(self.file_path)[1].lower()
                if ext == '.xlsx':
                    df = pd.read_excel(self.file_path)
                    cols = [col.lower() for col in df.columns]
                    if "siteid" in cols:
                        site_ids = df["siteid"].astype(str).tolist()
                    else:
                        raise Exception("Excel file must contain a column named 'siteid'.")
                elif ext == '.csv':
                    df = pd.read_csv(self.file_path)
                    cols = [col.lower() for col in df.columns]
                    if "siteid" in cols:
                        site_ids = df["siteid"].astype(str).tolist()
                    else:
                        raise Exception("CSV file must contain a column named 'siteid'.")
                else:
                    with open(self.file_path, "r", encoding="utf-8") as f:
                        content = f.read()
                    if "," in content:
                        site_ids = [s.strip() for s in content.split(",") if s.strip()]
                    else:
                        site_ids = [s.strip() for s in content.splitlines() if s.strip()]
            except Exception as e:
                messagebox.showerror("File Error", f"Error reading file: {e}")
                return
        else:
            site_id = self.siteid_entry.get().strip()
            if not site_id:
                messagebox.showerror("Error", "Please enter a Site ID or select a file with Site IDs.")
                return
            site_ids = [site_id]
        
        # Process each site and each query.
        for sid in site_ids:
            self.output_text.insert(tk.END, f"--- Processing Site ID: {sid} ---\n")
            try:
                conn = connect_to_database(sid, username, password, database, custom_ip if custom_ip else None)
                query_count = 1
                for q in queries:
                    self.output_text.insert(tk.END, f"--- Running Query {query_count} ---\n")
                    cur = conn.cursor()
                    cur.execute(q)
                    if q.strip().lower().startswith("select"):
                        rows = cur.fetchall()
                        columns = [col[0] for col in cur.description] if cur.description else []
                        
                        rows = [tuple(row) for row in rows]
                        if rows and len(rows[0]) == 1 and isinstance(rows[0][0], str):
                            new_rows = []
                            for r in rows:
                                splitted = str(r[0]).split(',')
                                if len(splitted) == len(columns):
                                    new_rows.append(tuple(splitted))
                                else:
                                    new_rows.append(r)
                            rows = new_rows
                        
                        df = pd.DataFrame(rows, columns=columns) if columns else pd.DataFrame(rows)
                        # Add a column for the Site ID
                        df.insert(0, "SiteID", sid)
                        # Store results grouped by query number.
                        key = f"Q{query_count}"
                        if key in self.results_by_query:
                            self.results_by_query[key].append(df)
                        else:
                            self.results_by_query[key] = [df]
                        
                        if columns:
                            self.output_text.insert(tk.END, "\t".join(columns) + "\n")
                            self.output_text.insert(tk.END, "-" * 50 + "\n")
                        if rows:
                            for row in rows:
                                self.output_text.insert(tk.END, "\t".join([str(item) for item in row]) + "\n")
                        else:
                            self.output_text.insert(tk.END, "No rows returned.\n")
                    else:
                        conn.commit()
                        affected = cur.rowcount
                        self.output_text.insert(tk.END, f"Query executed successfully. Rows affected: {affected}\n")
                    self.output_text.insert(tk.END, "\n")
                    cur.close()
                    query_count += 1
                self.output_text.insert(tk.END, f"Optional Date: {optional_date.strftime('%Y-%m-%d')}\n\n")
                conn.close()
            except Exception as e:
                self.output_text.insert(tk.END, f"Error processing Site ID {sid}: {e}\n\n")
        
        # Enable Download Report if any SELECT queries were run.
        if any(q.strip().lower().startswith("select") for q in queries):
            self.download_button.config(state="normal")
        else:
            self.download_button.config(state="disabled")
    
    def download_to_excel(self):
        if not self.results_by_query:
            messagebox.showerror("Error", "No results to export.")
            return
        
        downloads_folder = "C:\\Users\\APL41051\\Downloads"
        if not os.path.exists(downloads_folder):
            os.makedirs(downloads_folder, exist_ok=True)
        
        base_filename = "query"
        extension = ".xlsx"
        file_path = os.path.join(downloads_folder, base_filename + extension)
        counter = 1
        while os.path.exists(file_path):
            file_path = os.path.join(downloads_folder, f"{base_filename}{counter}{extension}")
            counter += 1
        
        # Combine results for each query number and write each to its own sheet.
        try:
            with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
                for key, df_list in self.results_by_query.items():
                    # Concatenate all DataFrames for this query.
                    combined_df = pd.concat([d.reset_index(drop=True) for d in df_list], ignore_index=True)
                    sheet_name = key if len(key) <= 31 else key[:31]
                    combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
            messagebox.showinfo("Success", f"Results downloaded successfully to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Download Error", f"Error downloading to Excel: {e}")
    
    def clear_output(self):
        self.output_text.delete("1.0", tk.END)
        self.download_button.config(state="disabled")
    
    def clear_query(self):
        self.query_text.delete("1.0", tk.END)

if __name__ == "__main__":
    app = SQLQueryApp()
    app.mainloop()
