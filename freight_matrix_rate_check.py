import os
import json
import time
import threading
import requests
import pandas as pd
from datetime import datetime
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
MAX_REQUESTS_PER_MINUTE = 150
REQUEST_INTERVAL = 60 / MAX_REQUESTS_PER_MINUTE
API_URL = "https://..."
API_KEY = "api_key"
RETRY_ATTEMPTS = 3
BACKOFF_FACTOR = 2
def get_output_path(prefix="transformed_data"):
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"{prefix}_{current_time}.xlsx"
    output_dir = os.path.expanduser('~/Downloads')
    return os.path.join(output_dir, output_filename)
def read_input_file(file_path):
    if file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file type. Please select an Excel or CSV file.")
def transform_data(df):
    if not {'SKU', 'Postcode'}.issubset(df.columns):
        raise ValueError("Input file must contain 'SKU' and 'Postcode' columns.")
    return [{"productCode": row['SKU'], "postCode": row['Postcode']} for _, row in df.iterrows()]
def send_to_api(json_data):
    headers = {"api-key": API_KEY, "Content-Type": "application/json"}
    attempt = 0
    while attempt < RETRY_ATTEMPTS:
        try:
            response = requests.post(API_URL, headers=headers, data=json.dumps(json_data))
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError:
            status_code = response.status_code
            error_map = {
                400: "Bad Request: Check your input format.",
                403: "Forbidden: Invalid API key or insufficient permissions.",
                404: "Not Found: API endpoint or resource missing.",
                409: "Conflict: Duplicate or invalid data detected.",
                500: "Internal Server Error: Please try again later."
            }
            message = error_map.get(status_code, f"Unexpected Error: {status_code}")
            if attempt < RETRY_ATTEMPTS - 1:
                time.sleep(BACKOFF_FACTOR ** attempt)
                attempt += 1
            else:
                raise Exception(f"API Error {status_code}: {message}")
        except requests.exceptions.RequestException as req_err:
            if attempt < RETRY_ATTEMPTS - 1:
                time.sleep(BACKOFF_FACTOR ** attempt)
                attempt += 1
            else:
                raise Exception(f"Network Error: {str(req_err)}")
def save_results_to_excel(data, output_path):
    pd.DataFrame(data).to_excel(output_path, index=False)
def create_template():
    sample_data = pd.DataFrame({"SKU": ["ABC123", "XYZ789"], "Postcode": ["2000", "3000"]})
    output_path = get_output_path(prefix="template")
    sample_data.to_excel(output_path, index=False)
    return output_path
class IndividualCheckTab:
    def __init__(self, parent, app):
        self.app = app
        self.frame = tb.Frame(parent)
        self.build_ui()
    def build_ui(self):
        tb.Label(self.frame, text="Individual Rate Check", font=("Helvetica", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
        tb.Label(self.frame, text="SKU:").grid(row=1, column=0, sticky=E, padx=10)
        self.sku_var = tb.StringVar()
        tb.Entry(self.frame, textvariable=self.sku_var).grid(row=1, column=1, sticky=W, padx=10)
        tb.Label(self.frame, text="Postcode:").grid(row=2, column=0, sticky=E, padx=10)
        self.postcode_var = tb.StringVar()
        tb.Entry(self.frame, textvariable=self.postcode_var).grid(row=2, column=1, sticky=W, padx=10)
        tb.Button(self.frame, text="Check Rate", bootstyle=INFO, command=self.check_rate).grid(row=3, column=0, columnspan=2, pady=15)
        self.result_label = tb.Label(self.frame, text="", wraplength=350, justify=LEFT)
        self.result_label.grid(row=4, column=0, columnspan=2, pady=10)
    def check_rate(self):
        sku = self.sku_var.get().strip()
        postcode = self.postcode_var.get().strip()
        if not sku or not postcode:
            messagebox.showwarning("Input required", "Please enter both SKU and Postcode.")
            return
        try:
            json_data = [{"productCode": sku, "postCode": postcode}]
            response_data = send_to_api(json_data)[0]
            if response_data.get('deliveryPossible'):
                result_text = (
                    f"Delivery Possible: Yes\n"
                    f"Delivery Rate: ${response_data.get('deliveryRate')}\n"
                    f"Product Code: {response_data.get('productCode')}\n"
                    f"Postcode: {response_data.get('postCode')}"
                )
                self.result_label.configure(text=result_text, foreground="green")
            else:
                result_text = (
                    f"Delivery Not Possible\n"
                    f"Product Code: {response_data.get('productCode')}\n"
                    f"Postcode: {response_data.get('postCode')}"
                )
                self.result_label.configure(text=result_text, foreground="red")
            self.app.log(f"Checked rate for SKU {sku}, Postcode {postcode}", "INFO")
        except Exception as e:
            self.result_label.configure(text=f"Error: {str(e)}", foreground="red")
            self.app.log(str(e), "ERROR")
class BulkProcessingTab:
    def __init__(self, parent, app):
        self.app = app
        self.frame = tb.Frame(parent)
        self.file_path = None
        self.build_ui()
    def build_ui(self):
        tb.Label(self.frame, text="Bulk Processing", font=("Helvetica", 14, "bold")).grid(row=0, column=0, pady=10, sticky=W)
        self.select_btn = tb.Button(self.frame, text="Select File", bootstyle=PRIMARY, command=self.select_file)
        self.select_btn.grid(row=1, column=0, pady=5, sticky=W)
        self.template_btn = tb.Button(self.frame, text="Download Template", bootstyle=SECONDARY, command=self.download_template)
        self.template_btn.grid(row=2, column=0, pady=5, sticky=W)
        self.process_btn = tb.Button(self.frame, text="Process", bootstyle=SUCCESS, command=self.start_processing_thread)
        self.process_btn.grid(row=3, column=0, pady=5, sticky=W)
        self.export_btn = tb.Button(self.frame, text="Export Results", bootstyle=INFO, command=self.export_results)
        self.export_btn.grid(row=4, column=0, pady=5, sticky=W)
        self.file_label = tb.Label(self.frame, text="No file selected")
        self.file_label.grid(row=5, column=0, pady=10, sticky=W)
        self.progress = tb.Progressbar(self.frame, mode='determinate', length=300)
        self.progress.grid(row=6, column=0, pady=10)
    def set_buttons_state(self, state):
        for btn in [self.select_btn, self.template_btn, self.process_btn, self.export_btn]:
            btn.configure(state=state)
    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        self.file_label.configure(text=f"Selected: {os.path.basename(self.file_path)}" if self.file_path else "No file selected")
    def download_template(self):
        try:
            path = create_template()
            messagebox.showinfo("Template Downloaded", f"Template saved to:\n{path}")
            self.app.log(f"Template downloaded: {path}", "SUCCESS")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.app.log(str(e), "ERROR")
    def start_processing_thread(self):
        if not self.file_path:
            messagebox.showwarning("No file selected", "Please select a file.")
            return
        thread = threading.Thread(target=self.process_file)
        thread.start()
    def process_file(self):
        try:
            self.app.root.after(0, lambda: self.set_buttons_state(DISABLED))
            self.app.root.after(0, lambda: self.app.set_state("Processing...", style=PRIMARY))
            df = read_input_file(self.file_path)
            json_data = transform_data(df)
            response_data = []
            total_batches = (len(json_data) // MAX_REQUESTS_PER_MINUTE) + 1
            self.app.root.after(0, lambda: self.progress.configure(maximum=total_batches))
            for i in range(0, len(json_data), MAX_REQUESTS_PER_MINUTE):
                batch = json_data[i:i+MAX_REQUESTS_PER_MINUTE]
                try:
                    response_data.extend(send_to_api(batch))
                except Exception as e:
                    self.app.root.after(0, lambda: messagebox.showerror("Error", str(e)))
                    self.app.log(str(e), "ERROR")
                    break
                self.app.root.after(0, lambda: self.progress.step(1))
                time.sleep(REQUEST_INTERVAL)
            if response_data:
                output_path = get_output_path()
                save_results_to_excel(response_data, output_path)
                self.app.results_data = response_data
                self.app.root.after(0, lambda: self.app.update_results(response_data))
                self.app.root.after(0, lambda: messagebox.showinfo("Success", f"Processed data saved to:\n{output_path}"))
                self.app.log(f"Processed file: {output_path}", "SUCCESS")
                self.app.root.after(0, lambda: self.app.set_state("Completed", style=SUCCESS))
        finally:
            self.app.root.after(0, lambda: self.progress.configure(value=0))
            self.app.root.after(0, lambda: self.set_buttons_state(NORMAL))
    def export_results(self):
        try:
            if not self.app.results_data:
                messagebox.showwarning("No Results", "No results to export.")
                return
            path = get_output_path(prefix="exported_results")
            save_results_to_excel(self.app.results_data, path)
            messagebox.showinfo("Exported", f"Results exported to:\n{path}")
            self.app.log(f"Results exported: {path}", "SUCCESS")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.app.log(str(e), "ERROR")
class FreightRateApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Freight Matrix Rate App")
        self.root.geometry("400x600")
        self.results_data = []
        self.menu_bar = tb.Menu(root)
        root.config(menu=self.menu_bar)
        file_menu = tb.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Export Logs", command=self.export_logs)
        file_menu.add_command(label="Clear Results", command=self.clear_results)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=root.quit)
        self.menu_bar.add_cascade(label="File", menu=file_menu)
        theme_menu = tb.Menu(self.menu_bar, tearoff=0)
        for theme in tb.Style().theme_names():
            theme_menu.add_command(label=theme, command=lambda t=theme: self.change_theme(t))
        self.menu_bar.add_cascade(label="Theme", menu=theme_menu)
        self.notebook = tb.Notebook(root)
        self.notebook.pack(fill=BOTH, expand=True)
        self.individual_tab = IndividualCheckTab(self.notebook, self)
        self.bulk_tab = BulkProcessingTab(self.notebook, self)
        self.logs_tab = tb.Frame(self.notebook)
        self.results_tab = tb.Frame(self.notebook)
        self.notebook.add(self.individual_tab.frame, text="Individual Check")
        self.notebook.add(self.bulk_tab.frame, text="Bulk Processing")
        self.notebook.add(self.logs_tab, text="Logs")
        self.notebook.add(self.results_tab, text="Results")
        self.log_text = ScrolledText(self.logs_tab, height=20)
        self.log_text.pack(fill=BOTH, expand=True)
        results_frame = tb.Frame(self.results_tab)
        results_frame.pack(fill=BOTH, expand=True)
        self.tree = tb.Treeview(results_frame, columns=("ProductCode", "PostCode", "Rate"), show="headings")
        for col in ("ProductCode", "PostCode", "Rate"):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        vsb = tb.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        vsb.pack(side=RIGHT, fill=Y)
        self.status_label = tb.Label(root, text="Idle", bootstyle=INFO, anchor="w")
        self.status_label.pack(fill=X, side=BOTTOM)
    def log(self, message, level="INFO"):
        color = {"INFO": "white", "SUCCESS": "green", "ERROR": "red"}.get(level, "white")
        self.log_text.insert("end", f"[{level}] {message}\n", level)
        self.log_text.tag_config(level, foreground=color)
        self.log_text.see("end")
    def update_results(self, data):
        for row in data:
            self.tree.insert("", "end", values=(row.get("productCode"), row.get("postCode"), row.get("deliveryRate")))
    def set_state(self, state, style=INFO):
        self.status_label.configure(text=state, bootstyle=style)
    def export_logs(self):
        path = get_output_path(prefix="logs")
        with open(path.replace(".xlsx", ".txt"), "w") as f:
            f.write(self.log_text.get("1.0", "end"))
        messagebox.showinfo("Logs Exported", f"Logs saved to:\n{path.replace('.xlsx', '.txt')}")
        self.log(f"Logs exported to {path}", "SUCCESS")
    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results_data = []
        self.log("Results cleared", "INFO")
    def change_theme(self, theme_name):
        tb.Style().theme_use(theme_name)
        self.log(f"Theme changed to {theme_name}", "INFO")
if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = FreightRateApp(root)
    root.mainloop()
