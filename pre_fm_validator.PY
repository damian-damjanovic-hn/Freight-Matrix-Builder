import os
import csv
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.constants import PRIMARY, INFO, SUCCESS
from ttkbootstrap import ttk
import subprocess
import platform

class CSVConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to Clean CSV & JSON Converter")
        self.root.geometry("600x700")
        self.root.resizable(True, True)

        self.style = Style(theme="darkly")
        self.data_rows = []
        self.create_widgets()

    def create_widgets(self):
        self.root.columnconfigure(1, weight=1)
        self.root.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(self.root, padding=10)
        sidebar.grid(row=0, column=0, sticky="nsw")

        content = ttk.Frame(self.root, padding=10)
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        ttk.Label(sidebar, text="Statistics:", anchor="w").pack(pady=(0, 2), anchor="w")
        self.stats_box = tk.Text(sidebar, height=6, width=30, wrap=tk.WORD, state="disabled")
        self.stats_box.pack(pady=5, anchor="w")

        self.file_label = ttk.Label(sidebar, text="No file selected", wraplength=200, anchor="w")
        self.file_label.pack(pady=5, anchor="w")

        ttk.Button(sidebar, text="Select CSV File", command=self.load_file, bootstyle=PRIMARY).pack(pady=5, anchor="w")

        ttk.Label(sidebar, text="Paste CSV content:", anchor="w").pack(pady=5, anchor="w")
        self.paste_box = tk.Text(sidebar, height=8, width=30, wrap=tk.NONE)
        self.paste_box.pack(pady=5, anchor="w")

        ttk.Button(sidebar, text="Preview", command=self.preview_data, bootstyle=INFO).pack(pady=10, anchor="w")

        self.export_csv_var = tk.BooleanVar(value=True)
        self.export_json_var = tk.BooleanVar(value=True)

        ttk.Checkbutton(sidebar, text="Export CSV", variable=self.export_csv_var).pack(pady=2, anchor="w")
        ttk.Checkbutton(sidebar, text="Export JSON", variable=self.export_json_var).pack(pady=2, anchor="w")

        ttk.Button(sidebar, text="Export Files", command=self.export_files, bootstyle=SUCCESS).pack(pady=10, anchor="w")

        self.status_label = ttk.Label(sidebar, text="", foreground="lightgreen", wraplength=200, anchor="w")
        self.status_label.pack(pady=5, anchor="w")

        ttk.Label(content, text="Preview:").grid(row=0, column=0, sticky="w", pady=(0, 5))

        self.convert_preview_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(content, text="Convert price to float", variable=self.convert_preview_var).grid(row=1, column=0, sticky="w", pady=(0, 5))

        preview_frame = ttk.Frame(content)
        preview_frame.grid(row=2, column=0, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.preview_box = tk.Text(preview_frame, wrap=tk.NONE)
        self.preview_box.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_box.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.preview_box.config(yscrollcommand=scrollbar.set)

    def normalize_headers(self, row):
        return {k.strip().lower().strip('"'): v.strip().strip('"') for k, v in row.items()}

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            try:
                with open(file_path, newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    self.data_rows = [self.normalize_headers(row) for row in reader]
                self.status_label.config(text="File loaded successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file:\n{e}")

    def preview_data(self):
        if not self.data_rows and self.paste_box.get("1.0", tk.END).strip():
            try:
                pasted_text = self.paste_box.get("1.0", tk.END).strip().splitlines()
                reader = csv.DictReader(pasted_text)
                self.data_rows = [self.normalize_headers(row) for row in reader]
                self.status_label.config(text="Pasted content loaded.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to parse pasted CSV:\n{e}")
                return

        if not self.data_rows:
            messagebox.showerror("Error", "No data to preview.")
            return

        convert_mode = self.convert_preview_var.get()

        self.preview_box.delete("1.0", tk.END)
        preview_text = "postCode,sku,price\n"

        total_rows = len(self.data_rows)
        missing_postcode = 0
        invalid_price = 0
        missing_sku = 0
        unique_skus_set = set()

        for row in self.data_rows[:100]:
            post_code = str(row.get("postcode", "")).strip().strip('"')
            sku = str(row.get("productcode", "")).strip().strip('"')
            price_raw = str(row.get("price", "")).strip().strip('"')

            if not post_code:
                missing_postcode += 1
            if not sku:
                missing_sku += 1
            if not price_raw.replace('.', '', 1).isdigit():
                invalid_price += 1
            if sku:
                unique_skus_set.add(sku)

            if convert_mode:
                try:
                    price = float(price_raw)
                except:
                    price = 0.0
            else:
                price = price_raw

            preview_text += f"{post_code},{sku},{price}\n"

        self.preview_box.insert(tk.END, preview_text)

        stats_text = (
            f"Total rows: {total_rows}\n"
            f"Missing postCode: {missing_postcode}\n"
            f"Invalid price values: {invalid_price}\n"
            f"Missing SKUs: {missing_sku}\n"
            f"Unique SKUs: {len(unique_skus_set)}\n"
        )

        self.stats_box.config(state="normal")
        self.stats_box.delete("1.0", tk.END)
        self.stats_box.insert(tk.END, stats_text)
        self.stats_box.config(state="disabled")

        if missing_postcode > 0 or invalid_price > 0 or missing_sku > 0:
            self.stats_box.config(foreground="red")
        else:
            self.stats_box.config(foreground="lightgreen")

    def export_files(self):
        if not self.data_rows:
            messagebox.showerror("Error", "No data to export.")
            return

        output_data = []
        csv_data = []
        error_data = []

        for row in self.data_rows:
            post_code_raw = str(row.get("postcode", "")).strip().strip('"')
            post_code_csv = post_code_raw.lstrip("0") if post_code_raw else ""
            sku = str(row.get("productcode", "")).strip().strip('"')
            price_str = str(row.get("price", "")).strip().strip('"')

            try:
                price_val = float(price_str)
                is_price_valid = True
            except ValueError:
                price_val = 0.0
                is_price_valid = False

            is_valid = bool(post_code_raw) and bool(sku) and is_price_valid

            if is_valid:
                output_data.append({"postCode": post_code_raw, "sku": sku, "price": price_val})
                csv_data.append({"postCode": post_code_csv, "sku": sku, "price": price_val})
            else:
                error_data.append({"postCode": post_code_raw, "sku": sku, "price": price_str})

        base_name = os.path.splitext(os.path.basename(getattr(self, 'file_path', 'pasted')))[0]
        base_name_snake = base_name.lower().replace("-", "_").replace(" ", "_")
        export_folder = os.path.abspath("export")
        os.makedirs(export_folder, exist_ok=True)

        csv_path = os.path.join(export_folder, base_name_snake + "_converted.csv")
        json_path = os.path.join(export_folder, base_name_snake + "_converted.json")
        error_path = os.path.join(export_folder, base_name_snake + "_errors.csv")

        try:
            if self.export_csv_var.get():
                with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=["postCode", "sku", "price"])
                    writer.writeheader()
                    writer.writerows(csv_data)

            if self.export_json_var.get():
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(output_data, f, indent=4)

            if error_data:
                with open(error_path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=["postCode", "sku", "price"])
                    writer.writeheader()
                    writer.writerows(error_data)

            try:
                system_platform = platform.system()
                if system_platform == "Windows":
                    subprocess.Popen(f'explorer "{export_folder}"', shell=True)
                elif system_platform == "Darwin":
                    subprocess.Popen(["open", export_folder])
                elif system_platform == "Linux":
                    subprocess.Popen(["xdg-open", export_folder])
            except Exception:
                pass

            self.status_label.config(text="Files exported successfully.")
            messagebox.showinfo("Success", f"Export completed.\nFiles saved in:\n{export_folder}")
        except Exception as e:
            self.status_label.config(text="Export failed.")
            messagebox.showerror("Error", f"Failed to export:\n{e}")

# To run the app:
if __name__ == "__main__":
    root = tk.Tk()
    app = CSVConverterApp(root)
    root.mainloop()
