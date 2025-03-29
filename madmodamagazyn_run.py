import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import openpyxl
import datetime
from datetime import date


class InventoryApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Analiza stanu magazynu")
        self.all_data = []

        # Frame for buttons (in Polish)
        button_frame = tk.Frame(master)
        button_frame.pack(padx=10, pady=10, fill=tk.X)

        load_button = tk.Button(button_frame, text="Wybierz pliki i analizuj", command=self.load_files)
        load_button.pack(side=tk.LEFT, padx=5)

        save_button = tk.Button(button_frame, text="Zapisz do Excel", command=self.save_to_excel)
        save_button.pack(side=tk.LEFT, padx=5)

        # Frame for the detailed table (upper table)
        details_frame = tk.Frame(master)
        details_frame.pack(padx=10, pady=(10, 5), fill=tk.BOTH, expand=True)

        details_label = tk.Label(details_frame, text="Szczegóły")
        details_label.pack(anchor="w")

        # Detailed table columns: "Odzież" renamed to "Towar", with Index as an integer.
        self.details_columns = ["Towar", "Index", "Cena zakupu", "Szt.", "Rozmiar", "Cena przedaży", "Plik"]
        self.details_tree = ttk.Treeview(details_frame, columns=self.details_columns, show="headings")
        for col in self.details_columns:
            self.details_tree.heading(col, text=col)
            self.details_tree.column(col, width=100, anchor="center")
        self.details_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        details_scrollbar = tk.Scrollbar(details_frame, orient=tk.VERTICAL, command=self.details_tree.yview)
        self.details_tree.configure(yscrollcommand=details_scrollbar.set)
        details_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Frame for the summary table (lower table)
        summary_frame = tk.Frame(master)
        summary_frame.pack(padx=10, pady=(5, 5), fill=tk.BOTH, expand=True)

        summary_label = tk.Label(summary_frame, text="Podsumowanie (grupowanie po Index, Rozmiar i Towar)")
        summary_label.pack(anchor="w")

        # Summary table columns in the same order as detailed table
        self.summary_columns = ["Towar", "Index", "Cena zakupu", "Szt.", "Rozmiar", "Cena przedaży", "Plik"]
        self.summary_tree = ttk.Treeview(summary_frame, columns=self.summary_columns, show="headings")
        for col in self.summary_columns:
            self.summary_tree.heading(col, text=col)
            self.summary_tree.column(col, width=100, anchor="center")
        self.summary_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        summary_scrollbar = tk.Scrollbar(summary_frame, orient=tk.VERTICAL, command=self.summary_tree.yview)
        self.summary_tree.configure(yscrollcommand=summary_scrollbar.set)
        summary_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Frame for totals summary at the bottom of the window
        totals_frame = tk.Frame(master)
        totals_frame.pack(padx=10, pady=(5, 10), fill=tk.X)
        self.totals_label = tk.Label(totals_frame, text="Podsumowanie całkowite: ")
        self.totals_label.pack(anchor="w")

    def process_file(self, file_path):
        """
        Processes a single Excel file:
         - Reads the inventory date from cell G2 (formatted, but not used in the result).
         - Iterates over rows starting from row 5.
         - Adds a row only if:
             * The quantity ("Szt.") is greater than 0, and
             * The product name (column A) is not empty.
         - Renames the column "Komis" to "Cena przedaży".
         - If the value in column "Index" (cell C) is not numeric, it is set to 0.
         - Stores the product name under the key "Towar" and converts "Index" to an integer.
        """
        data_entries = []
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się otworzyć pliku:\n{file_path}\n{e}")
            return data_entries

        # Read and format the inventory date (not used further)
        inventory_date = ws['G2'].value
        if isinstance(inventory_date, (datetime.datetime, datetime.date)):
            inventory_date = inventory_date.strftime("%d-%m-%Y")

        # Process rows starting from row 5
        for row in ws.iter_rows(min_row=5, values_only=True):
            try:
                product_name = row[0]
                quantity = row[4]
            except IndexError:
                continue

            # Only add the row if quantity > 0 and product name is not empty
            if (quantity is not None and isinstance(quantity, (int, float)) and quantity > 0) and \
                    (product_name is not None and str(product_name).strip() != ""):
                # Convert "Index" to integer; if conversion fails, set to 0.
                try:
                    index_val = int(float(row[2]))
                except (ValueError, TypeError):
                    index_val = 0

                entry = {
                    "Towar": product_name,  # renamed from "Odzież"
                    "Index": index_val,
                    "Cena zakupu": row[3],
                    "Szt.": quantity,
                    "Rozmiar": row[5],
                    "Cena przedaży": row[6],  # renamed from "Komis"
                    "Plik": os.path.basename(file_path)
                }
                data_entries.append(entry)
        return data_entries

    def load_files(self):
        """
        Allows the user to select Excel files, processes each file,
         and stores the results. Then, it sorts the data numerically by the "Index" column,
         and updates both the detailed and summary tables.
        """
        file_paths = filedialog.askopenfilenames(
            title="Wybierz pliki Excel",
            filetypes=[("Pliki Excel", "*.xlsx")]
        )
        if not file_paths:
            return

        self.all_data = []
        for file_path in file_paths:
            entries = self.process_file(file_path)
            self.all_data.extend(entries)

        if not self.all_data:
            messagebox.showinfo("Informacja", "Nie znaleziono żadnych pozycji spełniających warunki.")
            return

        # Sort data by the "Index" column numerically
        self.all_data.sort(key=lambda entry: entry["Index"])
        self.update_details_tree()
        self.update_summary_tree()

    def update_details_tree(self):
        """Updates the detailed table with data from all_data."""
        for row in self.details_tree.get_children():
            self.details_tree.delete(row)
        for entry in self.all_data:
            self.details_tree.insert("", tk.END, values=(
                entry["Towar"],
                entry["Index"],
                entry["Cena zakupu"],
                entry["Szt."],
                entry["Rozmiar"],
                entry["Cena przedaży"],
                entry["Plik"]
            ))

    def generate_summary(self):
        """
        Groups data by the tuple (Index, Rozmiar, Towar) and sums the "Szt." values.
         Also, sums the values for "Cena zakupu" and "Cena przedaży",
         and collects file names from which the entries came.
         Returns a list of dictionaries with keys: "Towar", "Index", "Cena zakupu", "Szt.",
         "Rozmiar", "Cena przedaży", "Plik".
        """
        summary = {}
        for entry in self.all_data:
            key = (entry["Index"], entry["Rozmiar"], entry["Towar"])
            if key not in summary:
                summary[key] = {
                    "Index": entry["Index"],
                    "Rozmiar": entry["Rozmiar"],
                    "Towar": entry["Towar"],
                    "Szt.": 0,
                    "Cena zakupu": 0,
                    "Cena przedaży": 0,
                    "Plik": set()  # Use a set to avoid duplicate file names
                }
            summary[key]["Szt."] += entry["Szt."]
            try:
                summary[key]["Cena zakupu"] += float(entry["Cena zakupu"]) if entry["Cena zakupu"] is not None else 0
            except (ValueError, TypeError):
                pass
            try:
                summary[key]["Cena przedaży"] += float(entry["Cena przedaży"]) if entry[
                                                                                      "Cena przedaży"] is not None else 0
            except (ValueError, TypeError):
                pass
            summary[key]["Plik"].add(entry["Plik"])
        # Convert the file names set to a comma-separated string and transform the dictionary to a list
        result = []
        for key, data in summary.items():
            data["Plik"] = ", ".join(sorted(data["Plik"]))
            result.append(data)
        return result

    def update_summary_tree(self):
        """Updates the summary table with grouped data and updates the totals label."""
        for row in self.summary_tree.get_children():
            self.summary_tree.delete(row)
        summary_data = self.generate_summary()
        # Sort summary by Index, then by Rozmiar, then by Towar
        summary_data.sort(key=lambda x: (x["Index"], x["Rozmiar"], x["Towar"]))
        total_quantity = 0
        total_purchase = 0
        total_sale = 0
        for entry in summary_data:
            self.summary_tree.insert("", tk.END, values=(
                entry["Towar"],
                entry["Index"],
                entry["Cena zakupu"],
                entry["Szt."],
                entry["Rozmiar"],
                entry["Cena przedaży"],
                entry["Plik"]
            ))
            total_quantity += entry["Szt."]
            total_purchase += entry["Cena zakupu"]
            total_sale += entry["Cena przedaży"]
        # Update the totals label at the bottom of the window
        self.totals_label.config(text=f"Podsumowanie całkowite: Ilość = {total_quantity}, "
                                      f"Wartość zakupu = {total_purchase:.2f}, "
                                      f"Wartość sprzedaży = {total_sale:.2f}")

    def adjust_column_widths(self, worksheet):
        """
        Adjusts the column widths in the given worksheet based on the maximum length
        of the content in each column.
        """
        from openpyxl.utils import get_column_letter
        for col in worksheet.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            worksheet.column_dimensions[col_letter].width = max_length + 2

    def save_to_excel(self):
        """
        Saves the data displayed in both tables to an Excel file.
         The first worksheet ("Szczegóły") contains the detailed data,
         and the second worksheet ("Suma") contains the summary data.
         The default filename is "stan_magazynu_YYYY-MM-DD.xlsx".
         Column widths are automatically adjusted.
        """
        if not self.details_tree.get_children():
            messagebox.showwarning("Brak danych", "Brak danych do zapisu. Najpierw przetwórz pliki.")
            return

        today_str = date.today().strftime("%Y-%m-%d")
        default_filename = f"stan_magazynu_{today_str}.xlsx"
        save_path = filedialog.asksaveasfilename(
            title="Zapisz zbiorczą tabelę",
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Pliki Excel", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            from openpyxl import Workbook
            wb = Workbook()

            # Worksheet "Szczegóły" with detailed data
            ws_details = wb.active
            ws_details.title = "Szczegóły"
            headers_details = self.details_columns
            ws_details.append(headers_details)
            for child in self.details_tree.get_children():
                row = self.details_tree.item(child)['values']
                ws_details.append(row)
            # Adjust column widths for the detailed sheet
            self.adjust_column_widths(ws_details)

            # Worksheet "Suma" with summary data
            ws_summary = wb.create_sheet(title="Suma")
            headers_summary = self.summary_columns
            ws_summary.append(headers_summary)
            summary_data = self.generate_summary()
            summary_data.sort(key=lambda x: (x["Index"], x["Rozmiar"], x["Towar"]))
            for entry in summary_data:
                ws_summary.append([
                    entry["Towar"],
                    entry["Index"],
                    entry["Cena zakupu"],
                    entry["Szt."],
                    entry["Rozmiar"],
                    entry["Cena przedaży"],
                    entry["Plik"]
                ])
            # Adjust column widths for the summary sheet
            self.adjust_column_widths(ws_summary)

            wb.save(save_path)
            messagebox.showinfo("Sukces", f"Dane zostały zapisane do pliku:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Błąd", f"Nie udało się zapisać pliku:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap("app_icon.ico")
    app = InventoryApp(root)
    root.mainloop()
