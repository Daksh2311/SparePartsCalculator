import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os

# Load data
df = pd.read_excel("SpareParts_Price_Calculator.xlsx")
df.columns = [str(col) for col in df.columns]

# Build Part No + Description labels for dropdown
part_labels = [f"{row['Part No']} - {row['Description']}" for idx, row in df.iterrows()]
part_label_to_no = {label: row['Part No'] for label, (_, row) in zip(part_labels, df.iterrows())}

# Get model columns
model_columns = [col for col in df.columns if col not in ["Part No", "Description", "Qty/pc"]]

class PriceCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Spare Parts Price Calculator with Receipt")
        self.entries = []

        # Frame for rows
        self.frame = tk.Frame(root)
        self.frame.pack(padx=10, pady=10)

        # Add first row
        self.add_entry_row()

        # Buttons
        self.add_button = tk.Button(root, text="Add Another Part", command=self.add_entry_row)
        self.add_button.pack(pady=5)

        self.calc_button = tk.Button(root, text="Calculate Total Price", command=self.calculate_total)
        self.calc_button.pack(pady=5)

        # Result label
        self.result_label = tk.Label(root, text="Total Price: 0", font=('Arial', 14, 'bold'))
        self.result_label.pack(pady=10)

    def add_entry_row(self):
        row_frame = tk.Frame(self.frame)
        row_frame.pack(pady=2)

        # Part No + Description dropdown
        part_cb = ttk.Combobox(row_frame, values=part_labels, width=40)
        part_cb.grid(row=0, column=0, padx=5)
        part_cb.set(part_labels[0])

        # Model dropdown
        model_cb = ttk.Combobox(row_frame, values=model_columns, width=10)
        model_cb.grid(row=0, column=1, padx=5)
        model_cb.set(model_columns[0])

        # Quantity entry
        qty_entry = tk.Entry(row_frame, width=10)
        qty_entry.grid(row=0, column=2, padx=5)
        qty_entry.insert(0, "1")

        self.entries.append((part_cb, model_cb, qty_entry))

    def calculate_total(self):
        total_price = 0
        receipt_data = []

        for part_cb, model_cb, qty_entry in self.entries:
            part_label = part_cb.get().strip()
            model = model_cb.get().strip()

            # Validate quantity
            try:
                quantity = int(qty_entry.get())
            except ValueError:
                quantity = 0

            if part_label not in part_label_to_no or model not in df.columns or quantity <= 0:
                continue

            part_no = part_label_to_no[part_label]
            part_row = df[df['Part No'] == part_no]

            if part_row.empty:
                continue

            # Get unit price and qty/pc
            price_per_unit = part_row.iloc[0][model]
            qty_per_pc = part_row.iloc[0]['Qty/pc']
            description = part_row.iloc[0]['Description']

            total_item_price = price_per_unit * quantity / qty_per_pc
            total_price += total_item_price

            receipt_data.append({
                "Part No": part_no,
                "Description": description,
                "Model": model,
                "Unit Price": price_per_unit,
                "Qty Ordered": quantity,
                "Qty/pc": qty_per_pc,
                "Item Total": total_item_price
            })

        self.result_label.config(text=f"Total Price: {total_price:.2f}")

        if receipt_data:
            self.generate_excel_receipt(receipt_data, total_price)
            messagebox.showinfo("Success", "Receipt has been saved!")
        else:
            messagebox.showwarning("No Valid Items", "No valid items to add to receipt.")

    def generate_excel_receipt(self, receipt_data, total_price):
        df_receipt = pd.DataFrame(receipt_data)
        df_receipt.loc[len(df_receipt.index)] = ["", "", "", "", "", "Grand Total", total_price]

        now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"SpareParts_Receipt_{now}.xlsx"
        full_path = os.path.abspath(filename)

        df_receipt.to_excel(full_path, index=False)
        print(f"✅ Receipt saved as: {full_path}")

# Run the GUI
if __name__ == "__main__":
    root = tk.Tk()
    app = PriceCalculatorApp(root)
    root.mainloop()
