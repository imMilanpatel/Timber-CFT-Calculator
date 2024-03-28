import tkinter as tk
from tkinter import messagebox
import pandas as pd
from openpyxl import Workbook
from datetime import datetime

# Global constant for CFT calculation
CFT_CONSTANT = 144

class TimberCalculatorApp:
    def __init__(self, master):
        self.master = master
        master.title("Timber Wood Calculator")

        # Variables for input fields
        self.size_var = tk.StringVar()
        self.customer_name_var = tk.StringVar()
        self.customer_address_var = tk.StringVar()
        self.customer_phone_var = tk.StringVar()
        self.table_rows_var = tk.IntVar(value=10)

        # Create menu bar
        menubar = tk.Menu(master)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Export", command=self.export_to_excel)
        file_menu.add_command(label="Exit", command=self.confirm_exit)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about_info)
        help_menu.add_command(label="App Usage", command=self.show_app_usage)
        menubar.add_cascade(label="Help", menu=help_menu)

        master.config(menu=menubar)

        # Create input fields
        tk.Label(master, text="Size (e.g. 4x3):").grid(row=0, column=0, sticky="w")
        self.size_entry = tk.Entry(master, textvariable=self.size_var, width=10)
        self.size_entry.grid(row=0, column=1, sticky="w")

        tk.Label(master, text="Customer Name:").grid(row=1, column=0, sticky="w")
        self.customer_name_entry = tk.Entry(master, textvariable=self.customer_name_var, width=30)
        self.customer_name_entry.grid(row=1, column=1, columnspan=2, sticky="w")

        tk.Label(master, text="Customer Address:").grid(row=2, column=0, sticky="w")
        self.customer_address_entry = tk.Entry(master, textvariable=self.customer_address_var, width=50)
        self.customer_address_entry.grid(row=2, column=1, columnspan=2, sticky="w")

        tk.Label(master, text="Customer Phone:").grid(row=3, column=0, sticky="w")
        self.customer_phone_entry = tk.Entry(master, textvariable=self.customer_phone_var, width=15)
        self.customer_phone_entry.grid(row=3, column=1, sticky="w")

        tk.Label(master, text="Table Rows:").grid(row=4, column=0, sticky="w")
        self.table_rows_spinbox = tk.Spinbox(master, from_=1, to=100, textvariable=self.table_rows_var, width=5)
        self.table_rows_spinbox.grid(row=4, column=1, sticky="w")

        # Generate Table Button
        self.generate_table_btn = tk.Button(master, text="Generate Table", command=self.generate_table)
        self.generate_table_btn.grid(row=4, column=2, sticky="w")

        # Table Frame
        self.table_frame = tk.Frame(master)
        self.table_frame.grid(row=5, column=0, columnspan=3)

        # Calculate CFT Button
        self.calculate_cft_btn = tk.Button(master, text="Calculate CFT", command=self.calculate_cft, state=tk.DISABLED)
        self.calculate_cft_btn.grid(row=6, column=0, columnspan=3, pady=10)

    def generate_table(self):
        try:
            # Parse size input
            width, thick = map(int, self.size_var.get().split('x'))

            # Create headers
            headers = ["WIDTH", "THICK", "LENGTH", "PIECES", "CFT"]

            # Create table frame
            self.table_frame.destroy()
            self.table_frame = tk.Frame(self.master)
            self.table_frame.grid(row=5, column=0, columnspan=3)

            # Create header labels
            for col, header in enumerate(headers):
                tk.Label(self.table_frame, text=header).grid(row=0, column=col)

            # Create rows
            for row in range(self.table_rows_var.get()):
                # Width and Thick columns auto-filled
                tk.Label(self.table_frame, text=width).grid(row=row + 1, column=0)
                tk.Label(self.table_frame, text=thick).grid(row=row + 1, column=1)

                # Length, Pieces, and CFT entry boxes
                tk.Entry(self.table_frame).grid(row=row + 1, column=2)
                tk.Entry(self.table_frame).grid(row=row + 1, column=3)
                tk.Label(self.table_frame, text="").grid(row=row + 1, column=4)

            # Enable calculate button
            self.calculate_cft_btn.config(state=tk.NORMAL)

        except ValueError:
            messagebox.showerror("Error", "Invalid size input. Please enter in the format '4x3'.")

    def calculate_cft(self):
        total_cft = 0
        try:
            for row in self.table_frame.winfo_children():
                if isinstance(row, tk.Frame):
                    length = int(row.grid_slaves(row=0, column=2)[0].get())
                    pieces = int(row.grid_slaves(row=0, column=3)[0].get())
                    cft = (int(row.grid_slaves(row=0, column=0)[0].cget('text')) *
                           int(row.grid_slaves(row=0, column=1)[0].cget('text')) *
                           length * pieces) / CFT_CONSTANT
                    total_cft += cft
                    tk.Label(row, text=f"{cft:.2f}").grid(row=0, column=4)
            messagebox.showinfo("Calculation Complete", f"Total CFT: {total_cft:.2f}")
        except ValueError:
            messagebox.showerror("Error", "Invalid input. Please enter valid numbers.")

    def export_to_excel(self):
        try:
            filename = f"timber_calculator_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            df = pd.DataFrame(columns=["WIDTH", "THICK", "LENGTH", "PIECES", "CFT"])
            for row in self.table_frame.winfo_children():
                if isinstance(row, tk.Frame):
                    width = int(row.grid_slaves(row=0, column=0)[0].cget('text'))
                    thick = int(row.grid_slaves(row=0, column=1)[0].cget('text'))
                    length = int(row.grid_slaves(row=0, column=2)[0].get())
                    pieces = int(row.grid_slaves(row=0, column=3)[0].get())
                    cft = float(row.grid_slaves(row=0, column=4)[0].cget('text'))
                    df = df.append({"WIDTH": width, "THICK": thick, "LENGTH": length, "PIECES": pieces, "CFT": cft},
                                   ignore_index=True)
            df.to_excel(filename, index=False)
            messagebox.showinfo("Export Successful", f"Data exported to {filename}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred: {e}")

    def confirm_exit(self):
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.master.destroy()

    def show_about_info(self):
        messagebox.showinfo("About", "Timber Wood Calculator\nVersion 1.0\n\nDeveloper: [Developer Name]\nGitHub: [GitHub Link]")

    def show_app_usage(self):
        messagebox.showinfo("App Usage", "1. Enter the size of timber wood in the format '4x3'.\n"
                                        "2. Fill in customer details.\n"
                                        "3. Enter the number of table rows and click 'Generate Table'.\n"
                                        "4. Fill in length and pieces for each row.\n"
                                        "5. Click 'Calculate CFT' to compute CFT values.\n"
                                        "6. Use 'File' menu to export data or exit the application.")


def main():
    root = tk.Tk()
    app = TimberCalculatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()