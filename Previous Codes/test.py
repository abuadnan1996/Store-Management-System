import openpyxl
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from openpyxl import Workbook
import customtkinter as ctk

class StoreManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Store Management App")
        self.root.geometry("1000x800")
        self.product_list = []
        self.item_no_counter = 1
        self.setup_frames()
        self.setup_home_frame()
        self.setup_add_product_frame()
        self.setup_withdraw_product_frame()
        self.show_home_frame()
        self.find_missing_number()
        self.missing_numbers=[]

    def setup_frames(self):
        self.frame_home = ctk.CTkFrame(self.root)
        self.frame_add_product = ctk.CTkFrame(self.root)
        self.frame_withdraw_product = ctk.CTkFrame(self.root)

    def setup_home_frame(self):
        self.frame_home.pack(pady=10, padx=10, fill="both", expand=True)
        
        ctk.CTkLabel(self.frame_home, text="Super Petrochemical Limited", font=("Arial", 24, "bold")).pack(pady=10)
        ctk.CTkLabel(self.frame_home, text="Instrument and Control Department", font=("Arial", 15, "bold")).pack(pady=10)
        
        ctk.CTkButton(self.frame_home, text="Add Product", command=self.show_add_product_frame).pack(pady=10)
        ctk.CTkButton(self.frame_home, text="Withdraw Product", command=self.show_withdraw_product_frame).pack(pady=10)

    def setup_add_product_frame(self):
        self.frame_logo_add = ctk.CTkFrame(self.frame_add_product)
        self.frame_logo_add.pack(pady=10, padx=10, fill="x")

        self.company_name_add = ctk.CTkLabel(self.frame_logo_add, text="Super Petrochemical Limited", font=("Arial", 24, "bold"))
        self.company_name_add.pack(pady=5)
        self.department_name_add = ctk.CTkLabel(self.frame_logo_add, text="Instrument and Control Department", font=("Arial", 15, "bold"))
        self.department_name_add.pack(pady=5)

        self.frame_top_add = ctk.CTkFrame(self.frame_add_product)
        self.frame_top_add.pack(pady=10, padx=10, fill="x")
        
        self.frame_middle_add = ctk.CTkFrame(self.frame_add_product)
        self.frame_middle_add.pack(pady=10, padx=10, fill="x")
        
        self.frame_bottom_add = ctk.CTkFrame(self.frame_add_product)
        self.frame_bottom_add.pack(pady=10, padx=10, fill="both", expand=True)

        # Top Frame Widgets
        self.frame_entries_add = ctk.CTkFrame(self.frame_top_add)
        self.frame_entries_add.pack(pady=10, padx=10)

        ctk.CTkLabel(self.frame_entries_add, text="Product Name:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_product_name_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_product_name_add.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Tag:").grid(row=0, column=2, padx=5, pady=5)
        self.entry_tag_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_tag_add.grid(row=0, column=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Plant Name:").grid(row=1, column=0, padx=5, pady=5)
        self.entry_plant_name_add = ctk.CTkComboBox(self.frame_entries_add, values=["CHEMAX", "ZEHUA", "BTX", "HEXANE", "LCP", "TNS", "REFORMER 2", "CFU-10000BPD"])
        self.entry_plant_name_add.grid(row=1, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Store:").grid(row=1, column=2, padx=5, pady=5)
        self.entry_store_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_store_add.grid(row=1, column=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Type:").grid(row=2, column=0, padx=5, pady=5)
        self.entry_type_add = ctk.CTkComboBox(self.frame_entries_add, values=["PG", "PT", "LT","LG","FT","FI","TG","TT","TG","TE","FLAME","PLC ITEM","HEATER ITEM",])
        self.entry_type_add.grid(row=2, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Size:").grid(row=2, column=2, padx=5, pady=5)
        self.entry_size_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_size_add.grid(row=2, column=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Range:").grid(row=3, column=0, padx=5, pady=5)
        self.entry_range_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_range_add.grid(row=3, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Quantity:").grid(row=3, column=2, padx=5, pady=5)
        self.entry_quantity_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_quantity_add.grid(row=3, column=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Rack:").grid(row=4, column=0, padx=5, pady=5)
        self.entry_rack_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_rack_add.grid(row=4, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Column:").grid(row=4, column=2, padx=5, pady=5)
        self.entry_column_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_column_add.grid(row=4, column=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Box No:").grid(row=5, column=0, padx=5, pady=5)
        self.entry_box_no_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_box_no_add.grid(row=5, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Part No:").grid(row=5, column=2, padx=5, pady=5)
        self.entry_part_no_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_part_no_add.grid(row=5, column=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Description:").grid(row=6, column=0, padx=5, pady=5)
        self.entry_description_add = tk.Text(self.frame_entries_add, height=5, width=40)
        self.entry_description_add.grid(row=6, column=1, columnspan=3, padx=5, pady=5)

        self.button_add_product = ctk.CTkButton(self.frame_entries_add, text="Add Product", command=self.add_product)
        self.button_add_product.grid(row=7, column=0, columnspan=4, pady=10)

        # Middle Frame Widgets
        ctk.CTkLabel(self.frame_middle_add, text="Search Product:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_search_add = ctk.CTkEntry(self.frame_middle_add)
        self.entry_search_add.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkButton(self.frame_middle_add, text="Search", command=self.search_product).grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(self.frame_middle_add, text="Clear Search", command=self.refresh_product_list).grid(row=0, column=3, padx=5, pady=5)

        # Bottom Frame Widgets
        self.tree_add = ttk.Treeview(self.frame_bottom_add, columns=("Item No", "Name","Quantity","Size","Type","Range","Store","Rack","Column","Box", "Description", "Plant Name", "Part No", "Tag"), show="headings")
        self.tree_add.heading("Item No", text="Item No")
        self.tree_add.heading("Name", text="Name")
        self.tree_add.heading("Quantity", text="Quantity")
        self.tree_add.heading("Size", text="Size")
        self.tree_add.heading("Type", text="Type")
        self.tree_add.heading("Range", text="Range")
        self.tree_add.heading("Store", text="Store")
        self.tree_add.heading("Rack", text="Rack")
        self.tree_add.heading("Column", text="Column")
        self.tree_add.heading("Box", text="Box")
        self.tree_add.heading("Description", text="Description")
        self.tree_add.heading("Plant Name", text="Plant Name")
        self.tree_add.heading("Part No", text="Part No")
        self.tree_add.heading("Tag", text="Tag")

        self.tree_add.pack(fill="both", expand=True)

        self.button_export_add = ctk.CTkButton(self.frame_bottom_add, text="Export to Excel", command=self.export_to_excel)
        self.button_export_add.pack(pady=10)

    def setup_withdraw_product_frame(self):
        self.frame_withdraw_product.pack(pady=10, padx=10, fill="both", expand=True)
        
        ctk.CTkLabel(self.frame_withdraw_product, text="Withdraw Product", font=("Arial", 24, "bold")).pack(pady=10)
        
        # Add widgets to withdraw product frame here
        
    def show_home_frame(self):
        self.frame_add_product.pack_forget()
        self.frame_withdraw_product.pack_forget()
        self.frame_home.pack()

    def show_add_product_frame(self):
        self.frame_home.pack_forget()
        self.frame_withdraw_product.pack_forget()
        self.frame_add_product.pack()

    def show_withdraw_product_frame(self):
        self.frame_home.pack_forget()
        self.frame_add_product.pack_forget()
        self.frame_withdraw_product.pack()

    def add_product(self):
        product_name = self.entry_product_name_add.get()
        tag = self.entry_tag_add.get()
        plant_name = self.entry_plant_name_add.get()
        store = self.entry_store_add.get()
        type_ = self.entry_type_add.get()
        size = self.entry_size_add.get()
        range_ = self.entry_range_add.get()
        quantity = self.entry_quantity_add.get()
        rack = self.entry_rack_add.get()
        column = self.entry_column_add.get()
        box_no = self.entry_box_no_add.get()
        part_no = self.entry_part_no_add.get()
        description = self.entry_description_add.get("1.0", "end-1c")

        product = {
            "Item No": self.item_no_counter,
            "Name": product_name,
            "Tag": tag,
            "Plant Name": plant_name,
            "Store": store,
            "Type": type_,
            "Size": size,
            "Range": range_,
            "Quantity": quantity,
            "Rack": rack,
            "Column": column,
            "Box": box_no,
            "Part No": part_no,
            "Description": description
        }

        self.product_list.append(product)
        self.export_to_excel()
        self.item_no_counter += 1
        self.refresh_product_list()
        self.clear_entries()

    def search_product(self):
        query = self.entry_search_add.get().lower()
        results = [product for product in self.product_list if query in product["Name"].lower()]
        self.update_product_tree(results)

    def refresh_product_list(self):
        self.update_product_tree(self.product_list)

    def update_product_tree(self, products):
        for item in self.tree_add.get_children():
            self.tree_add.delete(item)
        for product in products:
            self.tree_add.insert("", "end", values=(
                product["Item No"], product["Name"], product["Quantity"], product["Size"], 
                product["Type"], product["Range"], product["Store"], product["Rack"], 
                product["Column"], product["Box"], product["Description"], 
                product["Plant Name"], product["Part No"], product["Tag"]
            ))

    def clear_entries(self):
        self.entry_product_name_add.delete(0, tk.END)
        self.entry_tag_add.delete(0, tk.END)
        self.entry_plant_name_add.set("")
        self.entry_store_add.delete(0, tk.END)
        self.entry_type_add.set("")
        self.entry_size_add.delete(0, tk.END)
        self.entry_range_add.delete(0, tk.END)
        self.entry_quantity_add.delete(0, tk.END)
        self.entry_rack_add.delete(0, tk.END)
        self.entry_column_add.delete(0, tk.END)
        self.entry_box_no_add.delete(0, tk.END)
        self.entry_part_no_add.delete(0, tk.END)
        self.entry_description_add.delete("1.0", tk.END)

    def find_missing_number(self):
        for product in self.product_list:
            self.missing_numbers.append(product["Item No"])
            print(self.missing_numbers)


    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Products"

        headers = ["Item No", "Name", "Quantity", "Size", "Type", "Range", "Store", "Rack", "Column", "Box", "Description", "Plant Name", "Part No", "Tag"]
        sheet.append(headers)

        for product in self.product_list:
            row = [product[header] for header in headers]
            sheet.append(row)

        try:
            workbook.save(file_path)
            messagebox.showinfo("Export Successful", f"Data exported successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"An error occurred while exporting: {e}")

if __name__ == "__main__":
    root = ctk.CTk()
    app = StoreManagementApp(root)
    root.mainloop()
