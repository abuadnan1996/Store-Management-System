import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from openpyxl import Workbook
from PIL import Image, ImageTk
from tkinter import filedialog
import os

class StoreManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Store Management App")
        self.root.geometry("1000x800")
        
        self.product_list = []
        self.item_no_counter = 1
        
        self.setup_ui()
    
    def setup_ui(self):
        # Home Page
        self.frame_home = ctk.CTkFrame(self.root)
        self.frame_home.pack(pady=10, padx=10, fill="both", expand=True)
        
        ctk.CTkLabel(self.frame_home, text="Welcome to Store Management App", font=("Arial", 24, "bold")).pack(pady=20)
        
        ctk.CTkButton(self.frame_home, text="Add Product", command=self.show_add_product_frame).pack(pady=10)
        ctk.CTkButton(self.frame_home, text="Withdraw Product", command=self.show_withdraw_product_frame).pack(pady=10)
        
        # Add Product Frame
        self.frame_add_product = ctk.CTkFrame(self.root)
        
        # Withdraw Product Frame
        self.frame_withdraw_product = ctk.CTkFrame(self.root)
        
        self.setup_add_product_frame()
        self.setup_withdraw_product_frame()
    def setup_add_product_frame(self):
        # Creating Frames
        # Company Name and Logo
        self.frame_logo = ctk.CTkFrame(self.frame_add_product)
        self.frame_logo.pack(pady=10, padx=10, fill="x")

        self.company_name = ctk.CTkLabel(self.frame_logo, text="Super Petrochemical Limited", font=("Arial", 24, "bold"))
        self.company_name.pack(pady=5)
        self.department_name = ctk.CTkLabel(self.frame_logo, text="Instrument and Control Department", font=("Arial", 15, "bold"))
        self.department_name.pack(pady=5)

        self.frame_top = ctk.CTkFrame(self.frame_add_product)
        self.frame_top.pack(pady=10, padx=10, fill="x")
        
        self.frame_middle = ctk.CTkFrame(self.frame_add_product)
        self.frame_middle.pack(pady=10, padx=10, fill="x")
        
        self.frame_bottom = ctk.CTkFrame(self.frame_add_product)
        self.frame_bottom.pack(pady=10, padx=10, fill="both", expand=True)

        # Top Frame Widgets
        self.frame_entries_add = ctk.CTkFrame(self.frame_top)
        self.frame_entries_add.pack(pady=10, padx=10)

        ctk.CTkLabel(self.frame_entries_add, text="Product Name:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_product_name_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_product_name_add.grid(row=0, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Tag:").grid(row=0, column=2, padx=5, pady=5)
        self.entry_tag_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_tag_add.grid(row=0, column=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Plant Name:").grid(row=1, column=0, padx=5, pady=5)
        self.entry_plant_name_add = ctk.CTkComboBox(self.frame_entries_add, values=["CHEMAX", "ZEHUA", "BTX","HEXANE","LCP","TNS","REFORMER 2","CFU-10000BPD",])
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

        ctk.CTkLabel(self.frame_entries_add, text="Description:").grid(row=5, column=0, padx=5, pady=5)
        self.entry_description_add = ctk.CTkTextbox(self.frame_entries_add, height=50, width=400)
        self.entry_description_add.grid(row=5, column=1, columnspan=3, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Part No:").grid(row=6, column=0, padx=5, pady=5)
        self.entry_part_no_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_part_no_add.grid(row=6, column=1, padx=5, pady=5)

        ctk.CTkLabel(self.frame_entries_add, text="Box No:").grid(row=6, column=2, padx=5, pady=5)
        self.entry_box_no_add = ctk.CTkEntry(self.frame_entries_add)
        self.entry_box_no_add.grid(row=6, column=3, padx=5, pady=5)
        
        self.button_add_product = ctk.CTkButton(self.frame_entries_add, text="Add Product", command=self.add_product)
        self.button_add_product.grid(row=7, column=0, columnspan=4, pady=10)
        
        self.button_update_product = ctk.CTkButton(self.frame_entries_add, text="Update Product", command=self.update_product)
        self.button_update_product.grid(row=8, column=0, columnspan=4, pady=10)

        # Middle Frame Widgets
        ctk.CTkLabel(self.frame_middle, text="Search Product:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_search_add = ctk.CTkEntry(self.frame_middle)
        self.entry_search_add.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkButton(self.frame_middle, text="Search", command=self.search_product).grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(self.frame_middle, text="Clear Search", command=self.refresh_product_list).grid(row=0, column=3, padx=5, pady=5)

        # Bottom Frame Widgets
        self.tree_add = ttk.Treeview(self.frame_bottom, columns=("Item No", "Name","Quantity","Type","Range","Store","Rack","Column","Box", "Description", "Plant Name", "Part No","Tag"), show="headings")
        self.tree_add.heading("Item No", text="Item No")
        self.tree_add.heading("Name", text="Product Name")
        self.tree_add.heading("Quantity", text="Quantity")
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

        self.tree_add.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.tree_add.bind("<Double-1>", self.load_selected_product_add)
        
        ctk.CTkButton(self.frame_bottom, text="Back to Home", command=self.show_home_frame).pack(pady=10)

    def setup_withdraw_product_frame(self):
        # Implement withdraw product frame similar to add product frame
        # Withdraw Product Frame Setup
        ctk.CTkLabel(self.frame_withdraw_product, text="Withdraw Product functionality to be implemented.", font=("Arial", 16)).pack(pady=20)
        self.frame_logo = ctk.CTkFrame(self.frame_withdraw_product)
        self.frame_logo.pack(pady=10, padx=10, fill="x")

        self.company_name = ctk.CTkLabel(self.frame_logo, text="Super Petrochemical Limited", font=("Arial", 24, "bold"))
        self.company_name.pack(pady=10)

        self.frame_top = ctk.CTkFrame(self.frame_withdraw_product)
        self.frame_top.pack(pady=10, padx=10, fill="x")
        
        self.frame_middle = ctk.CTkFrame(self.frame_withdraw_product)
        self.frame_middle.pack(pady=10, padx=10, fill="x")
        
        self.frame_bottom = ctk.CTkFrame(self.frame_withdraw_product)
        self.frame_bottom.pack(pady=10, padx=10, fill="both", expand=True)
        # textbox = customtkinter.CTkTextbox(app)

        # Top Frame Widgets
        self.frame_entries_withdraw = ctk.CTkFrame(self.frame_top)
        self.frame_entries_withdraw.pack(pady=10, padx=10)

        ctk.CTkLabel(self.frame_entries_withdraw, text="Product Name:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_product_name_withdraw = ctk.CTkEntry(self.frame_entries_withdraw)
        self.entry_product_name_withdraw.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries_withdraw, text="Description:").grid(row=1, column=0, padx=5, pady=5)
        self.entry_description_withdraw = ctk.CTkEntry(self.frame_entries_withdraw)
        self.entry_description_withdraw.grid(row=1, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries_withdraw, text="Plant Name:").grid(row=2, column=0, padx=5, pady=5)
        self.entry_plant_name_withdraw = ctk.CTkEntry(self.frame_entries_withdraw)
        self.entry_plant_name_withdraw.grid(row=2, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries_withdraw, text="Part No:").grid(row=3, column=0, padx=5, pady=5)
        self.entry_Part_no_withdraw = ctk.CTkEntry(self.frame_entries_withdraw)
        self.entry_Part_no_withdraw.grid(row=3, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries_withdraw, text="Quantity:").grid(row=4, column=0, padx=5, pady=5)
        self.entry_quantity_withdraw = ctk.CTkEntry(self.frame_entries_withdraw)
        self.entry_quantity_withdraw.grid(row=4, column=1, padx=5, pady=5)
        
        self.button_withdraw_product = ctk.CTkButton(self.frame_entries_withdraw, text="Withdraw Product", command=self.withdraw_product)
        self.button_withdraw_product.grid(row=5, column=0, columnspan=2, pady=10)
        
        self.button_update_product_withdraw = ctk.CTkButton(self.frame_entries_withdraw, text="Update Product", command=self.update_product)
        self.button_update_product_withdraw.grid(row=6, column=0, columnspan=2, pady=10)
        
        # Middle Frame Widgets
        ctk.CTkLabel(self.frame_middle, text="Search Product:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_search_withdraw = ctk.CTkEntry(self.frame_middle)
        self.entry_search_withdraw.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkButton(self.frame_middle, text="Search", command=self.search_product).grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(self.frame_middle, text="Clear Search", command=self.refresh_product_list).grid(row=0, column=3, padx=5, pady=5)
        
        # Bottom Frame Widgets
        self.tree_withdraw = ttk.Treeview(self.frame_bottom, columns=("Item No", "Name", "Description", "Plant Name", "Part No", "Quantity"), show="headings")
        self.tree_withdraw.heading("Item No", text="Item No")
        self.tree_withdraw.heading("Name", text="Product Name")
        self.tree_withdraw.heading("Description", text="Description")
        self.tree_withdraw.heading("Plant Name", text="Plant Name")
        self.tree_withdraw.heading("Part No", text="Part No")
        self.tree_withdraw.heading("Quantity", text="Quantity")
        self.tree_withdraw.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.tree_withdraw.bind("<Double-1>", self.load_selected_product_withdraw)
        
        ctk.CTkButton(self.frame_bottom, text="Back to Home", command=self.show_home_frame).pack(pady=10)


    def show_home_frame(self):
        self.frame_add_product.pack_forget()
        self.frame_withdraw_product.pack_forget()
        self.frame_home.pack(pady=10, padx=10, fill="both", expand=True)
    
    def show_add_product_frame(self):
        self.frame_home.pack_forget()
        self.frame_withdraw_product.pack_forget()
        self.frame_add_product.pack(pady=10, padx=10, fill="both", expand=True)
        self.refresh_product_list()
    
    def show_withdraw_product_frame(self):
        self.frame_home.pack_forget()
        self.frame_add_product.pack_forget()
        self.frame_withdraw_product.pack(pady=10, padx=10, fill="both", expand=True)
        self.refresh_product_list()


    def print_product_list(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                 filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                                                 title="Choose file location")
        if file_path:
            try:
                with open(file_path, 'w') as file:
                    file.write("Item No\tProduct Name\tTag\tPlant Name\tStore\tType\tSize\tRange\tQuantity\tRack\tColumn\tDescription\tPart No\tBox No\n")
                    for product in self.product_list:
                        product_details = f"{product['Item No']}\t{product['Name']}\t{product['Tag']}\t{product['Plant Name']}\t{product['Store']}\t{product['Type']}\t{product['Size']}\t{product['Range']}\t{product['Quantity']}\t{product['Rack']}\t{product['Column']}\t{product['Description']}\t{product['Part No']}\t{product['Box No']}\n"
                        file.write(product_details)
                os.startfile(file_path, "print")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to print file: {e}")

    # Add this method to your StoreManagementApp class

    # Example of how to integrate the print functionality in the UI
        

    def add_product(self, event=None):
        name = self.entry_product_name_add.get()
        quantity = self.entry_quantity_add.get()
        item_type = self.entry_type_add.get()
        item_range = self.entry_range_add.get()
        store = self.entry_store_add.get()
        rack = self.entry_rack_add.get()
        column = self.entry_column_add.get()
        box = self.entry_box_no_add.get()
        description = self.entry_description_add.get("1.0", "end").strip()
        plant_name = self.entry_plant_name_add.get()
        part_no = self.entry_part_no_add.get()
        
        
        if not name or not description or not plant_name or not part_no or not quantity:
            messagebox.showerror("Input Error", "All fields must be filled")
            return
        
        product = {
            "Item No": self.item_no_counter,
            "Name": self.entry_product_name_add.get(),
            "Quantity": self.entry_quantity_add.get(),
            "Type": self.entry_type_add.get(),
            "Range": self.entry_range_add.get(),
            "Store": self.entry_store_add.get(),
            "Rack": self.entry_rack_add.get(),
            "Column": self.entry_column_add.get(),
            "Box No": self.entry_box_no_add.get(),
            "Description": self.entry_description_add.get("1.0", "end").strip(),
            "Plant Name": self.entry_plant_name_add.get(),
            "Part No": self.entry_part_no_add.get(),           
            "Size": self.entry_size_add.get(),
            "Tag": self.entry_tag_add.get()
            
        }
        
        self.product_list.append(product)
        self.item_no_counter += 1
        self.refresh_product_list()
        self.clear_entries_add()
    
    def update_product(self):
        selected_item = self.tree_add.selection()
        if not selected_item:
            messagebox.showerror("Selection Error", "No item selected to update")
            return

        item_index = self.tree_add.index(selected_item)
        self.product_list[item_index] = {
            "Item No": self.tree_add.item(selected_item)['values'][0],
            "Name": self.entry_product_name_add.get(),
            "Description": self.entry_description_add.get(),
            "Plant Name": self.entry_plant_name_add.get(),
            "Part No": self.entry_Part_no_add.get(),
            "Quantity": self.entry_quantity_add.get()
        }
        
        self.refresh_product_list()
        self.clear_entries_add()
    
    def withdraw_product(self):
        selected_item = self.tree_withdraw.selection()
        if not selected_item:
            messagebox.showerror("Selection Error", "No item selected to withdraw")
            return

        item_index = self.tree_withdraw.index(selected_item)
        self.product_list[item_index]["Quantity"] = str(int(self.product_list[item_index]["Quantity"]) - int(self.entry_quantity_withdraw.get()))
        
        self.refresh_product_list()
        self.clear_entries_withdraw()
    
    def load_selected_product_add(self, event):
        selected_item = self.tree_add.selection()
        if not selected_item:
            return

        item_index = self.tree_add.index(selected_item)
        selected_product = self.product_list[item_index]
        
        self.entry_product_name_add.delete(0, tk.END)
        self.entry_product_name_add.insert(0, selected_product["Name"])
        self.entry_quantity_add.delete(0, tk.END)
        self.entry_quantity_add.insert(0, selected_product["Quantity"])
        self.entry_type_add.set(selected_product["Type"])
        self.entry_range_add.delete(0, tk.END)
        self.entry_range_add.insert(0, selected_product["Range"])
        self.entry_store_add.delete(0, tk.END)
        self.entry_store_add.insert(0, selected_product["Store"])
        self.entry_rack_add.delete(0, tk.END)
        self.entry_rack_add.insert(0, selected_product["Rack"])
        self.entry_column_add.delete(0, tk.END)
        self.entry_column_add.insert(0, selected_product["Column"])
        self.entry_box_no_add.delete(0, tk.END)
        self.entry_box_no_add.insert(0, selected_product["Box No"])
        self.entry_plant_name_add.set(selected_product["Plant Name"])
        self.entry_part_no_add.delete(0, tk.END)
        self.entry_part_no_add.insert(0, selected_product["Part No"])
        self.entry_tag_add.delete(0, tk.END)
        self.entry_tag_add.insert(0, selected_product["Tag"])
        self.entry_size_add.delete(0, tk.END)
        self.entry_size_add.insert(0, selected_product["Size"])
        self.entry_description_add.delete('1.0', 'end')
        self.entry_description_add.insert("1.0", selected_product["Description"],tags=None)

    def load_selected_product_withdraw(self, event):
        selected_item = self.tree_withdraw.selection()
        if not selected_item:
            return

        item_index = self.tree_withdraw.index(selected_item)
        selected_product = self.product_list[item_index]
        
        self.entry_product_name_withdraw.delete(0, tk.END)
        self.entry_product_name_withdraw.insert(0, selected_product["Name"])
        self.entry_description_withdraw.delete(0, tk.END)
        self.entry_description_withdraw.insert(0, selected_product["Description"])
        self.entry_plant_name_withdraw.delete(0, tk.END)
        self.entry_plant_name_withdraw.insert(0, selected_product["Plant Name"])
        self.entry_Part_no_withdraw.delete(0, tk.END)
        self.entry_Part_no_withdraw.insert(0, selected_product["Part No"])
        self.entry_quantity_withdraw.delete(0, tk.END)
        self.entry_quantity_withdraw.insert(0, selected_product["Quantity"])
    
    def refresh_product_list(self):
        for tree in [self.tree_add, self.tree_withdraw]:
            for i in tree.get_children():
                tree.delete(i)

            for product in self.product_list:
                tree.insert("", "end", values=(product["Item No"], product["Name"], product["Quantity"], product["Type"], product["Range"], product["Store"],product["Rack"],product["Column"],product["Box No"],product["Description"],product["Plant Name"],product["Part No"],product["Tag"]))
    def search_product(self):
        search_term = self.entry_search_add.get().lower()
        
        for row in self.tree_add.get_children():
            self.tree_add.delete(row)
        
        for product in self.product_list:
            if (search_term in product[1].lower() or 
                search_term in product[2].lower() or 
                search_term in product[3].lower() or
                search_term in product[4].lower() or
                search_term in product[5].lower() or
                search_term in product[6].lower() or
                search_term in product[7].lower() or
                search_term in product[8].lower() or
                search_term in product[9].lower() or
                search_term in product[10].lower() or 
                search_term in product[11].lower() or 
                search_term in product[12].lower() or 
                search_term in product[13].lower()):
                self.tree_add.insert("", tk.END, values=product)

    # def search_product(self):
    #     search_term = self.entry_search_add.get()
    #     filtered_list = [product for product in self.product_list if search_term.lower() in product["Name"].lower()]

    #     for tree in [self.tree_add, self.tree_withdraw]:
    #         for i in tree.get_children():
    #             tree.delete(i)
                
    #         for product in filtered_list:
    #             tree.insert("", "end", values=(product["Item No"], product["Name"],product["Quantity"], product["Description"], product["Plant Name"], product["Part No"], product["Quantity"]))

    def clear_entries_add(self):
        self.entry_product_name_add.delete(0, 'end')
        self.entry_tag_add.delete(0, 'end')
        self.entry_plant_name_add.set("")
        self.entry_store_add.delete(0, 'end')
        self.entry_type_add.set("")
        self.entry_size_add.delete(0, 'end')
        self.entry_range_add.delete(0, 'end')
        self.entry_quantity_add.delete(0, 'end')
        self.entry_rack_add.delete(0, 'end')
        self.entry_column_add.delete(0, 'end')
        self.entry_description_add.delete('1.0', 'end')
        self.entry_part_no_add.delete(0, 'end')
        self.entry_box_no_add.delete(0, 'end')
    
    def clear_entries_withdraw(self):
        self.entry_product_name_withdraw.delete(0, tk.END)
        self.entry_description_withdraw.delete(0, tk.END)
        self.entry_plant_name_withdraw.delete(0, tk.END)
        self.entry_Part_no_withdraw.delete(0, tk.END)
        self.entry_quantity_withdraw.delete(0, tk.END)


if __name__ == "__main__":
    ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
    ctk.set_default_color_theme("green")  # Themes: "blue" (default), "green", "dark-blue"
    root = ctk.CTk()
    app = StoreManagementApp(root)
    root.mainloop()