import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import barcode
from barcode.writer import ImageWriter
from openpyxl import Workbook
from PIL import Image, ImageTk

class StoreManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Store Management App")
        self.root.geometry("800x800")
        
        self.product_list = []
        self.item_no_counter = 1
        
        self.setup_ui()
    
    def setup_ui(self):
        # Creating Frames
        self.frame_top = ctk.CTkFrame(self.root)
        self.frame_top.pack(pady=10, padx=10, fill="x")
        
        self.frame_middle = ctk.CTkFrame(self.root)
        self.frame_middle.pack(pady=10, padx=10, fill="x")
        
        self.frame_bottom = ctk.CTkFrame(self.root)
        self.frame_bottom.pack(pady=10, padx=10, fill="both", expand=True)

        # Company Name and Logo
        self.frame_logo = ctk.CTkFrame(self.root)
        self.frame_logo.pack(pady=10, padx=10, fill="x")

        self.company_name = ctk.CTkLabel(self.frame_logo, text="My Company", font=("Arial", 24, "bold"))
        self.company_name.pack(side="left", padx=10)

        self.logo_image = Image.open("app.png")
        self.logo_image = self.logo_image.resize((50, 50), Image.LANCZOS)
        self.logo_photo = ImageTk.PhotoImage(self.logo_image)
        self.logo_label = ctk.CTkLabel(self.frame_logo, image=self.logo_photo)
        self.logo_label.pack(side="left", padx=10)

        # Top Frame Widgets
        self.frame_entries = ctk.CTkFrame(self.frame_top)
        self.frame_entries.pack(pady=10, padx=10)

        ctk.CTkLabel(self.frame_entries, text="Product Name:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_product_name = ctk.CTkEntry(self.frame_entries)
        self.entry_product_name.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries, text="Description:").grid(row=1, column=0, padx=5, pady=5)
        self.entry_description = ctk.CTkEntry(self.frame_entries)
        self.entry_description.grid(row=1, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries, text="Plant Name:").grid(row=2, column=0, padx=5, pady=5)
        self.entry_plant_name = ctk.CTkEntry(self.frame_entries)
        self.entry_plant_name.grid(row=2, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries, text="Catalogue No:").grid(row=3, column=0, padx=5, pady=5)
        self.entry_catalogue_no = ctk.CTkEntry(self.frame_entries)
        self.entry_catalogue_no.grid(row=3, column=1, padx=5, pady=5)
        
        ctk.CTkLabel(self.frame_entries, text="Quantity:").grid(row=4, column=0, padx=5, pady=5)
        self.entry_quantity = ctk.CTkEntry(self.frame_entries)
        self.entry_quantity.grid(row=4, column=1, padx=5, pady=5)
        
        self.button_add_product = ctk.CTkButton(self.frame_entries, text="Add Product", command=self.add_product)
        self.button_add_product.grid(row=5, column=0, columnspan=2, pady=10)
        
        self.button_update_product = ctk.CTkButton(self.frame_entries, text="Update Product", command=self.update_product)
        self.button_update_product.grid(row=6, column=0, columnspan=2, pady=10)
        
        # Middle Frame Widgets
        ctk.CTkLabel(self.frame_middle, text="Search Product:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_search = ctk.CTkEntry(self.frame_middle)
        self.entry_search.grid(row=0, column=1, padx=5, pady=5)
        
        ctk.CTkButton(self.frame_middle, text="Search", command=self.search_product).grid(row=0, column=2, padx=5, pady=5)
        ctk.CTkButton(self.frame_middle, text="Clear Search", command=self.refresh_product_list).grid(row=0, column=3, padx=5, pady=5)
        
        # Bottom Frame Widgets
        self.tree = ttk.Treeview(self.frame_bottom, columns=("Item No", "Name", "Description", "Plant Name", "Catalogue No", "Quantity"), show="headings")
        self.tree.heading("Item No", text="Item No")
        self.tree.heading("Name", text="Product Name")
        self.tree.heading("Description", text="Description")
        self.tree.heading("Plant Name", text="Plant Name")
        self.tree.heading("Catalogue No", text="Catalogue No")
        self.tree.heading("Quantity", text="Quantity")
        self.tree.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.tree.bind("<Double-1>", self.load_selected_product)
        
        ctk.CTkButton(self.frame_bottom, text="Remove Product", command=self.remove_product).pack(pady=10)
        ctk.CTkButton(self.frame_bottom, text="Generate Excel", command=self.generate_excel).pack(pady=10)
        ctk.CTkButton(self.frame_bottom, text="Refresh", command=self.refresh_product_list).pack(pady=10)
    
    def add_product(self):
        name = self.entry_product_name.get()
        description = self.entry_description.get()
        plant_name = self.entry_plant_name.get()
        catalogue_no = self.entry_catalogue_no.get()
        try:
            quantity = int(self.entry_quantity.get())
        except ValueError:
            messagebox.showerror("Invalid input", "Please enter a valid quantity.")
            return
        
        if not (name and description and plant_name and catalogue_no):
            messagebox.showerror("Invalid input", "Please fill in all the fields.")
            return
        
        item_no = f"{self.item_no_counter:06d}"
        self.item_no_counter += 1
        
        self.product_list.append((item_no, name, description, plant_name, catalogue_no, quantity))
        self.generate_barcode(item_no)
        
        self.entry_product_name.delete(0, tk.END)
        self.entry_description.delete(0, tk.END)
        self.entry_plant_name.delete(0, tk.END)
        self.entry_catalogue_no.delete(0, tk.END)
        self.entry_quantity.delete(0, tk.END)
        
        self.refresh_product_list()
        messagebox.showinfo("Success", "Product added successfully.")
    
    def update_product(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Selection Error", "Please select a product to update.")
            return

        item_values = self.tree.item(selected_item, "values")
        item_no = item_values[0]
        
        name = self.entry_product_name.get()
        description = self.entry_description.get()
        plant_name = self.entry_plant_name.get()
        catalogue_no = self.entry_catalogue_no.get()
        try:
            quantity = int(self.entry_quantity.get())
        except ValueError:
            messagebox.showerror("Invalid input", "Please enter a valid quantity.")
            return
        
        if not (name and description and plant_name and catalogue_no):
            messagebox.showerror("Invalid input", "Please fill in all the fields.")
            return
        
        for index, product in enumerate(self.product_list):
            if product[0] == item_no:
                self.product_list[index] = (item_no, name, description, plant_name, catalogue_no, quantity)
                break
        
        self.entry_product_name.delete(0, tk.END)
        self.entry_description.delete(0, tk.END)
        self.entry_plant_name.delete(0, tk.END)
        self.entry_catalogue_no.delete(0, tk.END)
        self.entry_quantity.delete(0, tk.END)
        
        self.refresh_product_list()
        messagebox.showinfo("Success", f"Product with Item No {item_no} updated successfully.")
    
    def generate_barcode(self, item_no):
        barcode_format = barcode.get_barcode_class('code128')
        barcode_instance = barcode_format(item_no, writer=ImageWriter())
        filename = f"{item_no}.png"
        barcode_instance.save(filename)
        
        messagebox.showinfo("Barcode", f"Barcode for item {item_no} generated and saved as {filename}")
    
    def search_product(self):
        search_term = self.entry_search.get().lower()
        
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        for product in self.product_list:
            if (search_term in product[1].lower() or 
                search_term in product[2].lower() or 
                search_term in product[3].lower() or 
                search_term in product[4].lower()):
                self.tree.insert("", tk.END, values=product)
    
    def refresh_product_list(self):
        self.entry_search.delete(0, tk.END)
        for row in self.tree.get_children():
            self.tree.delete(row)
        
        for product in self.product_list:
            self.tree.insert("", tk.END, values=product)
    
    def remove_product(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Selection Error", "Please select a product to remove.")
            return

        item_values = self.tree.item(selected_item, "values")
        item_no = item_values[0]
        
        self.product_list = [product for product in self.product_list if product[0] != item_no]
        
        self.refresh_product_list()
        messagebox.showinfo("Success", f"Product with Item No {item_no} removed successfully.")
    
    def generate_excel(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Products"
        
        # Create headers
        headers = ["Item No", "Product Name", "Description", "Plant Name", "Catalogue No", "Quantity"]
        sheet.append(headers)
        
        for product in self.product_list:
            sheet.append(product)
        
        workbook.save("products.xlsx")
        messagebox.showinfo("Excel File", "Products saved to products.xlsx")

    def load_selected_product(self, event):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        
        item_values = self.tree.item(selected_item, "values")
        
        self.entry_product_name.delete(0, tk.END)
        self.entry_product_name.insert(0, item_values[1])
        
        self.entry_description.delete(0, tk.END)
        self.entry_description.insert(0, item_values[2])
        
        self.entry_plant_name.delete(0, tk.END)
        self.entry_plant_name.insert(0, item_values[3])
        
        self.entry_catalogue_no.delete(0, tk.END)
        self.entry_catalogue_no.insert(0, item_values[4])
        
        self.entry_quantity.delete(0, tk.END)
        self.entry_quantity.insert(0, item_values[5])

if __name__ == "__main__":
    ctk.set_appearance_mode("System")  # Modes: "System" (default), "Dark", "Light"
    ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"

    root = ctk.CTk()
    app = StoreManagementApp(root)
    root.mainloop()
