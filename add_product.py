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

