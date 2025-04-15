import csv
import os
from tkinter import *
from tkinter import messagebox, ttk, filedialog
from openpyxl import Workbook
from datetime import datetime

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GrosirDimas - Inventory Management System")
        self.root.geometry("950x650")
        self.root.configure(bg="#f0f0f0")
        self.root.resizable(True, True)
        
        # Variables
        self.produk_list = []
        self.id_counter = 1
        self.csv_file = "produk.csv"
        self.selected_item = None
        self.search_var = StringVar()
        self.search_var.trace("w", self.search_products)
        
        # Load Data
        self.load_data()
        
        # Set up the GUI
        self.setup_gui()
        
    def load_data(self):
        if os.path.exists(self.csv_file):
            with open(self.csv_file, mode='r', newline='', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                for row in reader:
                    produk = {
                        "id": int(row["id"]),
                        "nama": row["nama"],
                        "stok": int(row["stok"]),
                        "harga": int(row["harga"])
                    }
                    self.produk_list.append(produk)
                if self.produk_list:
                    self.id_counter = max(p["id"] for p in self.produk_list) + 1

    def simpan_data(self):
        with open(self.csv_file, mode='w', newline='', encoding='utf-8') as file:
            fieldnames = ["id", "nama", "stok", "harga"]
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            for produk in self.produk_list:
                writer.writerow(produk)
        self.status_var.set(f"Data disimpan: {datetime.now().strftime('%H:%M:%S')}")

    def clear_entries(self):
        self.entry_nama.delete(0, END)
        self.entry_stok.delete(0, END)
        self.entry_harga.delete(0, END)
        self.selected_item = None
        self.btn_tambah.config(text="Tambah Produk", bg="#4CAF50", command=self.tambah_produk)
        self.btn_batal.config(state=DISABLED)

    def validate_inputs(self):
        nama = self.entry_nama.get().strip()
        stok = self.entry_stok.get().strip()
        harga = self.entry_harga.get().strip()
        
        error_messages = []
        
        if not nama:
            error_messages.append("Nama produk tidak boleh kosong")
        
        if not stok:
            error_messages.append("Stok tidak boleh kosong")
        else:
            try:
                stok_val = int(stok)
                if stok_val < 0:
                    error_messages.append("Stok tidak boleh negatif")
            except ValueError:
                error_messages.append("Stok harus berupa angka")
        
        if not harga:
            error_messages.append("Harga tidak boleh kosong")
        else:
            try:
                harga_val = int(harga)
                if harga_val < 0:
                    error_messages.append("Harga tidak boleh negatif")
            except ValueError:
                error_messages.append("Harga harus berupa angka")
        
        if error_messages:
            messagebox.showwarning("Input Error", "\n".join(error_messages))
            return False
        
        return True

    def tambah_produk(self):
        if not self.validate_inputs():
            return
            
        nama = self.entry_nama.get().strip()
        stok = int(self.entry_stok.get())
        harga = int(self.entry_harga.get())

        produk = {
            "id": self.id_counter,
            "nama": nama,
            "stok": stok,
            "harga": harga
        }

        self.produk_list.append(produk)
        self.id_counter += 1
        self.simpan_data()
        self.update_listbox()
        self.clear_entries()
        messagebox.showinfo("Sukses", f"Produk '{nama}' berhasil ditambahkan!")
        self.status_var.set(f"Produk '{nama}' ditambahkan")

    def edit_produk(self):
        if not self.validate_inputs() or not self.selected_item:
            return
            
        idx = self.selected_item
        produk = self.produk_list[idx]
        
        nama = self.entry_nama.get().strip()
        stok = int(self.entry_stok.get())
        harga = int(self.entry_harga.get())
        
        produk["nama"] = nama
        produk["stok"] = stok
        produk["harga"] = harga
        
        self.simpan_data()
        self.update_listbox()
        self.clear_entries()
        messagebox.showinfo("Sukses", f"Produk '{nama}' berhasil diupdate!")
        self.status_var.set(f"Produk '{nama}' diupdate")

    def hapus_produk(self):
        if not self.selected_item and not self.tree.selection():
            messagebox.showwarning("Peringatan", "Silakan pilih produk terlebih dahulu!")
            return
            
        if not self.selected_item and self.tree.selection():
            selected_id = self.tree.item(self.tree.selection()[0])['values'][0]
            for i, produk in enumerate(self.produk_list):
                if produk['id'] == selected_id:
                    self.selected_item = i
                    break
        
        idx = self.selected_item
        produk = self.produk_list[idx]
        
        confirm = messagebox.askyesno("Konfirmasi", f"Yakin ingin menghapus produk '{produk['nama']}'?")
        if confirm:
            del self.produk_list[idx]
            self.simpan_data()
            self.update_listbox()
            self.clear_entries()
            self.status_var.set(f"Produk '{produk['nama']}' dihapus")

    def on_item_select(self, event):
        if not self.tree.selection():
            return
            
        selected_id = self.tree.item(self.tree.selection()[0])['values'][0]
        
        for i, produk in enumerate(self.produk_list):
            if produk['id'] == selected_id:
                self.entry_nama.delete(0, END)
                self.entry_stok.delete(0, END)
                self.entry_harga.delete(0, END)
                
                self.entry_nama.insert(0, produk['nama'])
                self.entry_stok.insert(0, produk['stok'])
                self.entry_harga.insert(0, produk['harga'])
                
                self.selected_item = i
                self.btn_tambah.config(text="Update Produk", bg="#FF9800", command=self.edit_produk)
                self.btn_batal.config(state=NORMAL)
                break

    def search_products(self, *args):
        search_text = self.search_var.get().lower()
        
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        for produk in self.produk_list:
            if search_text in produk['nama'].lower() or search_text in str(produk['id']):
                self.tree.insert("", END, values=(produk['id'], produk['nama'], produk['stok'], 
                                                f"Rp {produk['harga']:,}".replace(',', '.')))

    def update_listbox(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for produk in self.produk_list:
            self.tree.insert("", END, values=(produk['id'], produk['nama'], produk['stok'], 
                                            f"Rp {produk['harga']:,}".replace(',', '.')))
        self.tree.update()

    def export_ke_excel(self):
        if not self.produk_list:
            messagebox.showerror("Gagal", "Tidak ada data untuk diekspor.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Simpan File Excel"
        )
        
        if not file_path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Produk GrosirDimas"
        
        # Add header with style
        header = ["ID", "Nama Produk", "Stok", "Harga"]
        ws.append(header)
        
        # Add data
        for produk in self.produk_list:
            ws.append([produk["id"], produk["nama"], produk["stok"], produk["harga"]])
        
        # Auto-adjust column width
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(file_path)
        messagebox.showinfo("Sukses", f"Data berhasil diekspor ke '{file_path}'")
        self.status_var.set(f"Data diekspor ke Excel")

    def treeview_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        
        # Convert numeric columns to proper types for sorting
        if col in ("ID", "Stok"):
            l = [(int(i[0]), i[1]) for i in l]
        elif col == "Harga":
            l = [(int(i[0].replace("Rp ", "").replace(".", "")), i[1]) for i in l]
            
        l.sort(reverse=reverse)

        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # Reverse sort next time
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def setup_gui(self):
        # Main frame
        main_frame = Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=BOTH, expand=True, padx=20, pady=10)
        
        # Title
        title_frame = Frame(main_frame, bg="#f0f0f0")
        title_frame.pack(fill=X, pady=10)
        
        Label(title_frame, text="GrosirDimas - Inventory Management System", 
              font=("Helvetica", 16, "bold"), bg="#f0f0f0", fg="#333333").pack()
        
        # Content frame (split into left and right)
        content_frame = Frame(main_frame, bg="#f0f0f0")
        content_frame.pack(fill=BOTH, expand=True)
        
        # Left frame (entry form)
        left_frame = LabelFrame(content_frame, text="Detail Produk", bg="#f0f0f0", padx=15, pady=15)
        left_frame.pack(side=LEFT, fill=BOTH, expand=False, padx=(0,10))
        
        # Product form
        Label(left_frame, text="Nama Produk", bg="#f0f0f0", anchor="w").grid(row=0, column=0, sticky="w", pady=(10,5))
        self.entry_nama = Entry(left_frame, width=30, font=("Helvetica", 10))
        self.entry_nama.grid(row=1, column=0, sticky="ew", pady=(0,10))
        
        Label(left_frame, text="Stok", bg="#f0f0f0", anchor="w").grid(row=2, column=0, sticky="w", pady=(10,5))
        self.entry_stok = Entry(left_frame, width=30, font=("Helvetica", 10))
        self.entry_stok.grid(row=3, column=0, sticky="ew", pady=(0,10))
        
        Label(left_frame, text="Harga (Rp)", bg="#f0f0f0", anchor="w").grid(row=4, column=0, sticky="w", pady=(10,5))
        self.entry_harga = Entry(left_frame, width=30, font=("Helvetica", 10))
        self.entry_harga.grid(row=5, column=0, sticky="ew", pady=(0,10))
        
        # Buttons frame
        buttons_frame = Frame(left_frame, bg="#f0f0f0")
        buttons_frame.grid(row=6, column=0, sticky="ew", pady=20)
        
        self.btn_tambah = Button(buttons_frame, text="Tambah Produk", command=self.tambah_produk,
                               bg="#4CAF50", fg="white", font=("Helvetica", 10, "bold"),
                               width=15, relief=RAISED, bd=2)
        self.btn_tambah.pack(side=LEFT, padx=(0,5))
        
        self.btn_batal = Button(buttons_frame, text="Batal", command=self.clear_entries,
                              bg="#f44336", fg="white", font=("Helvetica", 10),
                              width=8, relief=RAISED, bd=2, state=DISABLED)
        self.btn_batal.pack(side=LEFT)
        
        self.btn_hapus = Button(left_frame, text="Hapus Produk", command=self.hapus_produk,
                              bg="#f44336", fg="white", font=("Helvetica", 10),
                              width=15, relief=RAISED, bd=2)
        self.btn_hapus.grid(row=7, column=0, sticky="w", pady=5)
        
        self.btn_export = Button(left_frame, text="Export ke Excel", command=self.export_ke_excel,
                              bg="#2196F3", fg="white", font=("Helvetica", 10),
                              width=15, relief=RAISED, bd=2)
        self.btn_export.grid(row=8, column=0, sticky="w", pady=5)
        
        # Right frame (table)
        right_frame = Frame(content_frame, bg="#f0f0f0")
        right_frame.pack(side=RIGHT, fill=BOTH, expand=True)
        
        # Search frame
        search_frame = Frame(right_frame, bg="#f0f0f0")
        search_frame.pack(fill=X, pady=(0, 10))
        
        Label(search_frame, text="Cari Produk:", bg="#f0f0f0").pack(side=LEFT, padx=(0, 5))
        Entry(search_frame, textvariable=self.search_var, width=25).pack(side=LEFT, fill=X, expand=True)
        
        # Treeview for product list
        table_frame = Frame(right_frame)
        table_frame.pack(fill=BOTH, expand=True)
        
        scroll_y = Scrollbar(table_frame)
        scroll_y.pack(side=RIGHT, fill=Y)
        
        scroll_x = Scrollbar(table_frame, orient='horizontal')
        scroll_x.pack(side=BOTTOM, fill=X)
        
        columns = ('ID', 'Nama', 'Stok', 'Harga')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                                yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        # Define column properties and headings
        self.tree.column('ID', width=50, anchor=CENTER)
        self.tree.column('Nama', width=200, anchor=W)
        self.tree.column('Stok', width=80, anchor=CENTER)
        self.tree.column('Harga', width=120, anchor=E)
        
        for col in columns:
            self.tree.heading(col, text=col, anchor=CENTER, 
                             command=lambda c=col: self.treeview_sort_column(self.tree, c, False))
        
        self.tree.pack(fill=BOTH, expand=True)
        
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        
        self.tree.bind("<ButtonRelease-1>", self.on_item_select)
        
        # Status bar
        self.status_var = StringVar()
        self.status_var.set("Aplikasi siap digunakan")
        status_bar = Label(self.root, textvariable=self.status_var, bd=1, relief=SUNKEN, anchor=W)
        status_bar.pack(side=BOTTOM, fill=X)
        
        # Apply style to treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
                      background="#f9f9f9",
                      foreground="black",
                      rowheight=25,
                      fieldbackground="#f9f9f9")
        style.configure("Treeview.Heading", 
                      font=('Helvetica', 10, 'bold'),
                      background="#ddd",
                      foreground="black")
        style.map('Treeview', background=[('selected', '#347083')])
        
        # Update treeview with data
        self.update_listbox()

if __name__ == "__main__":
    root = Tk()
    app = InventoryApp(root)
    root.mainloop()