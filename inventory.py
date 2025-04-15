import csv
import os
from openpyxl import Workbook

produk_list = []
id_counter = 1
csv_file = "produk.csv"

def load_data():
    global id_counter
    if os.path.exists(csv_file):
        with open(csv_file, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                produk = {
                    "id": int(row["id"]),
                    "nama": row["nama"],
                    "stok": int(row["stok"]),
                    "harga": int(row["harga"])
                }
                produk_list.append(produk)
            if produk_list:
                id_counter = max(p["id"] for p in produk_list) + 1

def simpan_data():
    with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
        fieldnames = ["id", "nama", "stok", "harga"]
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        for produk in produk_list:
            writer.writerow(produk)

def tambah_produk():
    global id_counter
    nama = input("Nama produk: ")
    stok = int(input("Stok produk: "))
    harga = int(input("Harga produk: "))
    
    produk = {
        "id": id_counter,
        "nama": nama,
        "stok": stok,
        "harga": harga
    }
    produk_list.append(produk)
    id_counter += 1
    simpan_data()
    print("✅ Produk berhasil ditambahkan.\n")

def lihat_produk():
    if not produk_list:
        print("❌ Belum ada produk.\n")
        return
    print("\n📦 Daftar Produk:")
    for produk in produk_list:
        print(f"ID: {produk['id']} | Nama: {produk['nama']} | Stok: {produk['stok']} | Harga: Rp{produk['harga']}")
    print()

def edit_produk():
    id_edit = int(input("Masukkan ID produk yang ingin diedit: "))
    for produk in produk_list:
        if produk["id"] == id_edit:
            produk["nama"] = input("Nama baru: ")
            produk["stok"] = int(input("Stok baru: "))
            produk["harga"] = int(input("Harga baru: "))
            simpan_data()
            print("✅ Produk berhasil diedit.\n")
            return
    print("❌ Produk dengan ID tersebut tidak ditemukan.\n")

def hapus_produk():
    id_hapus = int(input("Masukkan ID produk yang ingin dihapus: "))
    for produk in produk_list:
        if produk["id"] == id_hapus:
            produk_list.remove(produk)
            simpan_data()
            print("✅ Produk berhasil dihapus.\n")
            return
    print("❌ Produk dengan ID tersebut tidak ditemukan.\n")

def export_ke_excel():
    if not produk_list:
        print("❌ Tidak ada data untuk diekspor.\n")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Produk Warung"

    # Header
    ws.append(["ID", "Nama Produk", "Stok", "Harga"])

    # Data
    for produk in produk_list:
        ws.append([produk["id"], produk["nama"], produk["stok"], produk["harga"]])

    # Simpan file Excel
    wb.save("produk.xlsx")
    print("✅ Data berhasil diexport ke 'produk.xlsx'\n")


def menu():
    load_data()
    while True:
        print("=== INVENTORY PRODUK WARUNG ===")
        print("1. Tambah Produk")
        print("2. Lihat Produk")
        print("3. Edit Produk")
        print("4. Hapus Produk")
        print("5. Export ke Excel")
        print("6. Keluar")
        pilihan = input("Pilih menu (1-5): ")

        if pilihan == "1":
            tambah_produk()
        elif pilihan == "2":
            lihat_produk()
        elif pilihan == "3":
            edit_produk()
        elif pilihan == "4":
            hapus_produk()
        elif pilihan == "5":
            export_ke_excel()
        elif pilihan == "6":
            print("👋 Keluar dari program. Sampai jumpa!")
            break

        else:
            print("❌ Pilihan tidak valid.\n")

if __name__ == "__main__":
    menu()
