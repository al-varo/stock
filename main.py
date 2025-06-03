import psycopg2
import pandas as pd
from openpyxl import Workbook
from datetime import datetime

import os

# Konfigurasi koneksi PostgreSQL
DB_CONFIG = {
    "host": "app.manzada.net",
    "port": 5432,
    "dbname": "manzada",
    "user": "offline",
    "password": "ra#asia"
}

# Fungsi ambil harga average dari ir_property
def ambil_harga_average(conn):
    query = """
        SELECT
            REPLACE(SPLIT_PART(name, ',', 1), 'product.template,', '')::integer AS product_template_id,
            value_float
        FROM ir_property
        WHERE name = 'standard_price'
          AND res_id IS NOT NULL
    """
    with conn.cursor() as cursor:
        cursor.execute(query)
        hasil = cursor.fetchall()
        return {row[0]: float(row[1]) for row in hasil}  # Convert ke float

# Fungsi ambil ID lokasi internal
def ambil_lokasi_internal(conn):
    query = "SELECT id FROM stock_location WHERE usage = 'internal'"
    with conn.cursor() as cursor:
        cursor.execute(query)
        return [row[0] for row in cursor.fetchall()]

# Fungsi menghitung qty per product template pada tanggal tertentu
def ambil_qty_stok_per_tanggal(conn, tanggal, lokasi_ids):
    query = """
        SELECT
            pt.id AS product_template_id,
            SUM(CASE WHEN sm.location_dest_id = ANY(%(lokasi_ids)s) THEN sm.product_qty
                     WHEN sm.location_id = ANY(%(lokasi_ids)s) THEN -sm.product_qty
                     ELSE 0 END) AS qty
        FROM stock_move sm
        JOIN product_product pp ON sm.product_id = pp.id
        JOIN product_template pt ON pp.product_tmpl_id = pt.id
        WHERE sm.state = 'done'
          AND sm.date <= %(tanggal)s
        GROUP BY pt.id
        HAVING SUM(CASE WHEN sm.location_dest_id = ANY(%(lokasi_ids)s) THEN sm.product_qty
                        WHEN sm.location_id = ANY(%(lokasi_ids)s) THEN -sm.product_qty
                        ELSE 0 END) > 0
    """
    with conn.cursor() as cursor:
        cursor.execute(query, {"tanggal": tanggal, "lokasi_ids": lokasi_ids})
        return {row[0]: float(row[1]) for row in cursor.fetchall()}

# Fungsi mengambil nama produk
def ambil_nama_produk(conn):
    query = "SELECT id, name FROM product_template"
    with conn.cursor() as cursor:
        cursor.execute(query)
        return {row[0]: row[1] for row in cursor.fetchall()}

# Fungsi menghitung nilai stok
def hitung_nilai_stok(conn, tanggal):
    harga_avg = ambil_harga_average(conn)
    lokasi_internal = ambil_lokasi_internal(conn)
    qty_stok = ambil_qty_stok_per_tanggal(conn, tanggal, lokasi_internal)
    nama_produk = ambil_nama_produk(conn)

    hasil = []
    total_nilai = 0

    for product_template_id, qty in qty_stok.items():
        harga = harga_avg.get(product_template_id, 0.0)
        nilai = round(qty * harga, 2)
        total_nilai += nilai

        hasil.append({
            "Product": nama_produk.get(product_template_id, f"ID {product_template_id}"),
            "Qty": round(qty, 2),
            "Average Price": round(harga, 2),
            "Total Value": nilai
        })

    hasil.append({
        "Product": "TOTAL",
        "Qty": "",
        "Average Price": "",
        "Total Value": round(total_nilai, 2)
    })

    return hasil

# Fungsi simpan ke Excel
def simpan_ke_excel(data, tanggal):
    df = pd.DataFrame(data)
    folder = "/app" if os.path.exists("/app") else "."
    filename = f"{folder}/stok_averaging_{tanggal}.xlsx"
    df.to_excel(filename, index=False)
    print(f"File Excel disimpan: {filename}")

# Fungsi utama
if __name__ == "__main__":
    tanggal = "2024-12-31"  # Ubah sesuai kebutuhan

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        data = hitung_nilai_stok(conn, tanggal)
        simpan_ke_excel(data, tanggal)
    finally:
        conn.close()
