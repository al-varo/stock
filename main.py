import psycopg2
import pandas as pd
from datetime import datetime

def ambil_harga_average(conn):
    cursor = conn.cursor()
    query = """
        SELECT
            REPLACE(prop.res_id, 'product.template,', '')::int AS product_tmpl_id,
            prop.value_float AS standard_price
        FROM
            ir_property prop
        WHERE
            prop.name = 'standard_price'
            AND prop.res_id IS NOT NULL
            AND prop.value_float IS NOT NULL;
    """
    cursor.execute(query)
    return {row[0]: row[1] for row in cursor.fetchall()}

def ambil_lokasi_internal(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM stock_location WHERE usage = 'internal'")
    return [row[0] for row in cursor.fetchall()]

def hitung_nilai_stok(conn, tanggal_snapshot):
    lokasi_internal_ids = ambil_lokasi_internal(conn)
    harga_avg = ambil_harga_average(conn)

    cursor = conn.cursor()
    lokasi_tuple = tuple(lokasi_internal_ids)

    query = f"""
        SELECT
            sm.product_id,
            pt.id AS product_tmpl_id,
            pp.default_code,
            pt.name,
            uom.name AS uom_name,
            SUM(
                CASE
                    WHEN sm.location_dest_id IN %s THEN sm.product_qty
                    WHEN sm.location_id IN %s THEN -sm.product_qty
                    ELSE 0
                END
            ) AS qty
        FROM stock_move sm
        JOIN product_product pp ON sm.product_id = pp.id
        JOIN product_template pt ON pp.product_tmpl_id = pt.id
        LEFT JOIN product_uom uom ON pt.uom_id = uom.id
        WHERE sm.state = 'done'
          AND sm.date <= %s
          AND (
              sm.location_id IN %s OR
              sm.location_dest_id IN %s
          )
        GROUP BY sm.product_id, pt.id, pp.default_code, pt.name, uom.name
        HAVING SUM(
            CASE
                WHEN sm.location_dest_id IN %s THEN sm.product_qty
                WHEN sm.location_id IN %s THEN -sm.product_qty
                ELSE 0
            END
        ) > 0
    """

    params = (
        lokasi_tuple,
        lokasi_tuple,
        tanggal_snapshot,
        lokasi_tuple,
        lokasi_tuple,
        lokasi_tuple,
        lokasi_tuple
    )

    cursor.execute(query, params)
    rows = cursor.fetchall()

    hasil = []
    for row in rows:
        product_id, tmpl_id, kode, nama, uom, qty = row
        harga = harga_avg.get(tmpl_id, 0.0)
        nilai = round(qty * harga, 2)
        hasil.append({
            "Kode Produk": kode,
            "Nama Produk": nama,
            "Qty": round(qty, 2),
            "Satuan": uom,
            "Harga Rata-rata": round(harga, 2),
            "Nilai Stok": nilai
        })

    return hasil

def simpan_ke_excel(data, tanggal_snapshot):
    df = pd.DataFrame(data)
    filename = f"stok_average_{tanggal_snapshot}.xlsx"
    df.to_excel(filename, index=False)
    print(f"[✔] Disimpan ke file: {filename}")
    return filename

if __name__ == "__main__":
    # KONFIGURASI DATABASE
    conn = psycopg2.connect(
        host="app.manzada.net",
        port="5432",
        dbname="manzada",
        user="offline",
        password="ra#asia"
    )

    tanggal = "2024-12-31"
    data = hitung_nilai_stok(conn, tanggal)
    simpan_ke_excel(data, tanggal)
    conn.close()
