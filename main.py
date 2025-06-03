import psycopg2
import pandas as pd
from datetime import datetime

def nilai_stok_average_cost(conn, tanggal: str):
    lokasi_internal_id = 8  # ganti sesuai lokasi internal Anda
    query = """
    WITH
    semua_produk AS (
        SELECT DISTINCT product_id FROM stock_move
        WHERE state = 'done'
            AND date <= %(tanggal)s
            AND (
                location_id = %(lokasi_id)s OR
                location_dest_id = %(lokasi_id)s
            )
    ),
    masuk AS (
        SELECT
            sm.product_id,
            SUM(sm.product_qty) AS qty_in,
            SUM(sm.product_qty * COALESCE(sm.price_unit, ip.value_float)) AS nilai_in
        FROM stock_move sm
        LEFT JOIN product_product pp ON sm.product_id = pp.id
        LEFT JOIN product_template pt ON pp.product_tmpl_id = pt.id
        LEFT JOIN ir_property ip ON ip.name = 'standard_price'
            AND ip.res_id = ('product.template,' || pt.id)
        WHERE sm.state = 'done'
            AND sm.location_dest_id = %(lokasi_id)s
            AND sm.location_id != %(lokasi_id)s
            AND sm.date <= %(tanggal)s
        GROUP BY sm.product_id
    ),
    keluar AS (
        SELECT
            sm.product_id,
            SUM(sm.product_qty) AS qty_out
        FROM stock_move sm
        WHERE sm.state = 'done'
            AND sm.location_id = %(lokasi_id)s
            AND sm.location_dest_id != %(lokasi_id)s
            AND sm.date <= %(tanggal)s
        GROUP BY sm.product_id
    )

    SELECT
        pp.default_code AS kode,
        pt.name AS nama_produk,
        uom.name AS satuan,
        ROUND((COALESCE(masuk.qty_in, 0) - COALESCE(keluar.qty_out, 0))::numeric, 2) AS saldo_qty,
        ROUND((
            CASE
                WHEN COALESCE(masuk.qty_in, 0) > 0 THEN
                    COALESCE(masuk.nilai_in, 0) / COALESCE(masuk.qty_in, 1)
                ELSE
                    COALESCE(ip.value_float, 0)
            END
        )::numeric, 2) AS avg_cost,
        ROUND((
            (COALESCE(masuk.qty_in, 0) - COALESCE(keluar.qty_out, 0)) *
            CASE
                WHEN COALESCE(masuk.qty_in, 0) > 0 THEN
                    COALESCE(masuk.nilai_in, 0) / COALESCE(masuk.qty_in, 1)
                ELSE
                    COALESCE(ip.value_float, 0)
            END
        )::numeric, 2) AS nilai_stok
    FROM semua_produk sp
    LEFT JOIN masuk ON sp.product_id = masuk.product_id
    LEFT JOIN keluar ON sp.product_id = keluar.product_id
    LEFT JOIN product_product pp ON sp.product_id = pp.id
    LEFT JOIN product_template pt ON pp.product_tmpl_id = pt.id
    LEFT JOIN product_uom uom ON pt.uom_id = uom.id
    LEFT JOIN ir_property ip ON ip.name = 'standard_price'
        AND ip.res_id = ('product.template,' || pt.id)
    WHERE (COALESCE(masuk.qty_in, 0) - COALESCE(keluar.qty_out, 0)) > 0
    ORDER BY pt.name
    """

    with conn.cursor() as cursor:
        cursor.execute(query, {"tanggal": tanggal, "lokasi_id": lokasi_internal_id})
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        df = pd.DataFrame(rows, columns=columns)

    filename = f"stok_averaging_{tanggal}.xlsx"
    df.to_excel(filename, index=False)
    print(f"✅ File berhasil disimpan: {filename}")

if __name__ == "__main__":
    conn = psycopg2.connect(
        host="app.manzada.net",  # ganti dengan domain/server PostgreSQL
        port="5432",
        dbname="manzada",
        user="offline",
        password="ra#asia"
    )

    tanggal = "2024-12-31"  # ganti sesuai kebutuhan
    nilai_stok_average_cost(conn, tanggal)
    conn.close()
