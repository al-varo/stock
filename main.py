import psycopg2
import pandas as pd

def nilai_stok_average_cost(conn, tanggal, lokasi_internal_id=8):
    cursor = conn.cursor()

    query = """
        WITH
        masuk AS (
            SELECT
                sm.product_id,
                SUM(sm.product_qty) AS qty_in,
                SUM(sm.product_qty * COALESCE(sm.price_unit, ip.value_float)) AS nilai_in
            FROM stock_move sm
            JOIN product_product pp ON sm.product_id = pp.id
            JOIN product_template pt ON pp.product_tmpl_id = pt.id
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
            ROUND(COALESCE(masuk.qty_in, 0) - COALESCE(keluar.qty_out, 0), 2) AS saldo_qty,
            ROUND(
                CASE
                    WHEN COALESCE(masuk.qty_in, 0) > 0 THEN
                        COALESCE(masuk.nilai_in, 0) / COALESCE(masuk.qty_in, 1)
                    ELSE
                        COALESCE(ip.value_float, 0)
                END, 2
            ) AS avg_cost,
            ROUND((
                COALESCE(masuk.qty_in, 0) - COALESCE(keluar.qty_out, 0)) *
                CASE
                    WHEN COALESCE(masuk.qty_in, 0) > 0 THEN
                        COALESCE(masuk.nilai_in, 0) / COALESCE(masuk.qty_in, 1)
                    ELSE
                        COALESCE(ip.value_float, 0)
                END, 2
            ) AS nilai_stok
        FROM masuk
        LEFT JOIN keluar ON masuk.product_id = keluar.product_id
        JOIN product_product pp ON masuk.product_id = pp.id
        JOIN product_template pt ON pp.product_tmpl_id = pt.id
        LEFT JOIN product_uom uom ON pt.uom_id = uom.id
        LEFT JOIN ir_property ip ON ip.name = 'standard_price'
            AND ip.res_id = ('product.template,' || pt.id)
        WHERE (COALESCE(masuk.qty_in, 0) - COALESCE(keluar.qty_out, 0)) > 0
        ORDER BY pt.name
    """

    cursor.execute(query, {"tanggal": tanggal, "lokasi_id": lokasi_internal_id})
    columns = [desc[0] for desc in cursor.description]
    rows = cursor.fetchall()
    cursor.close()

    df = pd.DataFrame(rows, columns=columns)
    total = df["nilai_stok"].sum()

    print(f"\n📦 Nilai Stok (Average Cost) per {tanggal}")
    print(df.to_string(index=False))
    print(f"\n💰 Total Nilai Stok: Rp {total:,.2f}")

    filename = f"stok_average_{tanggal}.xlsx"
    df.to_excel(filename, index=False)
    print(f"✅ Disimpan ke: {filename}")

if __name__ == "__main__":
    conn = psycopg2.connect(
        host="app.manzada.net",  # ganti dengan domain PostgreSQL Anda
        database="manzada",
        user="offline",
        password="ra#asia"
    )
    nilai_stok_average_cost(conn, "2024-12-31")  # sesuaikan tanggal
    conn.close()
