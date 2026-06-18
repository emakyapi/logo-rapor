"""
Logo Veritabani -> JSON + Excel Rapor Uretici
=============================================
Bu script Logo DB'ye baglanir, verileri ceker ve
docs/data/ klasorune JSON + Excel olarak yazar.

Kurulum:
  pip install pyodbc pandas openpyxl python-dotenv

Calistirmak icin:
  python scripts/rapor_uret.py
"""

import os
import sys
import json
import pyodbc
import pandas as pd
from datetime import datetime
from pathlib import Path

# --- AYARLAR ---
# Bu bilgileri .env dosyasindan veya ortam degiskenlerinden aliyoruz
# .env dosyasinda bu satirlar olmali:
#   LOGO_SERVER=192.168.1.10
#   LOGO_DB=LOGODATA
#   LOGO_USER=sa
#   LOGO_PASS=sifreniz

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # GitHub Actions'ta .env gerekmez, Secret'lar zaten ortam degiskeni olarak gelir

LOGO_SERVER = os.environ.get("LOGO_SERVER", "localhost")
LOGO_DB     = os.environ.get("LOGO_DB",     "LOGODATA")
LOGO_USER   = os.environ.get("LOGO_USER",   "sa")
LOGO_PASS   = os.environ.get("LOGO_PASS",   "")

# Cikti klasoru (GitHub Pages bu klasoru yayinlar)
CIKTI_KLASORU = Path(__file__).parent.parent / "docs" / "data"
CIKTI_KLASORU.mkdir(parents=True, exist_ok=True)

# --- SQL SORGULARI ---
# Buraya istediginiz sorguları ekleyebilirsiniz
SORGULAR = {
    "musteriler": """
        SELECT TOP 200
            LOGICALREF    AS id,
            CODE          AS musteri_kodu,
            DEFINITION_   AS musteri_adi,
            ADDR1         AS adres,
            TELNRS1       AS telefon,
            EMAILADDR     AS email
        FROM LG_001_CLCARD
        WHERE ACTIVE = 0
          AND CARDTYPE = 3
        ORDER BY CODE
    """,

    "son_siparisler": """
        SELECT TOP 500
            ORF.LOGICALREF   AS siparis_id,
            ORF.FICHENO      AS siparis_no,
            CONVERT(VARCHAR, ORF.DATE_, 103) AS tarih,
            CL.CODE          AS musteri_kodu,
            CL.DEFINITION_   AS musteri_adi,
            ORF.NETTOTAL     AS toplam_tutar,
            ORF.REPORTNET    AS indirimli_tutar
        FROM LG_001_01_ORFICHE ORF
        LEFT JOIN LG_001_CLCARD CL ON CL.LOGICALREF = ORF.CLIENTREF
        WHERE ORF.TRCODE = 1
          AND ORF.DATE_ >= DATEADD(DAY, -30, GETDATE())
        ORDER BY ORF.DATE_ DESC
    """,

    "stok_ozeti": """
        SELECT TOP 300
            IT.CODE         AS stok_kodu,
            IT.NAME         AS stok_adi,
            IT.UNITSETREF   AS birim,
            ISNULL(ST.ONHAND, 0) AS mevcut_miktar,
            ISNULL(ST.ONORDER, 0) AS siparis_miktari
        FROM LG_001_ITEMS IT
        LEFT JOIN LG_001_01_STINVTOT ST ON ST.STOCKREF = IT.LOGICALREF
                                        AND ST.INVENNO = -1
        WHERE IT.ACTIVE = 0
        ORDER BY IT.CODE
    """
}


def baglan():
    """Logo veritabanina baglan"""
    print(f"[1/4] Veritabanina baglaniliyor: {LOGO_SERVER} / {LOGO_DB}")
    
    # Once SQL Server Authentication dene
    baglanti_string = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={LOGO_SERVER};"
        f"DATABASE={LOGO_DB};"
        f"UID={LOGO_USER};"
        f"PWD={LOGO_PASS};"
        f"TrustServerCertificate=yes;"
    )
    
    try:
        conn = pyodbc.connect(baglanti_string, timeout=10)
        print("    Baglanti basarili!")
        return conn
    except Exception as e:
        print(f"    HATA: {e}")
        print()
        print("    Lutfen .env dosyasindaki LOGO_SERVER, LOGO_DB,")
        print("    LOGO_USER, LOGO_PASS degerlerini kontrol edin.")
        sys.exit(1)


def sorgula_ve_kaydet(conn):
    """Tum sorgulari calistir, JSON ve Excel olarak kaydet"""
    ozet = {}
    
    for isim, sql in SORGULAR.items():
        print(f"[2/4] '{isim}' sorgusu calistiriliyor...")
        
        try:
            df = pd.read_sql(sql, conn)
            print(f"      {len(df)} kayit bulundu")
            
            # JSON olarak kaydet
            json_yolu = CIKTI_KLASORU / f"{isim}.json"
            df.to_json(json_yolu, orient="records", force_ascii=False, indent=2)
            
            # Excel olarak kaydet
            excel_yolu = CIKTI_KLASORU / f"{isim}.xlsx"
            df.to_excel(excel_yolu, index=False, engine="openpyxl")
            
            ozet[isim] = {
                "kayit_sayisi": len(df),
                "sutunlar": list(df.columns),
                "son_guncelleme": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
        except Exception as e:
            print(f"      UYARI: '{isim}' sorgusu basarisiz: {e}")
            ozet[isim] = {"hata": str(e)}
    
    return ozet


def meta_kaydet(ozet):
    """Site icin meta bilgi dosyasi olustur (hangi rapor ne zaman guncellendi)"""
    print("[3/4] Meta bilgi kaydediliyor...")
    
    meta = {
        "son_guncelleme": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tarih_turkce": datetime.now().strftime("%d %B %Y, %H:%M"),
        "raporlar": ozet
    }
    
    meta_yolu = CIKTI_KLASORU / "meta.json"
    with open(meta_yolu, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)
    
    print(f"      Kaydedildi: {meta_yolu}")


def main():
    print("=" * 50)
    print("Logo Rapor Uretici")
    print("=" * 50)
    
    conn = baglan()
    ozet = sorgula_ve_kaydet(conn)
    conn.close()
    meta_kaydet(ozet)
    
    print("[4/4] Tamamlandi!")
    print()
    print("Olusturulan dosyalar:")
    for dosya in sorted(CIKTI_KLASORU.iterdir()):
        boyut = dosya.stat().st_size
        print(f"  {dosya.name:30s}  {boyut:>8,} byte")
    
    print()
    print("Simdi 'git add . && git commit -m rapor && git push' calistirin.")


if __name__ == "__main__":
    main()
