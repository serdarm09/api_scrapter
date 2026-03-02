import pandas as pd
import requests
import time

# ─────────────────────────────────────────────────────────────
# AYARLAR — Token bitince yenisini buraya yapıştırın
# F12 > Network > productSuggest > Request Headers > authorization
# ─────────────────────────────────────────────────────────────
BEARER_TOKEN = "xx495b79cc-a506-4525-9c9a-9aa10f9fac5a"
CLIENT_ID    = "1890ae2c-0adf-47bb-bc00-31ff5e57305b"

EXCEL_PATH   = "LİSTE_328.xlsx"
OUTPUT_PATH  = "LİSTE_328_TAM_SONUC.xlsx"

KAYIT_ARALIGI = 1000   # Her 1000 satırda bir ara kayıt
SLEEP         = 0.2    # İstek arası bekleme (saniye) — hızı/ban riskini dengeler

URL = "https://agcocorporationproduction16naig7ze.org.coveo.com/rest/organizations/agcocorporationproduction16naig7ze/commerce/v2/search/productSuggest"

HEADERS = {
    "accept": "*/*",
    "accept-language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
    "authorization": f"Bearer {BEARER_TOKEN}",
    "cache-control": "no-cache",
    "content-type": "application/json",
    "origin": "https://parts.agcocorp.com",
    "pragma": "no-cache",
    "priority": "u=1, i",
    "referer": "https://parts.agcocorp.com/",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36"
}


def fetch_product(kod):
    payload = {
        "trackingId": "canada",
        "query": str(kod),
        "clientId": CLIENT_ID,
        "context": {
            "user": {
                "userAgent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36"
            },
            "view": {
                "url": "https://parts.agcocorp.com/ca/en/search",
                "referrer": "https://parts.agcocorp.com/"
            },
            "capture": True,
            "cart": [],
            "source": ["@coveo/headless@3.42.1"],
            "region": "NA"
        },
        "language": "en",
        "country": "CA",
        "currency": "CAD"
    }
    try:
        r = requests.post(URL, headers=HEADERS, json=payload, timeout=10)
        if r.status_code == 200:
            data  = r.json()
            prods = data.get("products", [])
            if prods:
                prod  = prods[0]
                extra = prod.get("additionalFields", {})
                imgs  = prod.get("ec_images", []) or []
                thumbs = prod.get("ec_thumbnails", []) or []
                resim  = ", ".join(str(x) for x in imgs) or ", ".join(str(x) for x in thumbs)
                return {
                    "Durum"       : "Bulundu",
                    "Ürün Adı"    : prod.get("ec_name", ""),
                    "Açıklama"    : prod.get("ec_description", ""),
                    "OEM"         : extra.get("partnumber", ""),
                    "Part Number" : extra.get("ec_prd_manufacturer_partnumber", ""),
                    "Kategori"    : " > ".join(prod.get("ec_category", [])),
                    "Marka"       : prod.get("ec_brand", ""),
                    "Fiyat"       : prod.get("ec_price", ""),
                    "Resim Linki" : resim,
                    "Coveo ID"    : prod.get("ec_product_id", ""),
                }
            return {"Durum": "Ürün Bulunamadı"}
        elif r.status_code == 401:
            return {"Durum": "TOKEN_EXPIRED_401"}
        else:
            return {"Durum": f"HTTP {r.status_code}"}
    except Exception as e:
        return {"Durum": f"Hata: {e}"}


def main():
    print(f"Excel okunuyor: {EXCEL_PATH}")
    df = pd.read_excel(EXCEL_PATH)

    if "KOD" not in df.columns:
        print(f"'KOD' kolonu yok! Mevcut: {df.columns.tolist()}")
        return

    toplam  = len(df)
    results = []
    islenen = 0

    print(f"Toplam {toplam} satır işlenecek. Her {KAYIT_ARALIGI} kayıtta bir ara kayıt yapılacak.\n")

    for i, (_, row) in enumerate(df.iterrows()):
        kod = row["KOD"]
        row_dict = row.to_dict()

        if pd.isna(kod) or str(kod).strip() == "":
            row_dict["Durum"] = "Boş KOD"
            results.append(row_dict)
            continue

        print(f"[{i+1}/{toplam}] {kod}", end=" -> ")
        result = fetch_product(kod)
        print(result.get("Durum", ""))
        row_dict.update(result)
        results.append(row_dict)
        islenen += 1

        # Token bittiyse dur
        if result.get("Durum") == "TOKEN_EXPIRED_401":
            print("\n>>> TOKEN SÜRESİ DOLDU! BEARER_TOKEN değerini güncelleyip yeniden çalıştırın.")
            break

        time.sleep(SLEEP)

        # Her 1000 işlenen kayıtta ara kayıt
        if islenen % KAYIT_ARALIGI == 0:
            try:
                pd.DataFrame(results).to_excel(OUTPUT_PATH, index=False)
                print(f"\n=== ARA KAYIT: {islened} satır -> {OUTPUT_PATH} ===\n")
            except Exception as e:
                print(f"\nAra kayıt hatası (Excel açık olabilir): {e}\n")

    # Nihai kayıt
    out_df = pd.DataFrame(results)
    try:
        out_df.to_excel(OUTPUT_PATH, index=False)
        print(f"\nTamamlandı! Toplam {len(results)} satır -> {OUTPUT_PATH}")
    except Exception as e:
        print(f"Kayıt hatası: {e}")

    # Özet
    if "Durum" in out_df.columns:
        print("\nDurum Özeti:")
        print(out_df["Durum"].value_counts().to_string())


if __name__ == "__main__":
    main()


