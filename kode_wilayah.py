import pandas as pd

# Baca CSV lokal tanpa header
df = pd.read_csv("wilayah.csv", header=None, names=["kode", "nama"])
df["kode"] = df["kode"].astype(str)

# Tentukan level berdasarkan jumlah segmen
def level(kode):
    return len(kode.split("."))

df["level"] = df["kode"].apply(level)
df["kode_clean"] = df["kode"].str.replace(".", "", regex=False)

# Pisahkan per level
provinsi_df  = df[df["level"] == 1].copy()
kabupaten_df = df[df["level"] == 2].copy()
kecamatan_df = df[df["level"] == 3].copy()
desa_df      = df[df["level"] == 4].copy()

# Fungsi pemilihan dengan validasi dan opsi kembali
def pilih_wilayah(df, label, kode_col="kode_clean", nama_col="nama"):
    while True:
        print(f"\n📍 Daftar {label}:")
        for _, row in df.iterrows():
            print(f"{row[kode_col]} - {row[nama_col]}")
        kode = input(f"\nMasukkan kode {label.lower()} yang ingin dicari (atau ketik 'kembali'): ").strip()
        if kode.lower() == "kembali":
            return "kembali"
        if kode in df[kode_col].values:
            return kode
        print(f"❌ Kode {label.lower()} tidak ditemukan. Silakan coba lagi.")

# Alur interaktif
while True:
    kode_prov = pilih_wilayah(provinsi_df, "Provinsi")
    if kode_prov == "kembali":
        continue
    kabupaten_terpilih = kabupaten_df[kabupaten_df["kode_clean"].str.startswith(kode_prov)]

    while True:
        kode_kab = pilih_wilayah(kabupaten_terpilih, "Kabupaten")
        if kode_kab == "kembali":
            break
        kecamatan_terpilih = kecamatan_df[kecamatan_df["kode_clean"].str.startswith(kode_kab)]

        while True:
            kode_kec = pilih_wilayah(kecamatan_terpilih, "Kecamatan")
            if kode_kec == "kembali":
                break
            desa_terpilih = desa_df[desa_df["kode_clean"].str.startswith(kode_kec)]

            while True:
                print("\n📍 Daftar Desa di Kecamatan tersebut:")
                for _, row in desa_terpilih.iterrows():
                    print(f"{row['kode_clean']} - {row['nama']}")
                kode_desa = input("\nMasukkan kode desa yang ingin dicari (atau ketik 'kembali'): ").strip()
                if kode_desa.lower() == "kembali":
                    break
                desa_final = desa_terpilih[desa_terpilih["kode_clean"] == kode_desa]
                if not desa_final.empty:
                    print(f"\n✅ Desa ditemukan: {desa_final.iloc[0]['nama']} (Kode: {kode_desa})")
                    exit()
                else:
                    print("❌ Kode desa tidak ditemukan. Silakan coba lagi.")