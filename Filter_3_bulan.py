import pandas as pd

# Path file terbaru
file_dec = "Balancepos20241230-2.xlsx"
file_jan = "Balancepos20250131.xlsx"
file_feb = "Balancepos20250228.xlsx"

# Membaca file Excel
df_dec = pd.read_excel(file_dec)
df_jan = pd.read_excel(file_jan)
df_feb = pd.read_excel(file_feb)

# Bersihkan nama kolom dari spasi berlebih
for df in [df_dec, df_jan, df_feb]:
    df.rename(columns=lambda x: x.strip(), inplace=True)

# Hitung total Scripless
for df in [df_dec, df_jan, df_feb]:
    df['Scripless'] = df['Total Lokal'] + df['Total Foreign']

# Gabungkan data berdasarkan kode saham
df_merge = df_dec[['Code', 'Local ID', 'Scripless']].merge(
    df_jan[['Code', 'Local ID', 'Scripless']], on='Code', suffixes=('_dec', '_jan')
).merge(
    df_feb[['Code', 'Local ID', 'Scripless']], on='Code'
).rename(columns={'Local ID': 'Local ID_feb', 'Scripless': 'Scripless_feb'})

# Fungsi untuk menghitung kenaikan & penurunan
def hitung_perubahan(df, bulan_awal, bulan_akhir, label):
    df[f'Nominal_Kenaikan_Local_ID_{label}'] = df[f'Local ID_{bulan_akhir}'] - df[f'Local ID_{bulan_awal}']
    df[f'Nominal_Penurunan_Local_ID_{label}'] = df[f'Local ID_{bulan_awal}'] - df[f'Local ID_{bulan_akhir}']

    # Hindari pembagian dengan nol
    df[f'Persentase_Kenaikan_Local_ID_{label}'] = df.apply(
        lambda row: ((row[f'Local ID_{bulan_akhir}'] / row[f'Local ID_{bulan_awal}']) - 1) * 100
        if row[f'Local ID_{bulan_awal}'] != 0 else None, axis=1
    )
    
    df[f'Persentase_Penurunan_Local_ID_{label}'] = df.apply(
        lambda row: ((row[f'Local ID_{bulan_akhir}'] / row[f'Local ID_{bulan_awal}']) - 1) * 100
        if row[f'Local ID_{bulan_awal}'] != 0 else None, axis=1
    )

# Hitung perubahan untuk masing-masing periode
hitung_perubahan(df_merge, 'dec', 'jan', 'Dec → Jan')
hitung_perubahan(df_merge, 'jan', 'feb', 'Jan → Feb')
hitung_perubahan(df_merge, 'dec', 'feb', 'Dec → Feb')

# Fungsi untuk menampilkan top 50 saham
def tampilkan_top_50(df, kolom_persen_up, kolom_nominal_up, kolom_persen_down, kolom_nominal_down, periode):
    print(f"\n📈 Top 50 Saham dengan kenaikan Local ID terbesar ({periode}):")
    print(df.nlargest(50, kolom_persen_up)[['Code', kolom_persen_up, kolom_nominal_up]])

    print(f"\n📉 Top 50 Saham dengan penurunan Local ID terbesar ({periode}):")
    print(df.nsmallest(50, kolom_persen_down)[['Code', kolom_persen_down, kolom_nominal_down]])

# Tampilkan hasil
tampilkan_top_50(
    df_merge, 
    'Persentase_Kenaikan_Local_ID_Dec → Jan', 'Nominal_Kenaikan_Local_ID_Dec → Jan',
    'Persentase_Penurunan_Local_ID_Dec → Jan', 'Nominal_Penurunan_Local_ID_Dec → Jan',
    'Dec → Jan'
)

tampilkan_top_50(
    df_merge, 
    'Persentase_Kenaikan_Local_ID_Jan → Feb', 'Nominal_Kenaikan_Local_ID_Jan → Feb',
    'Persentase_Penurunan_Local_ID_Jan → Feb', 'Nominal_Penurunan_Local_ID_Jan → Feb',
    'Jan → Feb'
)

tampilkan_top_50(
    df_merge, 
    'Persentase_Kenaikan_Local_ID_Dec → Feb', 'Nominal_Kenaikan_Local_ID_Dec → Feb',
    'Persentase_Penurunan_Local_ID_Dec → Feb', 'Nominal_Penurunan_Local_ID_Dec → Feb',
    'Dec → Feb'
)
