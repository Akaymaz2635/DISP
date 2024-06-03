import pandas as pd

# Veriyi bir DataFrame olarak oluşturma
data = {
    "Element": ["Datum D", "Datum D"],
    "Property": ["Parallelism", "Flatness"],
    "Nominal": [0, 0],
    "Tol -": [0, 0],
    "Tol +": [0.1, 0.01],
    "Min": [0.06, 0.04],
    "Max": [68, 45],
    "1T66L-1": ["64.000/65", "42/44"],
    "1T66L-13": ["62.000", "45/46"],
    "1T66L-14": ["68.000", "38/39"],
    "1T66L-15": ["0.064", "0.042/5"]
}

df = pd.DataFrame(data)

# Fonksiyonu tanımlama
def check_limits(row):
    nominal_min = row['Nominal'] - row['Tol -']
    nominal_max = row['Nominal'] + row['Tol +']
    if row['Min'] < nominal_min or row['Max'] > nominal_max:
        return f"{row['Element']} checks min {row['Min']} and max {row['Max']}."
    return ""

# Yeni kolonu ekleyip her satır için fonksiyonu uygulama
df['Reject'] = df.apply(check_limits, axis=1)

# DataFrame'i gösterme
print(df)
