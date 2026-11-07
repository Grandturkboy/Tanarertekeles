import pandas as pd

df = pd.read_csv(
    r"Tanárok névsora.csv",  # <== Ezt írd át, ha más a fájlnév
    header=None,
    names=["timestamp", "email", "class", "teachers"], 
    skiprows=1
)

df["class"] = df["class"].str.strip().str.lower()
df["teacher"] = df["teachers"].fillna("").str.split(";")

pairs = (
    df.explode("teacher")
    .assign(teacher=lambda d: d["teacher"].str.strip())
    .query("teacher != ''")[["class", "teacher"]]
    .drop_duplicates()
    .sort_values(["class", "teacher"])
)

pairs.to_csv(r"OsztályPerTanár.csv", index=False) # <== Ez lesz az elkészült fájl neve
