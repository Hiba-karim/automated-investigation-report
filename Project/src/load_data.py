#Data Loading & Processing
 
import pandas as pd
import numpy as np
 
  
DATA_PATH = r"data\donnees.xlsx"   # adjust path as needed
 
 
# Raw loading 
 
def load_ca_eff(path: str) -> pd.DataFrame:
    """
    Sheet 'CA Eff N' — Revenue & headcount evolution (opinion balance, %).
    Layout: row 0-1 are headers, column A is the date, B is CA, D is Effectifs.
    The 'Moyenne' columns contain Excel formulas → we recompute them ourselves.
    """
    df = pd.read_excel(
        path,
        sheet_name="CA Eff N",
        header=None,
        skiprows=2,        # skip the two header rows
        usecols=[0, 1, 3], # date | CA | Effectifs  (columns C and E are formulas)
    )
    df.columns = ["date", "ca", "effectifs"]
 
    # Drop rows without a date or with no data at all
    df = df.dropna(subset=["date"])
    df = df[df["ca"].notna() | df["effectifs"].notna()]
 
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)
 
    # Recompute historical averages (2000–2024) directly from the data
    mask_hist = (df["date"].dt.year >= 2000) & (df["date"].dt.year <= 2024)
    df["ca_mean_hist"]         = df.loc[mask_hist, "ca"].mean()
    df["effectifs_mean_hist"]  = df.loc[mask_hist, "effectifs"].mean()
 
    return df


def load_carnets(path: str) -> pd.DataFrame:
    """
    Sheet 'Carnets de commande' — Order book levels.
    Columns: date | carnet_passe (last 6 months) | carnet_futur (next 6 months).
    """
    df = pd.read_excel(
        path,
        sheet_name="Carnets de commande",
        header=None,
        skiprows=2,
        usecols=[0, 1, 3],
    )
    df.columns = ["date", "carnet_passe", "carnet_futur"]
 
    df = df.dropna(subset=["date"])
    df = df[df["carnet_passe"].notna() | df["carnet_futur"].notna()]
 
    df["date"] = pd.to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)
 
    mask_hist = (df["date"].dt.year >= 2000) & (df["date"].dt.year <= 2024)
    df["carnet_passe_mean_hist"] = df.loc[mask_hist, "carnet_passe"].mean()
    df["carnet_futur_mean_hist"] = df.loc[mask_hist, "carnet_futur"].mean()
 
    return df
 