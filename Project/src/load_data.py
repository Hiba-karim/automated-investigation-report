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
 
 def load_diff_appro(path:str) -> pd.DataFrame : 
    """
    Sheet 'Diff_appro' - Supply chain difficulty levels.
    Transported structure: rows = severity categories, columns = survey waves

    We pivot it so each row is a survey wave
    """
    df_row= pd.read_excel(
        path,
        sheet_name= "Diff_appro", 
        header= None , 
    )
    #Row 0 : survey wave labels (columns 1 onward)
    #Column 0: category labels 
    #We skip the formula row (row 7 = "sum of positif feedback")
    categories = df_row.iloc[1:6, 0].tolist()
    survey_labels = df_raw.iloc[0, 1:].tolist() #survey wave names
    values= df_row.iloc[1:6, 1:].values

    #Building a clean DataFrame: one row per survey wave, one col per category
    df=pd.DataFrame(values.T , columns= categories)
    df.insert(0, "enquete",survey_labels)

    #Extract a sortable date from the survey label (We take the year+ month hint)
    #e.g. 'Enquête PME (15 mai - 9 juin 2026)' -> approximate date=juin 2026
    def parse_survey_date(label:str) -> pd.Timestamp:
        import re
        months_fr = {
            "jan": 1, "fév":2,"mar": 3, "avr": 4, "mai": 5, "juin": 6,
            "juil": 7, "août": 8, "sep": 9, "oct": 10, "nov": 11, "déc": 12,
        }
        #We grap the last months + year in the label
        m=re.findall(r'(\d+)\s+([a-zéû\.]+)[,\s]+(\d{4})', label.lower())
        if m:
            day, month_str, year = m[-1]
            month_key = month_str[:3].replace(".","")
            month_num = months_fr.get(month_key, 6)
            return pd.Timestamp(int(year), month_num, int(day))
        #fallback: we try to find any 4-digit year 
        years = re.findall(r'\d{4}', label)
        return pd.Timestamp(int(years[-1], 6, 1)) if years else pd.NaT
    
    df["date"] = df["enquete"].apply(parse_survey_date)
    df = df.sort_values("date").reset_index(drop=True)

    
