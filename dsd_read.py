import pandas as pd
from dateutil import parser

def load_dsd(dsd_file_path):
    df = pd.read_excel(dsd_file_path, engine="openpyxl")

    def find_rate_value():
        for index, row in df.iterrows():
            for col_index, (col_name, value) in enumerate(row.items()):
                if isinstance(value, str) and "rate" in value.lower():
                    return df.iloc[index + 1, col_index] if index + 1 < len(df) else None
        return None

    def total_impressions():
        for index, row in df.iterrows():
            for col_index, (col_name, value) in enumerate(row.items()):
                if isinstance(value, str) and "impressions" in value.lower():
                    return df.iloc[index + 1, col_index] if index + 1 < len(df) else None
        return None

    def start_date():
        for index, row in df.iterrows():
            for col_index, (col_name, value) in enumerate(row.items()):
                if isinstance(value, str) and "start date" in value.lower():
                    return df.iloc[index + 1, col_index] if index + 1 < len(df) else None
        return None

    def end_date():
        for index, row in df.iterrows():
            for col_index, (col_name, value) in enumerate(row.items()):
                if isinstance(value, str) and "end date" in value.lower():
                    return df.iloc[index + 1, col_index] if index + 1 < len(df) else None
        return None

    def site():
        for index, row in df.iterrows():
            for col_index, (col_name, value) in enumerate(row.items()):
                if isinstance(value, str) and "site" in value.lower():
                    return df.iloc[index + 1, col_index] if index + 1 < len(df) else None
        return None

    def geo():
        for index, row in df.iterrows():
            for col_index, (col_name, value) in enumerate(row.items()):
                if isinstance(value, str) and "geo" in value.lower():
                    return df.iloc[index + 1, col_index] if index + 1 < len(df) else None
        return None

    def Fcap():
        for index, row in df.iterrows():
            for col_index, (col_name, value) in enumerate(row.items()):
                if isinstance(value, str) and "fcap" in value.lower():
                    return df.iloc[index + 1, col_index] if index + 1 < len(df) else None
        return None

    return {
        "rate": find_rate_value(),
        "total_impressions": total_impressions(),
        "start_date": start_date(),
        "end_date": end_date(),
        "site": site(),
        "geo": geo(),
        "fcap": Fcap()
    }
