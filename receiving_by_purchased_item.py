"""
Track purchases of selected items and create spreadsheet
"""

import json
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog


def get_file_path():
    root = tk.Tk()
    root.title("Select Receiving by Purchased Item Report")
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path


def get_table(file_path):
    df = pd.read_csv(
        file_path,
        skiprows=3,
        usecols=[
            "ItemName",
            "LocationName",
            "TransactionNumber",
            "VendorName",
            "Textbox11",
            "TransactionDate",
            "PurchaseUnit",
            "Quantity",
            "AmountEach",
            "ExtPrice2",
        ],
    )
    try:
        filter = df.Quantity.str.match(r"\((.+)\)")
        df = df[~filter]
    except:
        pass

    item_list = [df.ItemName.unique()]
    item_list = [item for sublist in item_list for item in sublist]
    item_list.sort()

    with open("/usr/local/share/UofM.json") as file:
        uofm = json.load(file)
    units = pd.DataFrame(uofm)
    df = df.merge(units, left_on="PurchaseUnit", right_on="Name", how="left")
    df.rename(columns={"Textbox11": "VendorNumber"}, inplace=True)
    for column in df.columns:
        # check if column contains parentheses and convert to negative float
        if df[column].dtype == "object":
            try:
                filter = df[column].str.match(r"\((.+)\)")
                df.loc[filter, column] = df.loc[filter, column].str.replace(r"\(|\)", "").astype(float) * -1
            except:
                pass
    try:
        df["BaseQty"] = df["BaseQty"].str.replace(",", "").astype(float)
    except:
        df["BaseQty"] = df["BaseQty"].astype(float)
    df["Quantity"] = df["Quantity"].astype(float)
    try:
        df["AmountEach"] = df["AmountEach"].str.replace(",", "").astype(float)
    except:
        df["AmountEach"] = df["AmountEach"].astype(float)
    try:
        df["ExtPrice2"] = df["ExtPrice2"].astype(str).str.replace(",", "").astype(float)
    except:
        df["ExtPrice2"] = df["ExtPrice2"].astype(float)
    # rename df["ExtPrice2"] to df["ExtPrice"] to match other reports
    df.rename(columns={"ExtPrice2": "ExtCost"}, inplace=True)
    # df.loc["Totals"] = df.sum(numeric_only=True)
    sorted_units = (
        df.groupby(["Name"]).mean(numeric_only=True).sort_values(by=["Quantity"], ascending=False).reset_index()
    )
    df_sorted = pd.DataFrame()
    for item in item_list:
        df_temp = df[df.ItemName == item]
        sorted_units = (
            df_temp.groupby(["Name"])
            .mean(numeric_only=True)
            .sort_values(by=["Quantity"], ascending=False)
            .reset_index()
        )
        report_unit = df_temp.iloc[0]["Name"]
        base_factor = df_temp.iloc[0]["BaseQty"]
        df_temp["reportUnit"] = report_unit
        df_temp["base_factor"] = base_factor
        df_temp["totalQuantity"] = df["Quantity"] * df["BaseQty"] / base_factor
        df_temp["unit"] = report_unit
        df_sorted = pd.concat([df_sorted, df_temp], ignore_index=True)

    return df_sorted


def make_pivot(table):
    vendor = pd.pivot_table(
        table,
        values=["totalQuantity", "ExtCost"],
        index=["ItemName", "VendorName", "unit"],
        aggfunc=np.sum,
    )
    vendor = vendor.reset_index().sort_values(["ItemName", "VendorName"]).set_index("VendorName")
    vendor.loc["Totals"] = vendor.sum(numeric_only=True)
    vendor["CostPerUnit"] = vendor["ExtCost"] / vendor["totalQuantity"]

    restaurant = pd.pivot_table(
        table,
        values=["totalQuantity", "ExtCost"],
        index=["ItemName", "LocationName", "unit"],
        aggfunc=np.sum,
    )
    restaurant = restaurant.reset_index().sort_values(["ItemName", "LocationName"]).set_index("LocationName")
    restaurant.loc["Totals"] = restaurant.sum(numeric_only=True)
    restaurant["CostPerUnit"] = restaurant["ExtCost"] / restaurant["totalQuantity"]
    restaurant.style.format(
        {
            "ExtCost": "${:,.2f}",
            "totalQuantity": "{:,.0f}",
            "CostPerUnit": "${:,.2f}",
        }
    )
    return [vendor, restaurant]


def save_file(table):
    filename = filedialog.asksaveasfilename(filetypes=(("XLSX Files", "*.xlsx"), ("All Files", "*.*")))
    with pd.ExcelWriter(filename) as writer:
        vendor, restaurant = make_pivot(df_table)
        vendor.to_excel(writer, sheet_name="Vendor")
        restaurant.to_excel(writer, sheet_name="Restaurant")
        table.to_excel(writer, sheet_name="Detail", index=False)


if __name__ == "__main__":
    file = get_file_path()
    df_table = get_table(file)
    save_file(df_table)
