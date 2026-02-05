import pandas as pd
import numpy as np


# used this to pull the sheet names in the below list (we need to remove the first sheet and the last two sheets so we only loop on the products sheets)
# pd.read_excel("Cost Menu Brazilio Bar & Bakery (1).xlsx",sheet_name=None).keys()

sheets = ['Extra Bar', 'Matcha', 'Turkish Coffee', 'Around The World', 'Hot Chocolate', 'Espresso Base', 'Juicy Coffee', 'ice Espresso', 'Frappe Cream', 'milk shake ', 'Ice Blended Coffee', 'mojito', 'Juice', 'Water & Soft Srink', 'Bakery & Desrets']


# **** note before you run the code on the raw data: the sheet "Ice Blended Coffee" has a few empty columns before the actual data, make sure to delete them first

dfs_list = []

for sheet in sheets:
    df = pd.read_excel("Cost Menu Brazilio Bar & Bakery (1).xlsx",sheet_name=sheet,header=None)

    # split the sheet into two data frames to combine them in one and renaming the columns
    df1 = df[[0,1,2,3,4,5]]
    df1.columns = ["Ingrediants", "Amount", "Unit", "G.unit","P/G.unit","C/amount"]
    df2 = df[[8,9,10,11,12,13]]
    df2.columns = ["Ingrediants", "Amount", "Unit", "G.unit","P/G.unit","C/amount"]
    # for df2 it is mandatory to add an empty row at the start of it so the indexing in the below block of code works correctly
    df2.index = df2.index+1
    df2.loc[0] = np.nan
    df2.sort_index(inplace=True)
    # combine the two dataframes
    combined = pd.concat([df1,df2],ignore_index=True)


    # this section is to fill empty product tables with any name so the below indexing codes work
    target_indices = combined[combined["Ingrediants"] == "Ingrediants"].index
    indices_to_fill = [i-1 for i in target_indices if i > 0 and combined.iloc[i-1].isnull().all()]
    combined.loc[indices_to_fill, "Ingrediants"] = "empty_product"

    # first we drop the empty row between product tables (we do this first because some sheets have them and some don't) (the regex to remove spaces in any cell in the sheet)
    combined = combined.replace(r'^\s*$', np.nan, regex=True).dropna(how='all').reset_index(drop=True)

    # extract the product name on a separate column
    combined.loc[combined.index % 14 == 0,"product"] = combined["Ingrediants"]


    # fill down the product name
    combined["product"] = combined["product"].ffill()
    # drop the actual product name row (the merged cell)
    combined = combined.drop(combined.index[0::14]).reset_index(drop=True)
    # drop the header row for each product
    combined = combined.drop(combined.index[0::13]).reset_index(drop=True)
    # drop the grand total row
    combined = combined.drop(combined.index[11::12]).reset_index(drop=True)
    # drop null rows that doesn't contain ingredients
    combined = combined.dropna(subset=["Ingrediants", "Amount", "Unit", "G.unit","P/G.unit"]).reset_index(drop=True)
    # re-arrange the dataframe
    combined = combined[["product","Ingrediants", "Amount", "Unit", "G.unit","P/G.unit","C/amount"]]

    # append every dataframe to a list
    dfs_list.append(combined)

# concatenate that list of dataframes
all_dfs = pd.concat(dfs_list,ignore_index=True)
# drop the empty products we added before and reset index
all_dfs.drop(index=all_dfs[all_dfs["product"] == "empty_product"].index,inplace=True)
all_dfs.reset_index(drop=True,inplace=True)


# i don't think you will need an output if you are going to integrate the script with Power BI, I am not sure how it works, so, good luck.
# all_dfs.to_csv("output.csv",index=False)