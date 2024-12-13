import csv
import os
from datetime import date, datetime

import pandas as pd

# py -m pip install openpyxl
# py -m pip install xlrd

intro_calls_file = pd.read_excel(f"{os.path.dirname(os.path.abspath(__file__))}/../Intro Calls.xlsx")
companies = [company for company in intro_calls_file["Company Name"]]
if(not os.path.exists(f"{os.path.dirname(os.path.abspath(__file__))}/../DUPLICATES.csv")):
    with open(f"{os.path.dirname(os.path.abspath(__file__))}/../DUPLICATES.csv", "a+", newline="") as dup_file:
        writer = csv.writer(dup_file)
        writer.writerow(["Date of Intro Call", "Company Name", "Found in Batch", "Date of Duplicate Recorded"])
dup_companies = pd.read_csv(f"{os.path.dirname(os.path.abspath(__file__))}/../DUPLICATES.csv")["Company Name"].to_list()
duplicate_file = open(f"{os.path.dirname(os.path.abspath(__file__))}/../DUPLICATES.csv", "a+", newline="")
writer = csv.writer(duplicate_file)

for filename in os.listdir(f"{os.path.dirname(os.path.abspath(__file__))}/.."):
        if(filename.endswith(".csv") and "DUPLICATES" not in filename):
            batch = pd.read_csv(f"{os.path.dirname(os.path.abspath(__file__))}/../{filename}")
            batch_copy = batch.copy(deep=True)
            for organization in batch["Organization"]:
                if organization in companies:
                    print(f"removing {organization} from {filename.strip(".csv")}")
                    batch_copy = batch_copy[batch_copy['Organization'] != organization]
                    if organization not in dup_companies:
                        writer.writerow([intro_calls_file["Date of Intro Call"][companies.index(organization)].date(), organization, filename.strip(".csv"), date.today()])
                        dup_companies.append(organization)
            batch_copy.to_csv(f"{os.path.dirname(os.path.abspath(__file__))}/../{filename}", index=False)

