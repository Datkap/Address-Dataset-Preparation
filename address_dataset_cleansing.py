# import modules and libraries
import pandas as pd
from translator import removeAccents
import openpyxl
print("Libraries loaded.")

# read full database (db) file
full_db = pd.read_csv("sample_data.csv")
print("Full DB loaded.")

# remove useless columns from db
db = full_db[[
    "Województwo",
    "Powiat",
    "Gmina",
    "Miejscowość (GUS)",
    "Ulica (cecha)",
    "Ulica (nazwa)",
    "Kod pocztowy (PNA)"
]]
print("Useless columns deleted")

# add ID columns to each row (ID must contain small letters and no Polish letters)
db["Kod województwa"] = ""
db["Kod powiatu"] = ""
db["Kod gminy"] = ""
db["Kod miejscowości"] = ""
db["Kod cechy adresu"] = ""
db["Kod adresu"] = ""
print("Empty ID columns created.")

for i in range(len(db)):
    # kod województwa - province_województwo
    db["Kod województwa"][i] = "province_" + removeAccents(db["Województwo"][i].lower().replace(" ", "_").replace(".","").replace("-","_"))
    # kod powiatu - district_województwo_powiat
    db["Kod powiatu"][i] = "district_" + removeAccents(db["Województwo"][i].lower()) + "_" + removeAccents(db["Powiat"][i].lower().replace(" ", "_").replace(".","").replace("-","_"))
    # kod gminy - commune_powiat_gmina
    db["Kod gminy"][i] = "commune_" + removeAccents(db["Powiat"][i].lower()) + "_" + removeAccents(db["Gmina"][i].lower().replace(" ", "_").replace(".","").replace("-","_"))
    # kod miejscowości - town_gmina_miejscowość
    db["Kod miejscowości"][i] = "town_" + removeAccents(db["Gmina"][i].lower()) + "_" + removeAccents(db["Miejscowość (GUS)"][i].lower().replace(" ", "_").replace(".","").replace("-","_"))
    # kod cechy adresu - address_prefix_'cecha adresu'
    if isinstance(db["Ulica (cecha)"][i], str):
        db["Kod cechy adresu"][i] = "address_prefix_" + removeAccents(db["Ulica (cecha)"][i].lower().replace(" ", "_").replace(".","").replace("-","_"))
    # kod adresu - address_miejscowość_ulica
    if isinstance(db["Ulica (nazwa)"][i], str):
        db["Kod adresu"][i] = "address_" + removeAccents(db["Miejscowość (GUS)"][i].lower().replace(" ", "_").replace(".","")) + "_" + removeAccents(db["Ulica (nazwa)"][i].lower().replace(" ", "_").replace(".","").replace("-","_"))
    print("IDs created for record: " + str(i))
print("IDs created for all the records.")

# save DB in case of error
db.to_excel("sample_results.xlsx")
print('DB saved.')

# make sure that dtype is correct for further manipulation
db['Kod pocztowy (PNA)'] = db['Kod pocztowy (PNA)'].astype(str)
print('Column "Kod pocztowy (PNA)" set to string format.')

# group rows by column "Miejscowość (GUS)", stacking unique values in column "kod pocztowy"
db_miasto = pd.DataFrame(db.groupby([
    'Województwo', 
    'Powiat', 
    'Gmina', 
    'Miejscowość (GUS)', 
    'Kod gminy', 
    'Kod miejscowości'], as_index=False)['Kod pocztowy (PNA)'].apply(lambda x: ",".join(x)))
print("DataFrame for cities created")

for i in range(len(db_miasto)):
    db_miasto['Kod pocztowy (PNA)'][i] = list(dict.fromkeys(list(db_miasto['Kod pocztowy (PNA)'][i].split(","))))
    db_miasto['Kod pocztowy (PNA)'][i] = ",".join(db_miasto['Kod pocztowy (PNA)'][i])
    print("Unique codes are saved for city: " + str(i))
print("Unique codes are saved for all cities.")

db_miasto =  db_miasto.filter([
    'Gmina', 
    'Kod gminy', 
    'Miejscowość (GUS)', 
    'Kod pocztowy (PNA)', 
    "Kod miejscowości"], axis=1)
print("Shape for cities file set.")

# group rows by column "Ulica (nazwa)", stacking unique values in column "kod pocztowy"
db_adres = pd.DataFrame(db.groupby([
    'Województwo', 
    'Powiat', 
    'Gmina', 
    'Miejscowość (GUS)', 
    'Ulica (cecha)', 
    'Ulica (nazwa)', 
    'Kod miejscowości', 
    'Kod cechy adresu', 
    'Kod adresu'], as_index=False)['Kod pocztowy (PNA)'].apply(lambda x: ",".join(x)))
print("DataFrame for addresses created")

for i in range(len(db_adres)):
    db_adres['Kod pocztowy (PNA)'][i] = list(dict.fromkeys(list(db_adres['Kod pocztowy (PNA)'][i].split(","))))
    db_adres['Kod pocztowy (PNA)'][i] = ",".join(db_adres['Kod pocztowy (PNA)'][i])
    print("Unique codes are saved for address: " + str(i))
print("Unique codes are saved for all addresses.")

db_adres = db_adres.filter([
    'Ulica (nazwa)', 
    'Kod pocztowy (PNA)', 
    'Miejscowość (GUS)', 
    'Kod miejscowości', 
    'Ulica (cecha)', 
    'Kod cechy adresu', 
    'Kod adresu'], axis=1)
print("Shape for addresses file set.")

# create xlsx file "ADRES"
db_adres['Operation'] = 1
db_adres.set_index('Operation', inplace=True)
db_adres.columns = [
    'Adres (Wymagalność: Nie,  code:description, sort:description_01_, type:TEXT)', 
    'Kod pocztowy (Wymagalność: Tak,  code:subject, sort:subject_01_, type:TEXT)', 
    'Miejscowość (Wymagalność: Tak,  code:town, sort:town_01_Lista_(etykieta), type:LIST)', 
    'Miejscowość (Wymagalność: Tak,  code:town, sort:town_02_Lista_(kod), type:LIST)',
    'Cecha Adresu (Wymagalność: Tak,  code:address_prefix, sort:address_prefix_01_Lista_(etykieta), type:LIST)',
    'Cecha Adresu (Wymagalność: Tak,  code:address_prefix, sort:address_prefix_02_Lista_(kod), type:LIST)',
    'Unikalny kod dokumentu (Wymagalność: Tak,  code:code, sort:code_01_, type:TEXT)'
]
db_adres.to_excel('adresy_aktualne/ADRES.xlsx')
print("ADRES.xlsx saved successfully.")

# create xlsx file "CECHA ADRESU"
db_prefix = pd.DataFrame(db[[
    'Ulica (cecha)', 
    'Kod cechy adresu'
]])
db_prefix.dropna(inplace=True)
db_prefix['Operation'] = 1
db_prefix.set_index('Operation', inplace=True)
db_prefix.drop_duplicates(inplace=True)
db_prefix.columns = [
    'Adres cecha (Wymagalność: Tak,  code:subject, sort:subject_01_, type:TEXT)',
    'Unikalny kod dokumentu (Wymagalność: Tak,  code:code, sort:code_01_, type:TEXT)'
]
db_prefix.to_excel('adresy_aktualne/CECHA ADRESU.xlsx')
print("CECHA ADRESU.xlsx saved successfully.")

# create xlsx file "MIASTO"
db_miasto['Operation'] = 1
db_miasto.set_index('Operation', inplace=True)
db_miasto.columns = [
    'Gmina (Wymagalność: Tak,  code:commune, sort:commune_01_Lista_(etykieta), type:LIST)',
    'Gmina (Wymagalność: Tak,  code:commune, sort:commune_02_Lista_(kod), type:LIST)',
    'Miejscowość (Wymagalność: Tak,  code:town, sort:town_01_, type:TEXT)',
    'Kod pocztowy (Wymagalność: Tak,  code:postcode, sort:postcode_01_, type:TEXT)',
    'Unikalny kod dokumentu (Wymagalność: Tak,  code:code, sort:code_01_, type:TEXT)'
]
db_miasto.to_excel('adresy_aktualne/MIASTO.xlsx')
print("MIASTO.xlsx saved successfully.")

# create xlsx file "GMINA"
db_gmina = pd.DataFrame(db[[
    'Gmina', 
    'Powiat', 
    'Kod powiatu', 
    'Kod gminy'
]])
db_gmina.dropna(inplace=True)
db_gmina['Operation'] = 1
db_gmina.set_index('Operation', inplace=True)
db_gmina.drop_duplicates(inplace=True)
db_gmina.columns = [
    'Gmina (Wymagalność: Tak,  code:commune, sort:commune_01_, type:TEXT)',
    'Powiat (Wymagalność: Tak,  code:district, sort:district_01_Lista_(etykieta), type:LIST)',
    'Powiat (Wymagalność: Tak,  code:district, sort:district_02_Lista_(kod), type:LIST)',
    'Unikalny kod dokumentu (Wymagalność: Tak,  code:code, sort:code_01_, type:TEXT)'
]
db_gmina.to_excel('adresy_aktualne/GMINA.xlsx')
print("GMINA.xlsx saved successfully.")

# create xlsx file "POWIAT"
db_powiat = pd.DataFrame(db[[
    'Powiat', 
    'Województwo', 
    'Kod województwa', 
    'Kod powiatu'
]])
db_powiat.dropna(inplace=True)
db_powiat['Operation'] = 1
db_powiat.set_index('Operation', inplace=True)
db_powiat.drop_duplicates(inplace=True)
db_powiat.columns = [
    'Powiat (Wymagalność: Tak,  code:district, sort:district_01_, type:TEXT)',
    'Województwo (Wymagalność: Tak,  code:province, sort:province_01_Lista_(etykieta), type:LIST)',
    'Województwo (Wymagalność: Tak,  code:province, sort:province_02_Lista_(kod), type:LIST)',
    'Unikalny kod dokumentu (Wymagalność: Tak,  code:code, sort:code_01_, type:TEXT)'
]
db_powiat.to_excel('adresy_aktualne/POWIAT.xlsx')
print("POWIAT.xlsx saved successfully.")

# create xlsx file "WOJEWÓDZTWO"
db_woje = pd.DataFrame(db[[
    'Województwo', 
    'Kod województwa'
]])
db_woje.dropna(inplace=True)
db_woje['Operation'] = 1
db_woje.set_index('Operation', inplace=True)
db_woje.drop_duplicates(inplace=True)
db_woje.columns = [
    'Województwo (Wymagalność: Tak,  code:province, sort:province_01_, type:TEXT)',
    'Unikalny kod dokumentu (Wymagalność: Tak,  code:code, sort:code_01_, type:TEXT)'
]
db_woje.to_excel('adresy_aktualne/WOJEWÓDZTWO.xlsx')
print("WOJEWÓDZTWO.xlsx saved successfully.")

# communicate the end of the process
print("Process finished.")