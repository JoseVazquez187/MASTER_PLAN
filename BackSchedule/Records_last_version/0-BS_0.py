import os
import pandas as pd

# === Detectar usuario preferente (.ONE si existe) ===
# users_dir = r"C:\Users"
# usernames = [name for name in os.listdir(users_dir) if os.path.isdir(os.path.join(users_dir, name))]
print('done')
# preferred_user = next((user for user in usernames if user.endswith(".ONE")), os.getlogin())
preferred_user = "j.vazquez"
# === Construir path base
path_base = os.path.join("C:\\Users", preferred_user, "Desktop", "PARCHE EXPEDITE")
expedite_path = os.path.join(path_base, "Expedite.csv")
print(expedite_path)
if not os.path.exists(expedite_path):
    raise FileNotFoundError(f"❌ No se encontró el archivo: {expedite_path}")
else:
    print(expedite_path)
# === Leer en chunks solo columnas necesarias ===
chunks = pd.read_csv(expedite_path, dtype=str,skiprows=1, usecols=['Demand-Source', 'Ship-Date'], chunksize=500_000)

# Unir todos los fragmentos
demand_list = []
for chunk in chunks:
    chunk['Demand-Source'] = chunk['Demand-Source'].str.upper()
    demand_list.append(chunk)

demand = pd.concat(demand_list).drop_duplicates()
demand = demand.rename(columns={'Demand-Source': 'ItemNo', 'Ship-Date': 'ReqDate'})

# === Exportar resultados ===
fcst_path = os.path.join(path_base, "FCST.xlsx")
demand.to_excel(fcst_path, index=False)

validar_path = os.path.join(path_base, "validar_vs_bom.xlsx")
demand[['ItemNo']].drop_duplicates().to_excel(validar_path, index=False)
