import pandas as pd

l_filename = "Checkoutdetails.xlsx"


file_data = pd.read_excel(l_filename)
file_data1 = file_data
dict = {}
cnt = 0
for i in file_data:
    dict[i] = str(file_data[i][cnt])
print(dict)

for i in dict:
    if i == "Expiry":
        print(int(dict[i].split()[0].split('-')[2]))
