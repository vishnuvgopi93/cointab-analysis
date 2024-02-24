import numpy as np
import pandas as pd

# import csv files 
xorder_report = pd.read_excel(r"C:\Users\User\Desktop\coin\Company X - Order Report.xlsx")
xpincode = pd.read_excel(r"C:\Users\User\Desktop\coin\Company X - Pincode Zones.xlsx")
xweight = pd.read_excel(r"C:\Users\User\Desktop\coin\Company X - SKU Master.xlsx")
courier = pd.read_excel(r"C:\Users\User\Desktop\coin\Courier Company - Invoice.xlsx")

#handle duplicates
duplicate_weight = xweight[xweight.duplicated(subset=['SKU'])]
xweight = xweight.drop_duplicates(subset=['SKU'])
duplicate_pincode = xpincode[xpincode.duplicated(subset=['Customer Pincode'])]
xpincode = xpincode.drop_duplicates(subset=['Customer Pincode'])

# merge company x's sku master with company x's invoice
df = pd.merge(xorder_report, xweight, on='SKU')
df["Total Weight as per X(kg)"] = ((df["Order Qty"] * df["Weight (g)"])/1000)

# group by externorderno and aggregate quantities and weights
aggr = df.groupby('ExternOrderNo').agg({'Order Qty': 'sum','Total Weight as per X(kg)':'sum'}).reset_index()

df = pd.merge(xpincode, courier, on='Customer Pincode', how='inner') 
df = df.rename(columns={'Zone_x': 'Delivery Zone as per X', 'Zone_y': 'Delivery Zone charged by Courier Company'})

df = pd.merge(aggr, df,left_on = "ExternOrderNo",right_on="Order ID", how="left")
df.drop(columns=['ExternOrderNo','Warehouse Pincode_x', 'Warehouse Pincode_y', 'Order Qty','Customer Pincode',], inplace=True)
df.rename(columns= {"Billing Amount (Rs.)":"Charges Billed by Courier Company (Rs.)","Charged Weight":"Total weight as per Courier Company (KG)"}, inplace = True)

df["Delivery Zone as per X"].unique()

#########################################################################
#creating zone for assinging weight slabs for company x
zone_b = df[df["Delivery Zone as per X"] == 'b']
zone_d = df[df["Delivery Zone as per X"] == 'd']
zone_e = df[df["Delivery Zone as per X"] == 'e']

def weight_slab_b(weight):
    i = round(weight % 1, 2)
    if i == 0.00:
        return weight
    elif weight<1:
        return 1 
    else: 
        return (weight//1)*1+1

#applying weight slabs as per rate card
zone_b["Weight slab as per X (KG)"] = zone_b["Total Weight as per X(kg)"].apply(weight_slab_b)


def weight_slab_d(weight):
    i = round(weight % 1.5, 2)
    if i == 0.00:
        return weight
    elif weight<1.5:
        return 1.5 
    else: 
        return (weight//1.5)*1.5+1.5
#applying weight slabs as per rate card

zone_d["Weight slab as per X (KG)"] = zone_d["Total Weight as per X(kg)"].apply(weight_slab_d)


def weight_slab_e(weight):
    i = round(weight % 2, 2)
    if i == 0.00:
        return weight
    elif weight<2:
        return 2 
    else: 
        return (weight//2)*2+2
#applying weight slabs as per rate card

zone_e["Weight slab as per X (KG)"] = zone_e["Total Weight as per X(kg)"].apply(weight_slab_e)

data1 = pd.concat([zone_d , zone_b , zone_e])

#caluclation for additional weight slabs
def calculate_additional_weight_slabs(zone, weight):
    weight_slabs = {'a': 0.5, 'b': 1, 'c': 1.25, 'd': 1.5, 'e': 2}
    zone = zone.strip().lower()
    base_slab = weight_slabs.get(zone)
    if base_slab is None:
        return "Invalid zone", None
    
    additional_weight = max(0, weight - base_slab)
    additional_weight_slabs = additional_weight // base_slab
    return additional_weight_slabs

# apply the calculation to each row
data1['Additional Weight SlabsX'] = data1.apply(lambda row: calculate_additional_weight_slabs(row['Delivery Zone as per X'], row['Weight slab as per X (KG)']), axis=1)

# define the rate card
rate_card = {
    'a': {'Weight Slabs': 0.5, 'Forward Fixed Charge': 29.5, 'Forward Additional Weight Slab Charge': 23.6, 'RTO Fixed Charge': 13.6, 'RTO Additional Weight Slab Charge': 23.6},
    'b': {'Weight Slabs': 1, 'Forward Fixed Charge': 33, 'Forward Additional Weight Slab Charge': 28.3, 'RTO Fixed Charge': 20.5, 'RTO Additional Weight Slab Charge': 28.3},
    'c': {'Weight Slabs': 1.25, 'Forward Fixed Charge': 40.1, 'Forward Additional Weight Slab Charge': 38.9, 'RTO Fixed Charge': 31.9, 'RTO Additional Weight Slab Charge': 38.9},
    'd': {'Weight Slabs': 1.5, 'Forward Fixed Charge': 45.4, 'Forward Additional Weight Slab Charge': 44.8, 'RTO Fixed Charge': 41.3, 'RTO Additional Weight Slab Charge': 44.8},
    'e': {'Weight Slabs': 2, 'Forward Fixed Charge': 56.6, 'Forward Additional Weight Slab Charge': 55.5, 'RTO Fixed Charge': 50.7, 'RTO Additional Weight Slab Charge': 55.5}
}

# calculate the total charge for each row
def calculate_total_charge(row):
    zone = row['Delivery Zone charged by Courier Company']
    additional_slabs = row['Additional Weight SlabsX']
    shipment_type = row['Type of Shipment']
    
    # base charge based on the delivery zone
    base_charge = rate_card[zone]['Forward Fixed Charge']
    
    # additional charge based on additional weight slabs
    additional_charge = additional_slabs * rate_card[zone]['Forward Additional Weight Slab Charge']
    
    # rto check
    if shipment_type == 'Forward and RTO charges':
        additional_charge += rate_card[zone]['RTO Fixed Charge']
        additional_charge += additional_slabs * rate_card[zone]['RTO Additional Weight Slab Charge']
    
    #  total charge
    total_charge = base_charge + additional_charge
    
    return total_charge

# apply the calculation to each row
data1['Expected Charge as per X (Rs.)'] = data1.apply(calculate_total_charge, axis=1)
################################################################################
#COURIER
#zones for courier
zone_B = df[df["Delivery Zone charged by Courier Company"] == 'b']
zone_D = df[df["Delivery Zone charged by Courier Company"] == 'd']
zone_E = df[df["Delivery Zone charged by Courier Company"] == 'e']

def weight_slab_B(weight):
    i = round(weight % 1, 2)
    if i == 0.00:
        return weight
    elif weight<1:
        return 1 
    else: 
        return (weight//1)*1+1
#applying weight slabs as per rate card
zone_B["Weight slab charged by Courier Company (KG)"] = zone_b["Total weight as per Courier Company (KG)"].apply(weight_slab_B)


def weight_slab_D(weight):
    i = round(weight % 1.5, 2)
    if i == 0.00:
        return weight
    elif weight<1.5:
        return 1.5 
    else: 
        return (weight//1.5)*1.5+1.5
#applying weight slabs as per rate card
zone_D["Weight slab charged by Courier Company (KG)"] = zone_D["Total weight as per Courier Company (KG)"].apply(weight_slab_D)


def weight_slab_E(weight):
    i = round(weight % 2, 2)
    if i == 0.00:
        return weight
    elif weight<2:
        return 2 
    else: 
        return (weight//2)*2+2

#applying weight slabs as per rate card
zone_E["Weight slab charged by Courier Company (KG)"] = zone_E["Total weight as per Courier Company (KG)"].apply(weight_slab_E)

data2 = pd.concat([zone_B , zone_D , zone_E])

# apply the calculation to each row
data2['Additional Weight SlabsC'] = data2.apply(lambda row: calculate_additional_weight_slabs(row['Delivery Zone charged by Courier Company'], row['Weight slab charged by Courier Company (KG)']), axis=1)

# function to calculate the total charge for each row
def calculate_total_charge(row):
    zone = row['Delivery Zone charged by Courier Company']
    additional_slabs = row['Additional Weight SlabsC']
    shipment_type = row['Type of Shipment']
    
    # determine base charge based on the delivery zone
    base_charge = rate_card[zone]['Forward Fixed Charge']
    
    # calculate additional charge based on additional weight slabs
    additional_charge = additional_slabs * rate_card[zone]['Forward Additional Weight Slab Charge']
    
    # check if rto is there
    if shipment_type == 'Forward and RTO charges':
        additional_charge += rate_card[zone]['RTO Fixed Charge']
        additional_charge += additional_slabs * rate_card[zone]['RTO Additional Weight Slab Charge']
    
    # calculate total charge
    total_charge = base_charge + additional_charge
    
    return total_charge

# apply the calculation to each row
data2['Total Charge BY C'] = data2.apply(calculate_total_charge, axis=1)

# merge data1 and data2 by order id
columns_to_merge = ['Order ID', 'Additional Weight SlabsC', 'Total Charge BY C','Weight slab charged by Courier Company (KG)']

# merge data1 and selected columns from data2 by order id
merged_data = pd.merge(data1, data2[columns_to_merge], on='Order ID', how='outer')

# calculate the difference between expected charges and billed charges
merged_data['Difference Between Expected Charges and Billed Charges (Rs.)'] = merged_data['Charges Billed by Courier Company (Rs.)'] - merged_data['Expected Charge as per X (Rs.)']

# reorder columns
desired_order = ['Order ID', 'AWB Code', 'Total Weight as per X(kg)', 'Weight slab as per X (KG)', 'Total weight as per Courier Company (KG)', 'Weight slab charged by Courier Company (KG)', 'Delivery Zone as per X', 'Delivery Zone charged by Courier Company', 'Expected Charge as per X (Rs.)', 'Charges Billed by Courier Company (Rs.)', 'Difference Between Expected Charges and Billed Charges (Rs.)']
merged_data = merged_data.reindex(columns=desired_order)

# save to csv
merged_data.to_csv("cointabanalysis.csv", index=False)

merged_data['Charge Status'] = np.where(merged_data['Difference Between Expected Charges and Billed Charges (Rs.)'] == 0, 'Correctly Charged',
                                        np.where(merged_data['Difference Between Expected Charges and Billed Charges (Rs.)'] > 0, 'Overcharged', 'Undercharged'))

#summary table
grouped = merged_data.groupby('Charge Status')

# calculate counts and total amounts for each type of charge
summary = pd.DataFrame({'Count': grouped.size(),'Amount (Rs.)': grouped['Total Invoice Amount'].sum()})

# display the summary table
print(summary)
summary.to_excel('summarytable.xlsx', index=True)  


# pie chart for visualization
import matplotlib.pyplot as plt
summary['Count'].plot(kind='pie', figsize=(8, 8), autopct='%1.1f%%', startangle=90)
plt.title('Summary of Charges', y=1.08)
plt.axis('equal')
plt.tight_layout()
plt.show()
