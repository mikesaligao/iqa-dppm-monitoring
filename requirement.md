open the ThisWorkbook and look for IQA Database sheet
columns in particulars
B = Supplier Name
C = Part Number
F = Quantity In
K = Shipment Date
O = Inspected By
AB = End Date
BC = Reject Quantity

for each data in the IQA database until the last row, generate new table in another sheet 'dppm'
Heres the table schema (header)
Shipment Date | Supplier Name | Part Number | Inspected By | Overall Quantity Received | Overall Units Reject | Overall DPPM

whereas,  Overall DPPM  = Overall Rejects / Overall Quantity Received *1M

The Shipment Date entries will be sort from Oldest to Newest.
The Shipment Date - Supplier Name - Material should be unique and will be one  entries in the table.
All the data for that key like Quantity In and Reject Quantity will be added together and the sum will be entered in the corresponding table.