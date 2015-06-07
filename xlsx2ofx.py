'''
Created on 7 Jun 2015

@author: Mike
'''

import csv     
import sys  
import os



try:
    in_file_name = sys.argv[1]    
    out_file_name = sys.argv[2]
except:
    print("Usage: python create_wbs.py <input_csv_file> <output_xlsx_file>")
    print("Where <input_csv_file is export from RTC query")
    print("      <output_csv_file> will contain report from script")
    sys.exit(0)


# Load the input file
try:
    print("Opening File...")
    input_file = open(in_file_name, encoding='utf-16') # opens the csv file
except IOError:
    print ('Error. Cannot open', in_file_name)
    sys.exit(0)

print("Reading File...")
reader = csv.reader(input_file,delimiter='\t')  # creates the reader 
raw_data = [r for r in reader]
# Getting creation time to be used as the basis of serial number for records.
serial_num=int(os.path.getctime(in_file_name))*1000

# Create output file
try:
    print("Opening Output File...")
    out_file = open(out_file_name, encoding='utf-8',mode='w') # opens the csv file
except IOError:
    print ('Error. Cannot open', out_file_name)
    sys.exit(0)
    
out_file.write('!Type:CCard\n')
for r in raw_data[1:]:
    print(r)
    
    date=r[0]
    description=r[1]
    amount_str=r[2]
    
    if amount_str.find(' DR')!=-1:
        amount=-1.0 * float (amount_str.replace(' DR',''))
    elif amount_str.find(' CR')!=-1:
        amount=1.0 * float (amount_str.replace(' CR',''))
    else:
        print('Amount string not recognised "'+amount_str+'"')
        exit(-1)
        
    out_file.write('D'+date+'\n') # write date
    out_file.write('P'+description+'\n') # write description
    out_file.write('T{:.2f}\n'.format(amount)) # amount
    out_file.write('N{:d}\n'.format(serial_num))
    serial_num+=1
    out_file.write('^\n') # End of record
    

input_file.close()      # close input file    
out_file.close() # close output file            
print("Done")