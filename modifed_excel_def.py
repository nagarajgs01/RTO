

import pandas as pd
from tkinter import messagebox


def excel_function(filePath):
    
    df=pd.read_excel(filePath,header=4,index_col=0,usecols="A,D,F,H,J")
    df=df.rename(columns={"S. No":"sr","Registration Number ":"Registration","Current Address ":"Address",
                          "Mobile Number ":"Number","Owner Name ":"Name"})
    


    file_name="UpdatedInformation.xlsx"
    df.to_excel(file_name)
    
    #print("Successfully created")
    
    
    
    new_df=pd.read_excel("UpdatedInformation.xlsx",header=0,index_col=0,dtype={"Number":str})
  
    reg_no=new_df["Registration"]
    user=0
    
    for index in reg_no:
        user=user+1
        #print(index)
        details=new_df[new_df["Registration"]==index]
        pd.options.display.max_colwidth = None
        
        address=details.iloc[:,2].values.tolist()
        address=str(address)
        address=address.strip('[]').strip("'")
        pincode=address[-6:]
        address=address.strip(pincode)
        comma=address[-1:]
        address=address.strip(comma)
        #print(address,pincode)
        new_df.at[user,"Address"]=address
        new_df.at[user,"Pincode"]=pincode
        
    #print(details.head())
   
    pd.options.display.max_colwidth = None
    #print(new_df.head())
    
   
    new_df.to_excel(file_name)
    #print("File Saved")
   
def print_card(values):
    
  
    file_name="print_details.xlsx"
    
    new_excel=pd.read_excel("UpdatedInformation.xlsx",header=0,index_col=0,dtype={"Number":str})
    
    printing_sheet=pd.DataFrame(columns=["Registration","Name","Address","Number","Pincode"],data=None)
    printing_sheet.to_excel(file_name)
    
    
    #print(new_excel.head())
    printing_sheet=pd.read_excel(file_name,header=0,index_col=0,dtype={"Number":str})
    #print(printing_sheet.head())
    user=0
    for index in values: 
        user=user+1
        user_information=new_excel[new_excel["Registration"]==index]
        
        temp1=str(user_information['Registration'].values.tolist()).strip("[]").strip("'")
        printing_sheet.at[user,"Registration"]=temp1
        
        temp1=str(user_information['Name'].values.tolist()).strip("[]").strip("'")
        printing_sheet.at[user,"Name"]=temp1
        
        temp1=str(user_information['Address'].values.tolist()).strip("[]").strip("'")
        printing_sheet.at[user,"Address"]=temp1 
        
        temp1=str(user_information['Number'].values.tolist()).strip("[]").strip("'")
        printing_sheet.at[user,"Number"]=temp1 
        
        temp1=str(user_information['Pincode'].values.tolist()).strip("[]").strip("'")
        printing_sheet.at[user,"Pincode"]=temp1 
        
      
       
       
    
    
    
    printing_sheet.to_excel(file_name)
    #print("printing File created")
    
""" 
if __name__=='__main__':
    
    try: 
        
        extractedInformation=excel_function("/Users/JaySabnis/Downloads/smart card excel sheet.xlsx")
        val=input("enter values")
        new_list=list(val.split(","))
        print(new_list)
        values=[new_list]
        print_card(new_list)
    
    except FileNotFoundError:
        messagebox.showerror("Error","Resource File not found")
   """     