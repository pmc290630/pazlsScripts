import pathlib as p
import tkinter
import tkinter.filedialog
from numpy.core.numeric import NaN
import pandas as pd
import numpy as np

#test= "https://pazls.de/products/pazls-desk-one?variant=37857734426820"
def cutSku(link):
    sku = link[link.find("/products/")+10:]
    return sku

#get FilePath
root= tkinter.Tk()
root.withdraw()
filename = tkinter.filedialog.askopenfilename()

#get Name of the the File
name = str(filename)
l = name[len(name)-1]
i=1
while(l != "/"):
    i+=1
    l = name[len(name)-i]
    if l ==".":
        name = name[0:len(name)-i]
        i=0
name=name[len(name)-i+1:]

#create new file in location and call it the same
path = p.Path.home()
path = str(path)+ "/Desktop/"
#path = "/Users/Shared/"
myFile = open( path+ name +".csv","w")

# if myFile.closed:
#     tkinter.messagebox.showerror(title="closed", message= "was für ein mist") 
#     sys.exit(-1) 


#open 
df = pd.read_excel(str(filename)) #df =datafile
#print(df.head())

#TestCode
myFile.write("Handle,Title,Body (HTML),Vendor,Type,Tags,Published,Option1 Name,Option1 Value,Option2 Name,Option2 Value,Option3 Name,Option3 Value,Variant SKU,Variant Grams,Variant Inventory Tracker,Variant Inventory Qty,Variant Inventory Policy,Variant Fulfillment Service,Variant Price,Variant Compare At Price,Variant Requires Shipping,Variant Taxable,Variant Barcode,Image Src,Image Position,Image Alt Text,Gift Card,SEO Title,SEO Description,Google Shopping / Google Product Category,Google Shopping / Gender,Google Shopping / Age Group,Google Shopping / MPN,Google Shopping / AdWords Grouping,Google Shopping / AdWords Labels,Google Shopping / Condition,Google Shopping / Custom Product,Google Shopping / Custom Label 0,Google Shopping / Custom Label 1,Google Shopping / Custom Label 2,Google Shopping / Custom Label 3,Google Shopping / Custom Label 4,Variant Image,Variant Weight Unit,Variant Tax Code,Cost per item,Status")

df["Variant SKU"] = df["Variant SKU"].apply(cutSku)
#print(df["Variant SKU"])
#df["Body (HTLM)"] = '""'
# print(str(df["Body (HTML)"]))

writer = pd.ExcelWriter(path+"linksBereinigt_AlteDateiBleibtMutter_"+name+".xlsx") #path + name
df.to_excel(writer,name) #sheet
writer.save()

rows=df.shape[0]
cols = df.shape[1]

df = df.replace(np.nan, "",regex=True)

columnsToBeInQuotationMarks = [0,1,2,5,23]

for r in range(0,rows):
    line = ""
    for c in range(0,cols):
        if line != "": line+=","
        cel = str(df.iat[r,c])
        if c in columnsToBeInQuotationMarks:
            line+='"'+cel+'"'
        else:
            line +=cel
    myFile.write("\n"+line)


#df.to_csv(path + "neu" + name + ".csv") 