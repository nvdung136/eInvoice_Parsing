import xml.etree.ElementTree as ET
import os
import sys 
from csv import writer

####################################################################################

def WritingCSVout (CSVpath, DataLine):                              #Write data function
    with open(CSVpath,"a",encoding="utf-8",newline='') as f_object: #Open output csv file
        writer_object = writer(f_object)                            #Initiate the write tool
        writer_object.writerow(DataLine)                            #Write the line
    f_object.close()                                                #Close the CSV file

####################################################################################

def AppendCol (definedCols,targetBranch):
    List=[]    
    for col in definedCols:                                                #Weaving the writing line
        if targetBranch.find(col) is not None:                              #To unify the output > check each collumn respectively
            Text= targetBranch.find(col).text
            Text= Text.replace(".000000","")
            List.append(Text)                        #Weaving the output CSV line
        else:
            List.append('N/A')
    return List

####################################################################################

cols=["Ten","MST","DChi","SDThoai","DCTDTu","STKNHang","TNHang"]    #Initiation
Tien=["TgTCThue","TgTThue","TgTTTBSo"]
TTHanghoa = ["STT","MHHDVu","THHDVu","DVTinh","SLuong","DGia","TLCKhau","STCKhau","ThTien","TSuat"]
DataList = []
relative_path = "input"
csv = "NBan.csv"
IDcols = ["KHHDon","SHDon","NLap"]
TTien = []
IdList =[]
Line2W =[]
####################################################################################

if getattr(sys, 'frozen',False):
    absolute_path = os.path.dirname(sys.executable)                 #Access invoice folder when run as exe
elif __file__:
    absolute_path = os.path.dirname(__file__)                       #Access Invoice folder when run as .py file

####################################################################################

out_csv = os.path.join(absolute_path, csv)                          #Path to the csv file
path = os.path.join(absolute_path, relative_path)                   #Path to the input folder

for file in os.listdir(path):                                       #Loop through all file in the folder
    tree = ET.parse(os.path.join(path,file))                        #Parse the invoice 
    root = tree.getroot()                                           #Access the root of the XML tree
    TTChung = root[0][0]                                            #Point to TTChung section for Identifier info
    IdList = AppendCol(IDcols,TTChung)
    NDHDon = root[0][1]                                             #Point to the invoice content

###############################################################################################################
    TToan = NDHDon.find('TToan')                                    #Aim for DSHHDVu information           
    TTien = AppendCol(Tien,TToan)
    
###############################################################################################################
    NBan = NDHDon.find('NBan')                                      #Aim for NBan information
    csv = "NBan.csv"
    out_csv = os.path.join(absolute_path, csv)                      #Path to the csv file
    DataList = AppendCol(cols,NBan)
    Line2W = IdList + TTien + DataList 
    WritingCSVout(out_csv,Line2W)
    
###############################################################################################################
    NMua = NDHDon.find('NMua')                                      #Aim for NMua information
    csv = "NMua.csv"
    out_csv = os.path.join(absolute_path, csv)                      #Path to the csv file
    DataList = AppendCol(cols,NMua)
    Line2W = IdList + TTien + DataList
    WritingCSVout(out_csv,Line2W)
    
##################################################################################
    DSHHDVu = NDHDon.find('DSHHDVu')                                #Aim for NMua information
    csv = "DSHang.csv"
    out_csv = os.path.join(absolute_path, csv)                      #Path to the csv file
    Line2W = IdList + 10*["***"]
    WritingCSVout(out_csv,Line2W)
    for Child in DSHHDVu:
        DataList = AppendCol(TTHanghoa,Child)
        Line2W = 3*["***"] + DataList
        WritingCSVout(out_csv,Line2W)
    Line2W = IdList + 8*["***"] + TTien
    WritingCSVout(out_csv,Line2W)

    

    
