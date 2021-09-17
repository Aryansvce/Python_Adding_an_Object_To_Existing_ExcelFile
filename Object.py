##pip install pypiwin32 to work with windows operating sysytm and import the module as mentioned below.
import win32com.client
import time
# Creating an object for accessing excel application.
excel_app = win32com.client.Dispatch('Excel.Application')
# Set visible as 1. It is required to perform the desired task.
excel_app.Visible = 1
# Open the excel workbook from the desired location in read mode.
workbook = excel_app.Workbooks.Open(r"E:\RPA_IGSO\Python\Output\EO_TRUEUP_APJ\JE Output\SG20_Test.xlsx")
# Select worksheet by name.
worksheet = workbook.Sheets('Supporting Documents')
# To assign an object for OLEObject(=EMBED("Packager Shell Object","")).
Embedded_object = worksheet.OLEObjects()
# To assign loction of the image file that need to inserted as OBJECT in excel worksheet.
file_loction = r"E:\RPA_IGSO\Python\Input\EO_TRUEUP_APJ\EO_TRUUP_MAIL_ATTACHMENT\E&O True Up.zip"
dest_cell = worksheet.Range("B16")

# To add selected file to the excel worksheet. It will add the OBJECT to the A1 cell of the current worksheet.
Embedded_object.Add(ClassType=None, Filename=file_loction, Link=False, DisplayAsIcon=True,Left=dest_cell.Left, Top=dest_cell.Top, Width=50, Height=50)
# To Copy selected range of cells in the current worksheet.
# worksheet.Range('A1:A1').Copy()
# # To paste the copied data to a perticular range of cells in currnet worksheet.
# worksheet.Paste(Destination=worksheet.Range('C2:C2'))
# # To select fist item in the list of object i.e. first object.
# obj = Embedded_object.Item(1)
#Embedded_object.Delete()

# workbook.Save()

# worksheet2 = workbook.Sheets('Sheet2')
# # To assign an object for OLEObject(=EMBED("Packager Shell Object","")).
# Embedded_object2 = worksheet2.OLEObjects()
# # To assign loction of the image file that need to inserted as OBJECT in excel worksheet.
# file_loction2 = r"D:\test_attach_object\Dashboard.zip"
# # To add selected file to the excel worksheet. It will add the OBJECT to the A1 cell of the current worksheet.
# Embedded_object2.Add(ClassType=None, Filename=file_loction2, Link=False, DisplayAsIcon=True,Left=3, Top=0, Width=50, Height=50)

# time.sleep(1)

# # To Copy selected range of cells in the current worksheet.
# worksheet2.Range('A1:A1').Copy()
# # To paste the copied data to a perticular range of cells in currnet worksheet.
# worksheet2.Paste(Destination=worksheet2.Range('C2:C2'))
# # # To select fist item in the list of object i.e. first object.
# # obj2 = Embedded_object2.Item(1)
# #Embedded_object2.Delete()

# # To delete selected object from the worksheet.
# # obj.Delete()

workbook.Save()
time.sleep(1)
workbook.Close()
excel_app.Quit()
