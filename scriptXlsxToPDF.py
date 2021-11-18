import win32com.client
import os
import time

xlApp = win32com.client.Dispatch("Excel.Application")

xlApp.Visible = False

wb_path = r'C:\Users\LAP-338\SynologyDrive\git\kobo-to-paperform\Humanitarion_Situation_Overview_Syria_(HSOS)_October_2021_Questionnaire.xlsx'


excelFile = xlApp.Workbooks.Open(wb_path)

xlModule = excelFile.VBProject.VBComponents.Add(1)

VBACode = '''Sub PDF_Creation()
        With Worksheets("paperEnglish").PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With

        Worksheets("paperEnglish").ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:="C:/Users/LAP-338/SynologyDrive/git/kobo-to-paperform/PDFVersion.pdf", _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False
        End Sub'''

xlModule.CodeModule.AddFromString(VBACode)

excelFile.SaveAs('Table', FileFormat=52)
excelFile.Close()

xlApp.Workbooks.Open(Filename='Table.xlsm', ReadOnly=1)

xlApp.Application.Run('Table.xlsm!PDF_Creation')

xlApp.Application.Quit()



file_deleted = False
while file_deleted is False:
    try:
        os.remove('Table.xlsm')
        file_deleted = True
    except WindowsError:
        time.sleep(0.5)
