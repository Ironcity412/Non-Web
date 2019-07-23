Option Explicit

Dim xl
Dim xlBk
Dim i
Dim myDSNvan
Dim mySETvan
Dim Cn 'Connection
Dim rs 'Record Set
Dim strCon 'Connection Script
Dim strSQL 'Sql QUery

myDSNvan = "removed for security--- fix before run"
mySETvan = "set schema 'pub'"


Set cn = CreateObject("ADODB.Connection") '*******Create Connection*******
Set rs = CreateObject("ADODB.Recordset") '********Create Recordset******

cn.Open myDSNvan

strSQL = "SELECT JobMtl.Company, JobHead.JobNum, JobMtl.AssemblySeq, JobMtl.MtlSeq, JobMtl.PartNum, JobMtl.Description, JobMtl.QtyPer, JobMtl.RelatedOperation, JobMtl.IssuedQty, JobMtl.RequiredQty, JobMtl.FixedQty,JobMtl.Plant,JobMtl.Direct,JobMtl.BuyIt  FROM JobMtl INNER JOIN JobHead ON JobMtl.Company = JobHead.Company AND JobMtl.JobNum = JobHead.JobNum WHERE JobHead.JobClosed = 0 AND JobHead.Company ='BTest' AND JobHead.Plant = 'Vantage'"

MsgBox ("Started")

Set rs = cn.Execute(strSQL)

Set xl = CreateObject("Excel.Application")'***************************Create excel object********
Set xlBk = xl.Workbooks.Add '*****************************Create Workbook********

With xlbk.Worksheets(1)
    For i = 0 To rs.Fields.Count - 1
        .Cells(1, i + 1) = rs.Fields(i).Name
    Next

    .Cells(2, 1).CopyFromRecordset rs
    
End With
'****************Find and Replace ************** 
Dim objWorksheet
Set objWorksheet = xlBk.Worksheets(1)

Const CompanyFrom = "BTest"
Const CompanyTo = "90468"
Const PlantFrom= "VANTAGE"
Const PlantTo= "MfgSys"
objWorksheet.Cells.Replace CompanyFrom, CompanyTo
objWorksheet.Cells.Replace PlantFrom, PlantTo
'***************************CLOSE OUT****************************
MsgBox ("ExportComplete")
rs.Close
Set Rs = nothing
cn.Close
Set Cn = nothing
xl.ActiveWorkBook.SaveAs("C:\Users\jmm\Desktop\Epicor10Pilot5Data\DMTData4Import\JobMATImport.Xlsx")
xl.ActiveWorkBook.Close
