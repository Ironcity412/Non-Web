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

myDSNvan = "removed for security---please fix before run."
mySETvan = "set schema 'pub'"


Set cn = CreateObject("ADODB.Connection") '*******Create Connection*******
Set rs = CreateObject("ADODB.Recordset") '********Create Recordset******

cn.Open myDSNvan

strSQL = "SELECT JobProd.Company, JobHead.Plant, JobProd.JobNum, JobProd.OrderNum, JobProd.OrderLine, JobProd.OrderRelNum,JobProd.ProdQty,JobProd.WareHouseCode,JobProd.TargetJobNum,JobProd.TargetAssemblySeq,JobProd.TargetMtlSeq,JobProd.ShippedQty,JobProd.ReceivedQty,JobHead.PartNum  FROM JobProd INNER JOIN JobHead ON JobProd.Company = JobHead.Company AND JobProd.JobNum = JobHead.JobNum WHERE JobHead.JobClosed = 0 AND JobHead.Company ='BTest' AND JobHead.Plant = 'Vantage'"


MsgBox ("Started")

Set rs = cn.Execute(strSQL)

Set xl = CreateObject("Excel.Application")'***************************Create excel object********
Set xlBk = xl.Workbooks.Add '*****************************Create Workbook********

With xlbk.Worksheets(1)
    For i = 0 To rs.Fields.Count - 1
        .Cells(1, i + 1) = rs.Fields(i).Name
    Next

    .Cells(2, 1).CopyFromRecordset rs
    .Cells(1,15).Value = "MakeToType" '***********Add a Column*****
    .Cells(1,16).Value = "MakeToStockQty" '***********Add a Column*****
End With

'***************************************Compare and Fill MakeToType*************

Dim objWorksheet
Set objWorksheet = xlBk.Worksheets(1)

Const CompanyFrom = "BTest"
Const CompanyTo = "90468"
Const PlantFrom= "VANTAGE"
Const PlantTo= "MfgSys"
objWorksheet.Cells.Replace CompanyFrom, CompanyTo
objWorksheet.Cells.Replace PlantFrom, PlantTo
'***********************************************************************
 
Dim z
Dim OrderCheck 
Dim result
Dim LastRow

LastRow = objWorksheet.UsedRange.Rows.Count 

For z = LastRow To 2 Step -1
OrderCHeck = objWorksheet.Cells(z,4).Value

    If OrderCHeck > 0 Then
    
       objWorksheet.Cells(z,15).Value = "ORDER"
    Else
      objWorksheet.Cells(z,15).Value = "STOCK"
      objWorksheet.Cells(z,16).Value = objWorksheet.Cells(z,7).Value
    End If
    
Next
 '***********************Find and Replace ************** 


'***************************CLOSE OUT****************************
MsgBox ("ExportComplete")
rs.Close
Set Rs = nothing
cn.Close
Set Cn = nothing
xl.ActiveWorkBook.SaveAs("C:\Users\jmm\Desktop\Epicor10Pilot5Data\DMTData4Import\JobPRODForImport.Xlsx")
xl.ActiveWorkBook.Close
