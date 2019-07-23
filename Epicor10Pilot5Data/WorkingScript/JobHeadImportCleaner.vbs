'******************* THIS WILL MAKE THE JOB OPEN AND UN ENGINEERED AND UNRELEASED.

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

myDSNvan = "removed for security"
mySETvan = "set schema 'pub'"


Set cn = CreateObject("ADODB.Connection") '*******Create Connection*******
Set rs = CreateObject("ADODB.Recordset") '********Create Recordset******

cn.Open myDSNvan

strSQL = "SELECT JobHead.JobNum, JobHead.PartNum, JobHead.PartDescription, JobHead.ReqDueDate, JobHead.JobEngineered, JobHead.CommentText, JobHead.JobReleased, JobHead.Plant, JobHead.Company, JobHead.UserDate1 FROM JobHead WHERE JobClosed = 0 AND Company ='BTest' And Plant = 'Vantage'"

MsgBox ("Started")

Set rs = cn.Execute(strSQL)

Set xl = CreateObject("Excel.Application")'***************************Create excel object********
Set xlBk = xl.Workbooks.Add '*****************************Create Workbook********

With xlbk.Worksheets(1)
    For i = 0 To rs.Fields.Count - 1
        .Cells(1, i + 1) = rs.Fields(i).Name
    Next

    .Cells(2, 1).CopyFromRecordset rs
    .Cells(1,11).Value = "JobComplete" '***********Add a Column*****
    .Cells(1,10).Value = "EngDate_c"
    .Cells(1,12).Value = "InCopyList"
End With

'*****************Fix Cell Format***************

With xl
  .Sheets(1).Columns("D").NumberFormat = "mm/dd/yyyy"

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

'********************Fill New Column
Dim z
Dim Completed
Dim result
Dim LastRow

LastRow = objWorksheet.UsedRange.Rows.Count 

For z = LastRow To 2 Step -1
       objWorksheet.Cells(z,5).Value = "False" '<-------------UnEngineer Jobs
       objWorksheet.Cells(z,7).Value = "False" '<-------------UnRelease Jobs
       objWorksheet.Cells(z,11).Value = "False" '<------------UnComplete Jobs
       objWorksheet.Cells(z,12).Value = 1 '<------------Template Jobs
Next

'***************************CLOSE OUT****************************
MsgBox ("ExportComplete")
rs.Close
Set Rs = nothing
cn.Close
Set Cn = nothing
xl.ActiveWorkBook.SaveAs("C:\Users\jmm\Desktop\Epicor10Pilot5Data\DMTData4Import\JobHeadForImport.Xlsx")
