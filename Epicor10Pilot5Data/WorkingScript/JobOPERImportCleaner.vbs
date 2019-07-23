Option Explicit
PUBLIC xl
PUBLIC xlBk
Dim i
Dim myDSNvan
Dim mySETvan
Dim Cn 'Connection
PUBLIC rs 'Record Set
Dim strCon 'Connection Script
Dim strSQL 'Sql QUery
PUBLIC LastRow
PUBLIC objWorksheet


myDSNvan = "removed for security ---fix before run"
mySETvan = "set schema 'pub'"


Set cn = CreateObject("ADODB.Connection") '<------Create Connection
Set rs = CreateObject("ADODB.Recordset") '<------Create Recordset

cn.Open myDSNvan

strSQL = "SELECT  JobHead.Company, JobOper.JobNum, JobOper.AssemblySeq, JobOper.OprSeq , JobOper.WCCode, JobOper.OpCode, JobOper.ProdStandard,JobOper.QtyPer, JobOper.ActProdHours, JobOper.ActSetupHours,JobOper.QtyCompleted, JobOper.SubContract,JobOper.CommentText,JobHead.Plant,JobOper.Number01, JobOper.EstScrap, JobOper.EstScrapType,JobOper.PartNum,JobOper.Description,Vendor.VendorID FROM JobOper INNER JOIN JobHead ON JobOper.Company = JobHead.Company AND JobOper.JobNum = JobHead.JobNum LEFT JOIN Vendor ON JobOper.VendorNum = Vendor.VendorNum WHERE JobHead.JobClosed = 0 AND JobHead.Company ='BTest' AND JobHead.Plant = 'Vantage'"


MsgBox ("Started")

Set rs = cn.Execute(strSQL)
Set xl = CreateObject("Excel.Application")'<------Create excel object
Set xlBk = xl.Workbooks.Add '<------Create Workbook
Set objWorksheet = xlBk.Worksheets(1) '<------Create WorkSheet

With xlbk.Worksheets(1)
    For i = 0 To rs.Fields.Count - 1
        .Cells(1, i + 1) = rs.Fields(i).Name
    Next

    .Cells(2, 1).CopyFromRecordset rs
End With

    objworksheet.Cells(1,15).Value = "OpPieceRate_c" '<------UpdateHeader
    objworksheet.Cells(1,20).Value = "VendorNumVendorID" '<------UpdateHeader
    objworksheet.Cells(1,21).Value = "JobOpDtl#ResourceGrpID" '<------UpdateHeader
    objworksheet.Cells(1,22).Value = "JobOpDtl#ResourceID" '<------UpdateHeader


'*********************************************************Find and Replace************** 


Const CompanyFrom = "BTest"
Const CompanyTo = "90468"
Const PlantFrom= "VANTAGE"
Const PlantTo= "MfgSys"
objWorksheet.Cells.Replace CompanyFrom, CompanyTo
objWorksheet.Cells.Replace PlantFrom, PlantTo

'*********************************************************CALL MAPPING SUBS
LastRow = objWorksheet.UsedRange.Rows.Count '<------FIND THE NUMBER OF ROWS FOR MAPPING

CALL MapOPER
CALL MapWorkCenter

objworksheet.columns("E:E").Delete '<------Remove WCCode after MAPPING

'**********************************************************CLOSE OUT****************************
MsgBox ("ExportComplete")
rs.Close
Set Rs = nothing
cn.Close
Set Cn = nothing
xl.ActiveWorkBook.SaveAs("C:\Users\jmm\Desktop\Epicor10Pilot5Data\DMTData4Import\JobOPERForImport.Xlsx")
xl.ActiveWorkBook.Close


'********************************************************Compare and Fill *************
Sub MapOPER
Dim z
Dim OperationCode


For z = LastRow To 2 Step -1
    OperationCode= objWorksheet.Cells(z,6).Value

Select Case OperationCode
    CASE "VINYL"
        objWorksheet.Cells(z,6).Value = "APLYVLET"
    CASE "BAND" 
        objWorksheet.Cells(z,6).Value = "CUTTING"
    CASE "BELT"
        objWorksheet.Cells(z,6).Value = "BELTSAND"
    CASE "BEND"
        objWorksheet.Cells(z,6).Value = "BENDING"
    CASE "CLICK"
        objWorksheet.Cells(z,6).Value = "DIECUT"
    CASE "CNC"
        objWorksheet.Cells(z,6).Value = "ROUTE"
    CASE "CUT"
        objWorksheet.Cells(z,6).Value = "CUTTING"
    CASE "DIE"
        objWorksheet.Cells(z,6).Value = "DIECUT"
    CASE "HINGE"
        objWorksheet.Cells(z,6).Value = "DRILL"
    CASE "DRLMH"
        objWorksheet.Cells(z,6).Value = "DRILL"
    CASE "DRILL"
        objWorksheet.Cells(z,6).Value = "DRILL"
    CASE "EBAND"
        objWorksheet.Cells(z,6).Value = "EDGEBAND"
    CASE "EFI"
        objWorksheet.Cells(z,6).Value = "EDGEFIN"
    CASE "EMBO"
        objWorksheet.Cells(z,6).Value = "EMBOSS"
    CASE "EYEL"
        objWorksheet.Cells(z,6).Value = "EYELET"
    CASE "FLAM"
        objWorksheet.Cells(z,6).Value = "FLAMPLSH"
    CASE "FOLD"
        objWorksheet.Cells(z,6).Value = "HOTKNIFE"
    CASE "GBC"
        objWorksheet.Cells(z,6).Value = "GBCPUNCH"
    CASE "GASSY"
        objWorksheet.Cells(z,6).Value = "GENASM"
    CASE "GLUE"
        objWorksheet.Cells(z,6).Value = "GLUING"
    CASE "GROM"
        objWorksheet.Cells(z,6).Value = "GROMMET"
    CASE "HAND"
        objWorksheet.Cells(z,6).Value = "HANDSCRN"
    CASE "HEAT"
        objWorksheet.Cells(z,6).Value = "HEATSEAL"
    CASE "HOLL"
        objWorksheet.Cells(z,6).Value = "DRILL"
    CASE "RPT"
        objWorksheet.Cells(z,6).Value = "JOBRPT"
    CASE "LSLIT"
        objWorksheet.Cells(z,6).Value = "LAMSLIT"
    CASE "LAMIN"
        objWorksheet.Cells(z,6).Value = "LAMINATE"
    CASE "LASER"
        objWorksheet.Cells(z,6).Value = "LASERCUT"
    CASE "METAL"
        objWorksheet.Cells(z,6).Value = "METALWIR"
    CASE "MISC"
        objWorksheet.Cells(z,6).Value = "MISCFAB"
    CASE "PACK"
        objWorksheet.Cells(z,6).Value = "PACK"
    CASE "PAINT"
        objWorksheet.Cells(z,6).Value = "GENASM"
    CASE "PSAW"
        objWorksheet.Cells(z,6).Value = "CUTTING"
    CASE "POLI"
        objWorksheet.Cells(z,6).Value = "POLISH"
    CASE "SAW"
        objWorksheet.Cells(z,6).Value = "CUTTING"
    CASE "PRTG"
        objWorksheet.Cells(z,6).Value = "SCRNPRNT"
    CASE "SEWG"
        objWorksheet.Cells(z,6).Value = "SEWING"
    CASE "SHEE"
        objWorksheet.Cells(z,6).Value = "SHEETING"
    CASE "SLIT"
        objWorksheet.Cells(z,6).Value = "SLIPSHT"
    CASE "STRIP"
        objWorksheet.Cells(z,6).Value = "STRIP"
    CASE "STAIN"
        objWorksheet.Cells(z,6).Value = "GENASM"
    CASE "ROUT"
        objWorksheet.Cells(z,6).Value = "ROUTE"
    CASE "TSAW"
        objWorksheet.Cells(z,6).Value = "CUTTING"
    CASE "TAPE"
        objWorksheet.Cells(z,6).Value = "TAPING"
    CASE "THERM"
        objWorksheet.Cells(z,6).Value = "THERMO"
    CASE "WELD"
        objWorksheet.Cells(z,6).Value = "WELDING"
    CASE "WLAB"
        objWorksheet.Cells(z,6).Value = "GENASM"
    CASE "SNAPS"
        objWorksheet.Cells(z,6).Value = "APLYSNAP"
    CASE "CATCH"
        objWorksheet.Cells(z,6).Value = "CATCHING"   
    CASE "MAGN"
        objWorksheet.Cells(z,6).Value = "APLYMAG"
    CASE "TRIM"
        objWorksheet.Cells(z,6).Value = "TRIMBORD"
    CASE "ASSEM"
        objWorksheet.Cells(z,6).Value = "GENASM"
    CASE "STRP"
        objWorksheet.Cells(z,6).Value = "STRIP" 
    CASE "CLIPS"
        objWorksheet.Cells(z,6).Value = "APLYCLIP" 
    CASE "BNDBG"
        objWorksheet.Cells(z,6).Value = "BENDBAG"
    CASE "TAPEB"
        objWorksheet.Cells(z,6).Value = "TAPEBAG"
    CASE "STBKS"
        objWorksheet.Cells(z,6).Value = "APYSTKBK"
    CASE "VELCR"
        objWorksheet.Cells(z,6).Value = "APLYVELC"
    CASE "TAPEP"
        objWorksheet.Cells(z,6).Value = "TAPEPACK"
    CASE "BUMP"
        objWorksheet.Cells(z,6).Value = "APLYBUMP"
    CASE "CORN"
        objWorksheet.Cells(z,6).Value = "CORRND"
END SELECT
 'THESE OPERATIONS ID'S ARE THE SAME IN BOTH VERSIONS:BURN ,MASK ,EYELET ,PEEL ,RIVET
Next
End Sub

SUB MapWorkCenter
Dim c
Dim WCCode
Dim RowNumA
Dim RowNumB
RowNumA = 21
RowNumB = 22

For c = LastRow To 2 Step -1
    WCCode= objWorksheet.Cells(c,5).Value

    SELECT CASE WCCode
        CASE "VDL"
             objWorksheet.Cells(c,RowNumA).Value = "PSHOPLB"             
        CASE "BAND"
            objWorksheet.Cells(c,RowNumA).Value = "PROD SAW"
            objWorksheet.Cells(c,RowNumB).Value = "BANDSAW"
        CASE "BELT"
            objWorksheet.Cells(c,RowNumA).Value = "WS LABOR"
        CASE "AUTO"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "AUTO BENDER"
        CASE "BEND"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "BEND2"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "CLIK"
            objWorksheet.Cells(c,RowNumA).Value = "DIECUTRG"
        CASE "CNC"
            objWorksheet.Cells(c,RowNumA).Value = "CNC"
        CASE "CNCIH"
            objWorksheet.Cells(c,RowNumA).Value = "CNC"
            objWorksheet.Cells(c,RowNumB).Value = "IronHorse"
        CASE "GIL"
            objWorksheet.Cells(c,RowNumA).Value = "GUILL"
        CASE "DIE"
            objWorksheet.Cells(c,RowNumA).Value = "DIECUTRG"
        CASE "DOOR"
            objWorksheet.Cells(c,RowNumA).Value = "WS LABOR"
        CASE "DRL"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "DRILL PRESS"
        CASE "EBAND"
            objWorksheet.Cells(c,RowNumA).Value = "WS MACHINE"
            objWorksheet.Cells(c,RowNumB).Value = "EDGEBANDMACH"
        CASE "EFI"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "EFI #1"
        CASE "EMBOS"
            objWorksheet.Cells(c,RowNumA).Value = "MISCSEWM"
            objWorksheet.Cells(c,RowNumB).Value = "EMBOSSING"
        CASE "EYEL"
            objWorksheet.Cells(c,RowNumA).Value = "MISCSEWN"
            objWorksheet.Cells(c,RowNumB).Value = "EYELET"
        CASE "FLM"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "FLAMEPOLSIH"
        CASE "FLM2"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "FLAMEPOLSIH"
        CASE "HOTKF"
            objWorksheet.Cells(c,RowNumA).Value = "MISCSEWM"
            objWorksheet.Cells(c,RowNumB).Value = "HOTKNIFE" 
        CASE "ASSLM"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "ASSL2"
           objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR" 
        CASE "GLUE2"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "GLUE"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "GROM2"
            objWorksheet.Cells(c,RowNumA).Value = "GROMMACH"
        CASE "GRO"
            objWorksheet.Cells(c,RowNumA).Value = "GROMMACH"
        CASE "HAND"
            objWorksheet.Cells(c,RowNumA).Value = "PSHOPLB"
        CASE "HSL"
            objWorksheet.Cells(c,RowNumA).Value = "HEATSEAL"
        CASE "HSL2"
            objWorksheet.Cells(c,RowNumA).Value = "HEATSEAL"
        CASE "HSLIN"
            objWorksheet.Cells(c,RowNumA).Value = "HEATSEAL"
        CASE "HOL"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "HOLLOW DRILL"
        CASE "JOBCL"
            objWorksheet.Cells(c,RowNumA).Value = " "
        CASE "LSLIT"
            objWorksheet.Cells(c,RowNumA).Value = "WS MACH"
             objWorksheet.Cells(c,RowNumB).Value = "LAMSLITTER"
        CASE "LAM"
            objWorksheet.Cells(c,RowNumA).Value = "WS LABOR"
        CASE "LASER"
            objWorksheet.Cells(c,RowNumA).Value = "OUTSIDE"
            objWorksheet.Cells(c,RowNumB).Value = "FABOUT"
        CASE "METAL"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "OUTSI"
            objWorksheet.Cells(c,RowNumA).Value = "OUTSIDE"
            objWorksheet.Cells(c,RowNumB).Value = "FABOUT"
        CASE "EXP"
            objWorksheet.Cells(c,RowNumA).Value = "CSERVICE"
        CASE "EXPAC"
            objWorksheet.Cells(c,RowNumA).Value = "PACKING"
        CASE "EXPSH"
            objWorksheet.Cells(c,RowNumA).Value = "SHIPLBR"
        CASE "PACK"
            objWorksheet.Cells(c,RowNumA).Value = "PACKING"
        CASE "SPRAY"
            objWorksheet.Cells(c,RowNumA).Value = "WS LABOR"
        CASE "PSAW"
            objWorksheet.Cells(c,RowNumA).Value = "WS MACH"
            objWorksheet.Cells(c,RowNumB).Value = "PANELSAW#1"
        CASE "BUF"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "BUFFER"
        CASE "MITER"
            objWorksheet.Cells(c,RowNumA).Value = "WS LABOR"
        CASE "TSAW"
            objWorksheet.Cells(c,RowNumA).Value = "PROD SAW"
            objWorksheet.Cells(c,RowNumB).Value = "TABLESAW"
        CASE "SAW"
            objWorksheet.Cells(c,RowNumA).Value = "PROD SAW"
            objWorksheet.Cells(c,RowNumB).Value = "HENDRIX"
        CASE "PRINT"
            objWorksheet.Cells(c,RowNumA).Value = "OUTSIDE"
            objWorksheet.Cells(c,RowNumB).Value = "PRINTOUT"
        CASE "ACCUP"
            objWorksheet.Cells(c,RowNumA).Value = "PRINTSHP"
        CASE "SQU"
            objWorksheet.Cells(c,RowNumA).Value = "PSHOPLB"
        CASE "CAMEO"
            objWorksheet.Cells(c,RowNumA).Value = " "'******REMOVE****
        CASE "SEWG2"
            objWorksheet.Cells(c,RowNumA).Value = "SEWMACH"
        CASE "SEW"
            objWorksheet.Cells(c,RowNumA).Value = "SEWMACH"
        CASE "SHR"
            objWorksheet.Cells(c,RowNumA).Value = "MISCSEWM"
            objWorksheet.Cells(c,RowNumB).Value = "SHEETER"
        CASE "SLIT"
             objWorksheet.Cells(c,RowNumA).Value = "MISCSEWM"
            objWorksheet.Cells(c,RowNumB).Value = "SLITTER"
        CASE "ASSLY"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "ROU"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "ROUTER TBL#1"
        CASE "TSA"
            objWorksheet.Cells(c,RowNumA).Value = "PROD SAW"
            objWorksheet.Cells(c,RowNumB).Value = "TABLESAW"
        CASE "TAPE"
            objWorksheet.Cells(c,RowNumA).Value = "B2FLOOR"
        CASE "THERM"
            objWorksheet.Cells(c,RowNumA).Value = "THERMO"
        CASE "WELD"
            objWorksheet.Cells(c,RowNumA).Value = "ACR MACH"
            objWorksheet.Cells(c,RowNumB).Value = "PLASTICWELD"
        CASE "WOOD"
            objWorksheet.Cells(c,RowNumA).Value = "WS LABOR"
    END SELECT
Next
END Sub