Attribute VB_Name = "a05MassUpdateDescription"
'***************程序说明
'1.用于批量维护描述
'2.在批量创建code之后, 不需要把结果复制粘贴到A列, 而是原始运行 MassUpdateDescription
'3.
'***************程序输入
'1.在Sub LogonSAP() 选择是正式系统还是测试系统
'2.
'3.
'***************程序输出
'只是在立即窗口显示结果

Option Explicit

Dim logonControl As SAPLogonCtrl.SAPLogonControl
Dim connection As SAPLogonCtrl.connection

Public Sub LogonSAP()
    Set logonControl = New SAPLogonControl
    Set connection = logonControl.NewConnection()
    
    With connection
'        .System = "SC1"                           ' test system
'        .ApplicationServer = "172.31.92.86"  ' test system
'        .SystemNumber = "02"                     ' test system
'        .Client = "100"
'        .User = "107895"
'        .Password = "Color@do1215"
        
        .System = "SP1"                           ' formal system
        .ApplicationServer = "172.31.92.187"    ' formal system
        .SystemNumber = "00"                       ' formal system
        .Client = "100"
        .User = "107895"
        .Password = "Color@do1224"
    End With
    
    Call connection.Logon(0, True)
End Sub

Public Sub LogoffSAP()
    If Not connection Is Nothing And connection.IsConnected = tloRfcConnected Then
        connection.Logoff
    End If
End Sub

'20230605 对于massupdate 这里的数量不用改

Public Sub MassUpdateDescription()
' the 1st lind code : save the material code in sheet1 column A to plant 1005 , pre defined the parameters
' the 2nd lind code : save the material code in sheet1 column A to plant 1003, pre defined the parameters
' the 3rd lind code : save the material code in sheet1 column A to plant 1003, wont change the maintain Description , the Description parameter is blank
'    Call BAPI_MATERIAL_GETINTNUMBERRET(10, "MPR", "00601", "A", "English Description", "Chinese Description", "M", "1005", "F", "", "5001", "5001", "VX0", "4S", "5001")
'    Call BAPI_MATERIAL_GETINTNUMBERRET(10, "MPR", "00601", "A", "English Description", "Chinese Description", "M", "1003", "F", "", "3001", "3001", "H03", "EX", "3001")
    Call BAPI_MATERIAL_GETINTNUMBERRET(1, "MPR", "00601", "A", "", "", "", "1003", "F", "", "3001", "3001", "H03", "EX", "3001")
'    Call BAPI_MATERIAL_GETINTNUMBERRET(10, "MPR", "00601", "A", "", "", "这里是单位,空着也可以", "1003", "F", "", "3001", "3001", "H03", "EX", "3001")
End Sub
Public Sub BAPI_MATERIAL_GETINTNUMBERRET _
(i As Integer, MATERIAL_TYPE As String, MATERIAL_GROUP As String, INDUSTRY_SECTOR As String, ENGLISH_DESCRIPTION As String, CHINESE_DESCRIPTION, _
BASE_UOM As String, PLANT As String, PROC_TYPE As String, SPPROCTYPE As String, ISS_ST_LOC As String, SLOC_EXPRC As String, MRP_CTRLER As String, _
LOTSIZEKEY As String, STGE_LOC As String)

    Dim fms As SAPFunctionsOCX.SAPFunctions
    Dim fm2 As SAPFunctionsOCX.Function ' dim fm2
    Dim MATERIALDESCRIPTION As SAPTableFactoryCtrl.Table  'fm2 table notice
    Dim UNITSOFMEASURE As SAPTableFactoryCtrl.Table  'fm2 table notice
    Dim UNITSOFMEASUREX As SAPTableFactoryCtrl.Table  'fm2 table notice
    Dim TAXCLASSIFICATIONS As SAPTableFactoryCtrl.Table  'fm2 table notice
    

    Call LogonSAP
    
    If connection.IsConnected <> tloRfcConnected Then
        MsgBox "connection failed"
        Exit Sub
    End If
    
    Set fms = New SAPFunctions
    Set fms.connection = connection
    Set fm2 = fms.Add("BAPI_MATERIAL_SAVEDATA")
    Set MATERIALDESCRIPTION = fm2.Tables("MATERIALDESCRIPTION") 'fm2 table notice
    Set UNITSOFMEASURE = fm2.Tables("UNITSOFMEASURE") 'fm2 table notice
    Set UNITSOFMEASUREX = fm2.Tables("UNITSOFMEASUREX") 'fm2 table notice
    Set TAXCLASSIFICATIONS = fm2.Tables("TAXCLASSIFICATIONS") 'fm2 table notice
    
'*************fm2 call start
    ' for loop to make the length to 18 , no need to use if else
    Dim rng As Range, rngLen As Integer, rngToStr As String
    For Each rng In Sheet3.Range(Sheet3.Range("J2"), Sheet3.Range("J1048576").End(xlUp))
        On Error Resume Next
        If rng <> "" Then
        rngToStr = rng
        For rngLen = 1 To (18 - Len(rng))
        rngToStr = "0" & rngToStr
        Next rngLen

'************HEADDATA start
            fm2.Exports("HEADDATA")("MATERIAL") = rngToStr
            fm2.Exports("HEADDATA")("IND_SECTOR") = INDUSTRY_SECTOR
            fm2.Exports("HEADDATA")("MATL_TYPE") = MATERIAL_TYPE
'
'            fm2.Exports("HEADDATA")("BASIC_VIEW") = "X"
'            fm2.Exports("HEADDATA")("SALES_VIEW") = "X"
'            fm2.Exports("HEADDATA")("PURCHASE_VIEW") = "X"
'            fm2.Exports("HEADDATA")("MRP_VIEW") = "X"
'            fm2.Exports("HEADDATA")("FORECAST_VIEW") = "X"
'            fm2.Exports("HEADDATA")("WORK_SCHED_VIEW") = "X"
'            fm2.Exports("HEADDATA")("PRT_VIEW") = "X"
'            fm2.Exports("HEADDATA")("STORAGE_VIEW") = "X"
'            fm2.Exports("HEADDATA")("WAREHOUSE_VIEW") = "X"
'            fm2.Exports("HEADDATA")("QUALITY_VIEW") = "X"
'            fm2.Exports("HEADDATA")("ACCOUNT_VIEW") = "X"
'            fm2.Exports("HEADDATA")("COST_VIEW") = "X"
'************HEADDATA end

''************STORAGELOCATIONDATA start
'            fm2.Exports("STORAGELOCATIONDATAX")("PLANT") = PLANT
'            fm2.Exports("STORAGELOCATIONDATAX")("STGE_LOC") = ISS_ST_LOC
'            ' above is the X data , below is real data
'            fm2.Exports("STORAGELOCATIONDATA")("PLANT") = PLANT
'            fm2.Exports("STORAGELOCATIONDATA")("STGE_LOC") = ISS_ST_LOC
''************STORAGELOCATIONDATA end
'
'************MATERIALDESCRIPTION table data start
            MATERIALDESCRIPTION.FreeTable
            MATERIALDESCRIPTION.AppendRow
'            MATERIALDESCRIPTION(1, "LANGU") = "1" ' default 1 is chinese
            MATERIALDESCRIPTION(1, "LANGU_ISO") = "EN"
            MATERIALDESCRIPTION(1, "MATL_DESC") = rng.Offset(0, -2)
'            MATERIALDESCRIPTION(1, "DEL_FLAG") = "X"
            MATERIALDESCRIPTION.AppendRow
'            MATERIALDESCRIPTION(1, "LANGU") = "1"
            MATERIALDESCRIPTION(2, "LANGU_ISO") = "ZH"
            
            If InStr(rng.Offset(0, -2), "SECTOR") <> 0 Then
                MATERIALDESCRIPTION(2, "MATL_DESC") = "扇形段 " & rng.Offset(0, -1)
                
                ElseIf InStr(rng.Offset(0, -2), "COVER") <> 0 Then
                MATERIALDESCRIPTION(2, "MATL_DESC") = "壳 " & rng.Offset(0, -1)
                
                ElseIf InStr(rng.Offset(0, -2), "SHIM") <> 0 Then
                MATERIALDESCRIPTION(2, "MATL_DESC") = "垫片 " & rng.Offset(0, -1)
                
                ElseIf InStr(rng.Offset(0, -2), "BAR") <> 0 Then
                MATERIALDESCRIPTION(2, "MATL_DESC") = "棒 " & rng.Offset(0, -1)
                
                ElseIf InStr(rng.Offset(0, -2), "DRAWING PACKAGE") <> 0 Then
                MATERIALDESCRIPTION(2, "MATL_DESC") = "图纸包 " & rng.Offset(0, -1)
                
                ElseIf InStr(rng.Offset(0, -2), "RING") <> 0 Then
                MATERIALDESCRIPTION(2, "MATL_DESC") = "环 " & rng.Offset(0, -1)
                
                Else: MATERIALDESCRIPTION(2, "MATL_DESC") = "零件" & rng.Offset(0, -1)
                
            End If
            
'            MATERIALDESCRIPTION(1, "DEL_FLAG") = "X"
'************MATERIALDESCRIPTION table data end
'
''************CLIENTDATA start
'            fm2.Exports("CLIENTDATAX")("MATL_GROUP") = "X"
'            fm2.Exports("CLIENTDATAX")("BASE_UOM") = "X"
'            fm2.Exports("CLIENTDATAX")("NO_SHEETS") = "X"
'            '***** add 7 rows
'            fm2.Exports("CLIENTDATAX")("DSN_OFFICE") = "X"
'            fm2.Exports("CLIENTDATAX")("PUR_VALKEY") = "X"
'            fm2.Exports("CLIENTDATAX")("NET_WEIGHT") = "X"
'            fm2.Exports("CLIENTDATAX")("UNIT_OF_WT") = "X"
'            fm2.Exports("CLIENTDATAX")("DIVISION") = "X"
'            fm2.Exports("CLIENTDATAX")("QTY_GR_GI") = "X"
'            fm2.Exports("CLIENTDATAX")("LABEL_TYPE") = "X"
'            '***** add 7 rows
'            fm2.Exports("CLIENTDATAX")("LABEL_FORM") = "X"
'            fm2.Exports("CLIENTDATAX")("ALLOWED_WT") = "X"
'            fm2.Exports("CLIENTDATAX")("ALLWD_VOL") = "X"
'            fm2.Exports("CLIENTDATAX")("WT_TOL_LT") = "X"
'            fm2.Exports("CLIENTDATAX")("VOL_TOL_LT") = "X"
'            fm2.Exports("CLIENTDATAX")("VAR_ORD_UN") = "X"
'            fm2.Exports("CLIENTDATAX")("FILL_LEVEL") = "X"
'            fm2.Exports("CLIENTDATAX")("STACK_FACT") = "X"
'            fm2.Exports("CLIENTDATAX")("MAT_GRP_SM") = "X"
'            fm2.Exports("CLIENTDATAX")("MINREMLIFE") = "X"
'            fm2.Exports("CLIENTDATAX")("SHELF_LIFE") = "X"
'            fm2.Exports("CLIENTDATAX")("STOR_PCT") = "X"
'            fm2.Exports("CLIENTDATAX")("PVALIDFROM") = "X"
'            fm2.Exports("CLIENTDATAX")("SVALIDFROM") = "X"
'            fm2.Exports("CLIENTDATAX")("SLED_BBD") = "X"
'            ' above is the X data , below is real data
'            fm2.Exports("CLIENTDATA")("MATL_GROUP") = MATERIAL_GROUP
'            fm2.Exports("CLIENTDATA")("BASE_UOM") = BASE_UOM
'            fm2.Exports("CLIENTDATA")("NO_SHEETS") = "000"
'             '***** add 7 rows
'            fm2.Exports("CLIENTDATA")("DSN_OFFICE") = "182"
'            fm2.Exports("CLIENTDATA")("PUR_VALKEY") = "PROD"
'            fm2.Exports("CLIENTDATA")("NET_WEIGHT") = "0.000"
'            fm2.Exports("CLIENTDATA")("UNIT_OF_WT") = "KG"
'            fm2.Exports("CLIENTDATA")("DIVISION") = "BO"
'            fm2.Exports("CLIENTDATA")("QTY_GR_GI") = "0.000"
'            fm2.Exports("CLIENTDATA")("LABEL_TYPE") = "M7"
'            '***** add 7 rows
'            fm2.Exports("CLIENTDATA")("LABEL_FORM") = "E1"
'            fm2.Exports("CLIENTDATA")("ALLOWED_WT") = "0.000"
'            fm2.Exports("CLIENTDATA")("ALLWD_VOL") = "0.000"
'            fm2.Exports("CLIENTDATA")("WT_TOL_LT") = "0.0"
'            fm2.Exports("CLIENTDATA")("VOL_TOL_LT") = "0.0"
'            fm2.Exports("CLIENTDATA")("VAR_ORD_UN") = "1"
'            fm2.Exports("CLIENTDATA")("FILL_LEVEL") = "0"
'            fm2.Exports("CLIENTDATA")("STACK_FACT") = "0"
'            fm2.Exports("CLIENTDATA")("MAT_GRP_SM") = "0001"
'            fm2.Exports("CLIENTDATA")("MINREMLIFE") = "0"
'            fm2.Exports("CLIENTDATA")("SHELF_LIFE") = "0"
'            fm2.Exports("CLIENTDATA")("STOR_PCT") = "0"
'            fm2.Exports("CLIENTDATA")("PVALIDFROM") = "00000000"
'            fm2.Exports("CLIENTDATA")("SVALIDFROM") = "00000000"
'            fm2.Exports("CLIENTDATA")("SLED_BBD") = "B"
''************CLIENTDATA end
'
''************PLANTDATA start
'            fm2.Exports("PLANTDATAX")("PLANT") = PLANT ' if pass X , doesnot work , dont know why
'            fm2.Exports("PLANTDATAX")("MRP_TYPE") = "X"
'            fm2.Exports("PLANTDATAX")("MRP_CTRLER") = "X"
'            fm2.Exports("PLANTDATAX")("PLND_DELRY") = "X"
'            fm2.Exports("PLANTDATAX")("GR_PR_TIME") = "X"
'            fm2.Exports("PLANTDATAX")("PERIOD_IND") = "X"
'            fm2.Exports("PLANTDATAX")("ASSY_SCRAP") = "X"
'            fm2.Exports("PLANTDATAX")("LOTSIZEKEY") = "X"
'            fm2.Exports("PLANTDATAX")("PROC_TYPE") = "X"
'            fm2.Exports("PLANTDATAX")("SPPROCTYPE") = "X"
'            fm2.Exports("PLANTDATAX")("REORDER_PT") = "X"
'            fm2.Exports("PLANTDATAX")("SAFETY_STK") = "X"
'            fm2.Exports("PLANTDATAX")("MINLOTSIZE") = "X"
'            fm2.Exports("PLANTDATAX")("MAXLOTSIZE") = "X"
'            fm2.Exports("PLANTDATAX")("FIXED_LOT") = "X"
'            fm2.Exports("PLANTDATAX")("ROUND_VAL") = "X"
'            fm2.Exports("PLANTDATAX")("MAX_STOCK") = "X"
'            fm2.Exports("PLANTDATAX")("EFF_O_DAY") = "X"
'            fm2.Exports("PLANTDATAX")("SM_KEY") = "X"
'            fm2.Exports("PLANTDATAX")("PROC_TIME") = "X"
'            fm2.Exports("PLANTDATAX")("SETUPTIME") = "X"
'            fm2.Exports("PLANTDATAX")("INTEROP") = "X"
'            fm2.Exports("PLANTDATAX")("BASE_QTY") = "X"
'            fm2.Exports("PLANTDATAX")("INHSEPRODT") = "X"
'            fm2.Exports("PLANTDATAX")("STGEPERIOD") = "X"
'            fm2.Exports("PLANTDATAX")("OVER_TOL") = "X"
'            fm2.Exports("PLANTDATAX")("UNDER_TOL") = "X"
'            fm2.Exports("PLANTDATAX")("REPLENTIME") = "X"
'            fm2.Exports("PLANTDATAX")("SERV_LEVEL") = "X"
'            fm2.Exports("PLANTDATAX")("AVAILCHECK") = "X"
'            fm2.Exports("PLANTDATAX")("SETUP_TIME") = "X"
'            fm2.Exports("PLANTDATAX")("BASE_QTY_PLAN") = "X"
'            fm2.Exports("PLANTDATAX")("SHIP_PROC_TIME") = "X"
'            fm2.Exports("PLANTDATAX")("AUTO_P_ORD") = "X"
'            fm2.Exports("PLANTDATAX")("PL_TI_FNCE") = "X"
'            fm2.Exports("PLANTDATAX")("BWD_CONS") = "X"
'            fm2.Exports("PLANTDATAX")("FWD_CONS") = "X"
'            fm2.Exports("PLANTDATAX")("ISS_ST_LOC") = "X"
'            fm2.Exports("PLANTDATAX")("COMP_SCRAP") = "X"
'            fm2.Exports("PLANTDATAX")("CYCLE_TIME") = "X"
'            fm2.Exports("PLANTDATAX")("D_TO_REF_M") = "X"
'            fm2.Exports("PLANTDATAX")("MULT_REF_M") = "X"
'            fm2.Exports("PLANTDATAX")("EX_CERT_DT") = "X"
'            fm2.Exports("PLANTDATAX")("INSP_INT") = "X"
'            fm2.Exports("PLANTDATAX")("SLOC_EXPRC") = "X"
'            fm2.Exports("PLANTDATAX")("SAFETYTIME") = "X"
'            fm2.Exports("PLANTDATAX")("PVALIDFROM") = "X"
'            fm2.Exports("PLANTDATAX")("DEPLOY_HORIZ") = "X"
'            fm2.Exports("PLANTDATAX")("MIN_SAFETY_STK") = "X"
''           ' above is the X data , below is real data
'            fm2.Exports("PLANTDATA")("PLANT") = PLANT
'            fm2.Exports("PLANTDATA")("MRP_TYPE") = "PD"
'            fm2.Exports("PLANTDATA")("MRP_CTRLER") = MRP_CTRLER
''            fm2.Exports("PLANTDATA")("PLND_DELRY") = "15"   ' leadtime , input if necessary
'            fm2.Exports("PLANTDATA")("GR_PR_TIME") = "0"
'            fm2.Exports("PLANTDATA")("PERIOD_IND") = "M"
'            fm2.Exports("PLANTDATA")("ASSY_SCRAP") = "0.00"
'            fm2.Exports("PLANTDATA")("LOTSIZEKEY") = LOTSIZEKEY
'            fm2.Exports("PLANTDATA")("PROC_TYPE") = PROC_TYPE 'default F
'            fm2.Exports("PLANTDATA")("SPPROCTYPE") = SPPROCTYPE 'default blank
''            fm2.Exports("PLANTDATA")("REORDER_PT") = "0.000" ' reorder point  , input if necessary
''            fm2.Exports("PLANTDATA")("SAFETY_STK") = "0.500" ' safty point , input if necessary
''            fm2.Exports("PLANTDATA")("MINLOTSIZE") = "6.600" ' MINLOTSIZE , input if necessary
''            fm2.Exports("PLANTDATA")("MAXLOTSIZE") = "0.000" ' MAXLOTSIZE , input if necessary
''            fm2.Exports("PLANTDATA")("FIXED_LOT") = "0.000" ' FIXED_LOT , input if necessary
''            fm2.Exports("PLANTDATA")("ROUND_VAL") = "6.600" ' ROUND_VAL  , input if necessary
'            fm2.Exports("PLANTDATA")("MAX_STOCK") = "0.000"
'            fm2.Exports("PLANTDATA")("EFF_O_DAY") = "00000000"
'            fm2.Exports("PLANTDATA")("SM_KEY") = "000"
'            fm2.Exports("PLANTDATA")("PROC_TIME") = "0.00"
'            fm2.Exports("PLANTDATA")("SETUPTIME") = "0.00"
'            fm2.Exports("PLANTDATA")("INTEROP") = "0.00"
'            fm2.Exports("PLANTDATA")("BASE_QTY") = "0.000"
'            fm2.Exports("PLANTDATA")("INHSEPRODT") = "0"
'            fm2.Exports("PLANTDATA")("STGEPERIOD") = "0"
'            fm2.Exports("PLANTDATA")("OVER_TOL") = "0.0"
'            fm2.Exports("PLANTDATA")("UNDER_TOL") = "0.0"
'            fm2.Exports("PLANTDATA")("REPLENTIME") = "0"
'            fm2.Exports("PLANTDATA")("SERV_LEVEL") = "0.0"
'            fm2.Exports("PLANTDATA")("AVAILCHECK") = "02"
'            fm2.Exports("PLANTDATA")("SETUP_TIME") = "0.00"
'            fm2.Exports("PLANTDATA")("BASE_QTY_PLAN") = "0.000"
'            fm2.Exports("PLANTDATA")("SHIP_PROC_TIME") = "0.00"
'            fm2.Exports("PLANTDATA")("AUTO_P_ORD") = "X"
'            fm2.Exports("PLANTDATA")("PL_TI_FNCE") = "000"
'            fm2.Exports("PLANTDATA")("BWD_CONS") = "000"
'            fm2.Exports("PLANTDATA")("FWD_CONS") = "000"
'            fm2.Exports("PLANTDATA")("ISS_ST_LOC") = ISS_ST_LOC
'            fm2.Exports("PLANTDATA")("COMP_SCRAP") = "0.00"
'            fm2.Exports("PLANTDATA")("CYCLE_TIME") = "0"
'            fm2.Exports("PLANTDATA")("D_TO_REF_M") = "00000000"
'            fm2.Exports("PLANTDATA")("MULT_REF_M") = "0.00"
'            fm2.Exports("PLANTDATA")("EX_CERT_DT") = "00000000"
'            fm2.Exports("PLANTDATA")("INSP_INT") = "0"
'            fm2.Exports("PLANTDATA")("SLOC_EXPRC") = SLOC_EXPRC
'            fm2.Exports("PLANTDATA")("SAFETYTIME") = "00"
'            fm2.Exports("PLANTDATA")("PVALIDFROM") = "00000000"
'            fm2.Exports("PLANTDATA")("DEPLOY_HORIZ") = "0"
'            fm2.Exports("PLANTDATA")("MIN_SAFETY_STK") = "0.000"
''************PLANTDATA end
'
''************VALUATIONDATA start
'            fm2.Exports("VALUATIONDATAX")("VAL_AREA") = PLANT ' if pass X , doesnot work , dont know why
'            fm2.Exports("VALUATIONDATAX")("VAL_CLASS") = "X"
'            ' above is the X data , below is real data
'            fm2.Exports("VALUATIONDATA")("VAL_AREA") = PLANT
'            fm2.Exports("VALUATIONDATA")("VAL_CLASS") = "3001"
''************VALUATIONDATA end
'
''************SALESDATA start
'            fm2.Exports("SALESDATAX")("SALES_ORG") = "3003"
'            fm2.Exports("SALESDATAX")("DISTR_CHAN") = "CL"
'
'            fm2.Exports("SALESDATAX")("DELYG_PLNT") = "X"
'            fm2.Exports("SALESDATAX")("ACCT_ASSGT") = "X"
'
'            ' above is the X data , below is real data
'            fm2.Exports("SALESDATA")("SALES_ORG") = "3003"
'            fm2.Exports("SALESDATA")("DISTR_CHAN") = "CL"
'
'            fm2.Exports("SALESDATA")("DELYG_PLNT") = "1003"
'            fm2.Exports("SALESDATA")("ACCT_ASSGT") = "00"
''************SALESDATA end
'
''************TAXCLASSIFICATIONS table data start
'            TAXCLASSIFICATIONS.FreeTable
'            TAXCLASSIFICATIONS.AppendRow
'            TAXCLASSIFICATIONS(1, "DEPCOUNTRY_ISO") = "CN"
'            TAXCLASSIFICATIONS(1, "TAX_TYPE_1") = "MWST"
'            TAXCLASSIFICATIONS(1, "TAXCLASS_1") = "1"
''************TAXCLASSIFICATIONS table data end
            
            fm2.Call
                If fm2.Exception <> "" Then
                Debug.Print fm2.Exception
                Exit Sub
                End If
            If fm2.Imports("RETURN")(1) = "E" Then
            Debug.Print fm2.Imports("RETURN")(1)
            Debug.Print fm2.Imports("RETURN")(2)
            Debug.Print fm2.Imports("RETURN")(3)
            Debug.Print fm2.Imports("RETURN")(4)
'            Debug.Print fm2.Imports("RETURN")(5)
'            Debug.Print fm2.Imports("RETURN")(6)
'            Debug.Print fm2.Imports("RETURN")(7)
'            Debug.Print fm2.Imports("RETURN")(8)
'            Debug.Print fm2.Imports("RETURN")(9)
'            Debug.Print fm2.Imports("RETURN")(10)
'            Debug.Print fm2.Imports("RETURN")(11)
'            Debug.Print fm2.Imports("RETURN")(12)
'            Debug.Print fm2.Imports("RETURN")(13)
'            Debug.Print fm2.Imports("RETURN")(14)
            Exit Sub
            Else: rng.Offset(0, 0).Interior.Color = vbCyan
            End If
        End If
    Next rng
'*************fm2 call end
MsgBox "done"

    
    Call LogoffSAP
End Sub


