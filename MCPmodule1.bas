Attribute VB_Name = "Module1"
Option Explicit
    
'change log in Sheets("log")

Const ver As String = "3.4 (10-25-2026)"
'update customer classification, segmentation and channel (phase 2)
'generate delta customer master

Public Calc_Status As Long
Public RD_code As String
Public old_codes As Variant
Public new_codes As Variant
Public loc_codes As Variant
Public rtm_mod As Variant
Public zip_ref As Variant

'error counters
Public without_cml As Long
Public without_mcp As Long
Public inactive_cml As Long
Public different_dsp As Long
Public different_rtm As Long
Public no_address As Long
Public z_dup As Long

Public dup_mcp() As Variant
Public z As Long
Public y As Long
Public dict As Object
Public mcp_miss As Object
Public mcp_missed() As Variant
Public miss_count As Integer
Public cust_name As Variant
Public cust_code As Variant
Public seg_name As Variant
Public seg_code As Variant
Public ch_code As Variant
Public ch_name As Variant
Public mod_names As Variant
Public serv_model As Variant
Public errors_found As String
Public duplicated As Object
Public duplicates As Object
    
Sub Consolidate_MCP()
'fills up MCP and Summary sheets from MCP and CML user input

    Dim mcp_files As Variant, mcp_file As Variant, mcp_data As Variant
    Dim cml_files As Variant, cml_file As Variant, cml_data As Variant
    Dim file As Variant
    Dim mcp_wkbk As Workbook, cml_wkbk As Workbook
    Dim this_wk As Workbook, this_ws As Worksheet
    Dim rng As Range
    Dim LstRow As Long
    Dim startTime As Single, endTime As Single
    Dim msg As String
    Dim totalTime As Single
    Dim hdr As Variant
    Dim i As Long, j As Long
    Dim dsp_summary() As Variant
    Dim dsp As Object
    Dim lo As Object
    Dim col As Object
    Dim wkday As Variant
    Dim m As Object
    Dim key As Variant
    Dim oItem As Variant
    
    startTime = Timer
    
    SwitchOff (True)
    On Error GoTo Err_handler
    
    Erase dup_mcp
    Erase mcp_missed
    
    Set dsp = CreateObject("Scripting.Dictionary")
    Set mcp_miss = CreateObject("Scripting.Dictionary")
    Set duplicates = CreateObject("Scripting.Dictionary")
    Set duplicated = CreateObject("Scripting.Dictionary")
    miss_count = 0
    
    errors_found = ""
    
    different_rtm = 0
    different_dsp = 0
    inactive_cml = 0
    without_cml = 0
    without_mcp = 0
    z_dup = 0
    RD_code = ""
    
    y = 0
    z = 0
    
    Set this_wk = Application.ThisWorkbook
    
    mod_names = ThisWorkbook.Worksheets("ref_service_model").Range("serv_mod_name").value
    
    endTime = Timer
    totalTime = totalTime + (endTime - startTime)
    msg = "Initializing time: " & vbTab & vbTab & Round(endTime - startTime, 4) & " sec."
    
    startTime = Timer
    
    '1. Select MCP and CML files
    MsgBox "Add MCP template/s", vbOKOnly, Title:="MCP Consolidator v" & ver
    mcp_files = select_file(True)
    
    If IsArrayAllocated(mcp_files) Then
        If MsgBox("Add raw CML?", vbYesNo, Title:=ver) = vbYes Then
            cml_files = select_file(True)
        End If
    End If
    
    endTime = Timer
    totalTime = totalTime + (endTime - startTime)
    msg = msg & vbCrLf & "Asking user for CML: " & vbTab & Round(endTime - startTime, 4) & " sec."
    
    '2. Process MCP
    If Not IsArray(mcp_files) Then
        GoTo Exit_sub
    Else
        For Each file In mcp_files
                           
            z = z + 1
            startTime = Timer
            
            mcp_file = CStr(file)
            Set mcp_wkbk = open_workbook(mcp_file)
            
            endTime = Timer
            totalTime = totalTime + (endTime - startTime)
            msg = msg & vbCrLf & "Opening the MCP #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."
            
            startTime = Timer
            
            mcp_data = iterate_worksheets(mcp_wkbk, mcp_data)
            mcp_wkbk.Close
            
            endTime = Timer
            totalTime = totalTime + (endTime - startTime)
            msg = msg & vbCrLf & "Getting the MCP #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."
            
        Next
    
    End If
    
    '4. Close workbook to free up some memory
    Set mcp_wkbk = Nothing
    If Not IsArrayAllocated(mcp_data) Then
        MsgBox "No data to import."
        GoTo Exit_sub
    End If
    
    '5. Process CML workbook
    If IsArrayAllocated(cml_files) = False Then
        GoTo Paste_Values
    Else
        z = 0
        For Each file In cml_files
                           
            z = z + 1
            startTime = Timer
            
            cml_file = CStr(file)
            Set cml_wkbk = open_workbook(cml_file)
            
            endTime = Timer
            totalTime = totalTime + (endTime - startTime)
            msg = msg & vbCrLf & "Opening the CML #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."
            
            startTime = Timer
            
            cml_data = iterate_cml(cml_wkbk, cml_data)
            cml_wkbk.Close
            
            endTime = Timer
            totalTime = totalTime + (endTime - startTime)
            msg = msg & vbCrLf & "Getting the CML #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."

        Next
    
    End If
    
    Set cml_wkbk = Nothing
    
    '6. Process consolidation of MCP and CML data
    If IsArrayAllocated(cml_data) = True Then
        
        startTime = Timer
        
        mcp_data = consolidate(mcp_data, cml_data)
        
        endTime = Timer
        totalTime = totalTime + (endTime - startTime)
        msg = msg & vbCrLf & "Consolidating data: " & vbTab & Round(endTime - startTime, 4) & " sec."
        
        startTime = Timer
        errors_found = check_error(mcp_data)
        endTime = Timer
        totalTime = totalTime + (endTime - startTime)
        msg = msg & vbCrLf & "Checking for errors: " & vbTab & Round(endTime - startTime, 4) & " sec."
    
    End If
    
    
Paste_Values:
    
    '7. Create duplicates file
    
    startTime = Timer
    
    If IsArrayAllocated(mcp_missed) Then
        hdr = Array("Distributor_Code", "Customer_Code", "Customer_Name", _
            "Salesman_Code", "Salesman_Name", "Service_Model_Name", _
            "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", _
            "WK1", "WK2", "WK3", "WK4", "SheetName/FileName")
        
        ReDim d(1 To UBound(mcp_missed, 2) + 1, 1 To UBound(mcp_missed, 1))
        
        For i = 1 To UBound(mcp_missed, 2) + 1
            For j = 1 To UBound(mcp_missed, 1)
                 If i = 1 Then
                    d(i, j) = hdr(j - 1)
                Else
                    d(i, j) = mcp_missed(j, i - 1)
                End If
            Next
        Next
        
        AddNew d, "NO_MCP_WEEK_FREQUENCY_" & d(2, 1), "noWeekSchedule"
        
        If errors_found = "" Then
            errors_found = "Found " & vbCrLf & _
                            error_semantics(mcp_miss.count, "have no week schedule in MCP.")
        End If
        
    End If
    
    If IsArrayAllocated(dup_mcp) Then
        ReDim d(LBound(dup_mcp, 2) To UBound(dup_mcp, 2) + 1, LBound(dup_mcp, 1) To UBound(dup_mcp, 1))
        hdr = Array("Distributor_Code", "Customer_Code", "Customer_Name", _
            "Salesman_Code", "SalesmanName", "Service_Model_Name", _
            "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", _
            "WK1", "WK2", "WK3", "WK4", "SheetName/FileName")
        
        For i = 1 To UBound(dup_mcp, 2) + 1
            For j = 1 To UBound(dup_mcp, 1)
                 If i = 1 Then
                    d(i, j) = hdr(j - 1)
                Else
                    d(i, j) = dup_mcp(j, i - 1)
                End If
            Next
        Next
        
        AddNew d, "DUPLICATES_IN_MCP_" & d(3, 1), "duplicates"
        
        z_dup = UBound(dup_mcp, 2)
        
        If errors_found = "" Then
            errors_found = "Found " & vbCrLf & _
                error_semantics(z_dup, "are duplicates.")
        Else
            errors_found = errors_found & error_semantics(z_dup, "are duplicates.")
        End If
    ElseIf errors_found = "" Then
        errors_found = "Found no error." & vbCrLf
    End If
    
    endTime = Timer
    totalTime = totalTime + (endTime - startTime)
    msg = msg & vbCrLf & "Checking for duplicates: " & vbTab & Round(endTime - startTime, 4) & " sec."

    '8. Paste values into template
    
    startTime = Timer

    With this_wk
    
        Set this_ws = .Worksheets("MCP")
        
        
        If sheetProtect(this_ws, 0) = False Then
            
            GoTo Err_handler
        
        Else
            
            With this_ws
                
                'should remove existing data in consolidator template
                If .ListObjects("tblMCP").ListRows.count > 0 Then
                    .ListObjects("tblMCP").DataBodyRange.Delete
                End If
                
                .ListObjects("tblMCP").Resize .Range(.Cells(5, 1), .Cells(UBound(mcp_data, 1) + 5, 30))
                
                Set rng = .Range(.Cells(6, 1), .Cells(UBound(mcp_data, 1) + 5, 30)) '.Range("tblMCP") '
                rng.FormulaR1C1 = mcp_data
                
                'write the version number
                .Range("A2").value = "Version: " & ver
                                
            End With
            
            sheetProtect this_ws, 1
            'xEnableOutlining this_ws

        End If
        
        For i = LBound(mcp_data, 1) To UBound(mcp_data, 1)
            If mcp_data(i, 11) <> "" Then
                If Not dsp.Exists(mcp_data(i, 11)) Then
                    dsp.Add mcp_data(i, 11), 1
                End If
            End If
        Next
        
        Set this_ws = .Worksheets("Summary")

        If sheetProtect(this_ws, 0) = False Then

                GoTo Err_handler

        Else

            With this_ws

                .Activate

                'should remove existing data in summary sheet
                If .ListObjects("tbl_summary").ListRows.count > 0 Then
                    .ListObjects("tbl_summary").DataBodyRange.Delete
                End If
                
                If dsp.count < 3 Then
                    i = 17
                Else
                    i = dsp.count + 11
                End If
                
                .ListObjects("tbl_summary").Resize .Range(.Cells(10, 1), .Cells(i, 33))
                    
                With .ListObjects("tbl_summary")
                    .ListColumns("Distributor_Code").DataBodyRange.Cells(1).FormulaR1C1 = _
                                "=IF([@[Salesman_Code]]="""","""",INDEX(tblMCP[Distributor_Code],MATCH([Salesman_Code],tblMCP[Salesman_Code],0),1))"

                                '"=IF([@[Salesman_Code]]="""","""", IF(LEFT([@[Salesman_Code]],3)=""TFO"", ""TFM"", IF(LEFT([@[Salesman_Code]],3)=""SME"", ""SMEI"", IF(LEFT([@[Salesman_Code]],3)=""SFR"",""SFOR"", IF(LEFT([@[Salesman_Code]],3)=""DMN"", ""DMNV"", LEFT([@[Salesman_Code]],3))))))"
                    .ListColumns("Salesman_Code").DataBodyRange.Cells(1).FormulaArray = "=IFERROR(INDEX(tblMCP[Salesman_Code],MATCH(0,COUNTIF(R10C2:R[-1]C2,tblMCP[Salesman_Code]),0)),"""")"
                    .ListColumns("Salesman_Name").DataBodyRange.Cells(1).FormulaR1C1 = "=IFERROR(VLOOKUP([@[Salesman_Code]],tblMCP[[Salesman_Code]:[Salesman_Name]],2,FALSE),"""")"
                    .ListColumns("Service Model").DataBodyRange.Cells(1).FormulaR1C1 = "=IFERROR(VLOOKUP(VLOOKUP([@[Salesman_Code]],tblMCP[[Salesman_Code]:[Service_Model_Name]],3,FALSE),serv_model[[#All],[Service Model Name]:[Service Model Code]],2,FALSE),"""")"

                    wkday = Array("MON", "TUE", "WED", "THU", "FRI", "SAT")

                    For i = 1 To 4
                        For j = LBound(wkday) To UBound(wkday)
                            .ListColumns(wkday(j) & i).DataBodyRange.Cells(1).FormulaR1C1 = "=IF([@[Salesman_Code]]="""","""",COUNTIFS(tblMCP[Salesman_Code],[@[Salesman_Code]],tblMCP[" & wkday(j) & "],1,tblMCP[WK" & i & "],1))"
                        Next

                        'week summary
                        .ListColumns("WK" & i).DataBodyRange.Cells(1).FormulaR1C1 = "=IF([@[Salesman_Code]]="""","""",SUM(tbl_summary[@[MON" & i & "]:[SAT" & i & "]]))"
                    Next

                    .ListColumns("TOTAL").DataBodyRange.Cells(1).FormulaR1C1 = "=IF([@[Salesman_Code]]="""","""",SUM([@WK1],[@WK2],[@WK3],[@WK4]))"
                    
                    .DataBodyRange.FillDown
                    
                End With
                
                
                If .ListObjects("tbl_UBA").ListRows.count > 0 Then
                    .ListObjects("tbl_UBA").DataBodyRange.Delete
                End If
                
                If dsp.count < 3 Then
                    i = 17
                Else
                    i = dsp.count + 11
                End If
                
                .ListObjects("tbl_UBA").Resize .Range(.Cells(10, 35), .Cells(i, 38))
                
                With .ListObjects("tbl_UBA")

                    .ListColumns("CML").DataBodyRange.Cells(1).FormulaR1C1 = "=IF(tbl_summary[@[Salesman_Code]]>"":"",COUNTIF(tblMCP[Salesman_Code],tbl_summary[@[Salesman_Code]]),"""")"
                    .ListColumns("Target").DataBodyRange.Cells(1).FormulaR1C1 = "=IFERROR(IF(OR(tbl_summary[@[Service Model]]=""EXT"",tbl_summary[@[Service Model]]=""BTDT""),IF(tblFreq[@F2]>0,IF(tblFreq[@F4]>0,360,480), 240),IF(HLOOKUP(tbl_summary[@[Service Model]],uba_target,2,FALSE)=0,[@CML],HLOOKUP(tbl_summary[@[Service Model]],uba_target,2,FALSE))),"""")"
                    .ListColumns("% Index").DataBodyRange.Cells(1).FormulaR1C1 = "=IFERROR([@CML]/[@Target],"""")"
                    .ListColumns("Gap").DataBodyRange.Cells(1).FormulaR1C1 = "=IFERROR(IF([@Target]-[@CML]<0,0,[@Target]-[@CML]),"""")"
                    
                    .DataBodyRange.FillDown
                End With

                If .ListObjects("tblFreq").ListRows.count > 0 Then
                    .ListObjects("tblFreq").DataBodyRange.Delete
                End If
                
                If dsp.count < 3 Then
                    i = 17
                Else
                    i = dsp.count + 11
                End If
                
                .ListObjects("tblFreq").Resize .Range(.Cells(10, 41), .Cells(i, 44))
                
                With .ListObjects("tblFreq")

                    .ListColumns("F4").DataBodyRange.Cells(1).FormulaR1C1 = "=COUNTIFS(tblMCP[Frequency],tblFreq[[#Headers],[F4]],tblMCP[Salesman_Code],tbl_summary[@[Salesman_Code]]) + + COUNTIFS(tblMCP[Frequency],""F8"",tblMCP[Salesman_Code],tbl_summary[@[Salesman_Code]]) + COUNTIFS(tblMCP[Frequency],""F12"",tblMCP[Salesman_Code],tbl_summary[@[Salesman_Code]])"
                    .ListColumns("F3").DataBodyRange.Cells(1).FormulaR1C1 = "=COUNTIFS(tblMCP[Frequency],tblFreq[[#Headers],[F3]],tblMCP[Salesman_Code],tbl_summary[@[Salesman_Code]])"
                    .ListColumns("F2").DataBodyRange.Cells(1).FormulaR1C1 = "=COUNTIFS(tblMCP[Frequency],tblFreq[[#Headers],[F2]],tblMCP[Salesman_Code],tbl_summary[@[Salesman_Code]])"
                    .ListColumns("F1").DataBodyRange.Cells(1).FormulaR1C1 = "=COUNTIFS(tblMCP[Frequency],tblFreq[[#Headers],[F1]],tblMCP[Salesman_Code],tbl_summary[@[Salesman_Code]])"
                    
                    .DataBodyRange.FillDown
                End With
            
            xEnableOutlining this_ws
            
            End With

        End If

    End With
    
    endTime = Timer
    totalTime = totalTime + (endTime - startTime)
    msg = msg & vbCrLf & "Writing data to table: " & vbTab & Round(endTime - startTime, 4) & " sec."
    
'    If mcp_miss.count > 0 Then
'        errors_found = errors_found & vbCrLf & mcp_miss.count & " no schedule in MCP."
'    End If
    
    MsgBox "Timer " & vbCrLf & msg & vbCrLf & vbCrLf & "Total Time: " & Round(totalTime, 0) & " sec"
    MsgBox "MCP consolidated successfully!" & vbCrLf & vbCrLf & errors_found, , Title:=ver
    
Exit_sub:
    Set this_wk = Nothing
    Set this_ws = Nothing
    xEnableOutlining this_ws
    mcp_data = ""
    SwitchOff False   'turn these features back on
    Exit Sub
Err_handler:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Kindly email SOCD."
    GoTo Exit_sub
End Sub

Function iterate_worksheets(wb As Workbook, mcp As Variant) As Variant
'returns data for sheet MCP

    Dim ws As Worksheet
    Dim arr As Variant
    Dim temp() As Variant
    Dim a As Long
    Dim data_row As Long
    Dim data_col As Long
    Dim k As Long
    Dim seq As Long
    Dim x As Long
    Dim startData As Long
    Dim i As Long
    Dim WeeksCover As String
    Dim Route_Name() As Variant
    Dim route_day() As String
    Dim Route() As Long
    Dim Week1() As Long
    Dim Week2() As Long
    Dim Week3() As Long
    Dim Week4() As Long
    Dim Address() As Long
    Dim Segment() As Long
    Dim n As Long
    Dim mcp_ws As Worksheets
    Dim mcp_seq As Boolean
    Dim mcp_data As clsMCP
    Dim dup As String
    Dim dup2 As String
    Dim mcp_m As clsMCP
    Dim cust_code As String
    
    On Error GoTo mcp_error
    
    Set mcp_m = New clsMCP
    
    If VarType(mcp) = (vbVariant + vbArray) Then
        ReDim temp(LBound(mcp, 2) To UBound(mcp, 2), LBound(mcp, 1) To UBound(mcp, 1))
        temp = Application.WorksheetFunction.Transpose(mcp)
        k = UBound(mcp, 1)
    Else
        k = 0
    End If
    
    For Each ws In wb.Worksheets
        
        'read sheet data and put into array
        With ws
            arr = .Range(.Cells(1, 1), .Cells(.UsedRange.Rows.count, .UsedRange.Columns.count))
        End With
        
        mcp_seq = False
        
        'Aim is to get the require entries for
            'Salesman name
            'Salesman code
            'Service model
            'Supervisor name
        'from the column headings.
        
        'check if sheet is empty or not
        If VarType(arr) = vbString Or VarType(arr) = 0 Then GoTo Next_ws
        If UBound(arr, 2) < 10 Or UBound(arr, 1) < 10 Then GoTo Next_ws
        
        Set mcp_data = New clsMCP
        
        For data_row = 1 To 10
            
            For data_col = 1 To 10   'column
                
                'same consolidator template
                If Left(arr(data_row, 1), 16) = "Distributor_Code" And _
                    Left(arr(data_row, 2), 13) = "Customer_Code" And _
                    Left(arr(data_row, 11), 13) = "Salesman_Code" Then
                    
                    iterate_worksheets = Copy_Paste(arr, data_row)
                    Exit Function
                                
                'SFA route master template
                ElseIf Left(arr(data_row, 1), 16) = "Distributor_Code" And _
                    Left(arr(data_row, 2), 10) = "Route_Code" And _
                    Left(arr(data_row, 3), 10) = "Route_Name" Then
                    
                    iterate_worksheets = Copy_Paste(arr, data_row, 1)
                    Exit Function
                
                'salesman name
                ElseIf arr(data_row, data_col) = "Salesman Name:" Or _
                    arr(data_row, data_col) = "Salesman:" Then
                    
                    mcp_data.Salesman_Name = arr(data_row, data_col + 2)
                        
                    If mcp_data.Salesman_Name = "" Then
                        
                        mcp_data.Salesman_Name = arr(data_row, data_col + 1)
                        
                    End If
                
                'salesman code
                ElseIf arr(data_row, data_col) = "Salesman Code:" Or _
                    arr(data_row, data_col) = "DSP Code:" Then
                    
                    mcp_data.salesman_code = arr(data_row, data_col + 2)
                        
                    If mcp_data.salesman_code = "" Then
                        
                        mcp_data.salesman_code = arr(data_row, data_col + 1)
                        
                        'this is required
                        If mcp_data.salesman_code = "" Then GoTo Next_ws
                        
                    End If
                    
                    If mcp_data.Distributor_Code = "" Then
                    
                        If Left(mcp_data.salesman_code, 3) = "DMN" Then
                            
                            mcp_data.Distributor_Code = "DMNV"
                        
                        ElseIf Left(mcp_data.salesman_code, 3) = "SFR" Then
                        
                            mcp_data.Distributor_Code = "SFOR"
                        
                        ElseIf Left(mcp_data.salesman_code, 3) = "SME" Then
                        
                            mcp_data.Distributor_Code = "SMEI"
                            
                        ElseIf Left(mcp_data.salesman_code, 3) = "TFO" Then
                        
                            mcp_data.Distributor_Code = "TFM"
                        
                        ElseIf Left(mcp_data.salesman_code, 3) = "VKG" Or _
                            Left(mcp_data.salesman_code, 3) = "VKA" _
                        Then
                            
                            mcp_data.Distributor_Code = get_rd_code
                            
                        ElseIf IsNumeric(Mid(mcp_data.salesman_code, 2, 1)) Then
                        
                            mcp_data.Distributor_Code = get_rd_code
                            
                        ElseIf IsNumeric(Mid(mcp_data.salesman_code, 3, 1)) Then
                        
                            If Left(mcp_data.salesman_code, 3) = "KM6" Then
                                mcp_data.Distributor_Code = Left(mcp_data.salesman_code, 3)
                            Else
                                mcp_data.Distributor_Code = get_rd_code
                            End If
                        
                        '6-19-2021
                        ElseIf Left(mcp_data.salesman_code, 3) = "EVD" Then
                        
                            Select Case Right(mcp_data.salesman_code, 2)
                                Case 20, 29, 40, 42
                                    mcp_data.Distributor_Code = "EVB"
                                Case 16 To 18, 21, 28, 30 To 32, 41, 43 To 45
                                    mcp_data.Distributor_Code = "EVN"
                                Case Else
                                    mcp_data.Distributor_Code = "EVD"
                            End Select
                            
                        Else
                        
                            mcp_data.Distributor_Code = Left(mcp_data.salesman_code, 3)
                            mcp_data.Distributor_Code = mcp_data.Distributor_Code
                            
                        End If
                    Else
                    
                        mcp_data.Distributor_Code = mcp_data.Distributor_Code
                        
                    End If
                    
                    
                'service model
                ElseIf arr(data_row, data_col) = "Service Model:" Or _
                    arr(data_row, data_col) = "Position:" Then
                    
                    mcp_data.service_model = arr(data_row, data_col + 2)
                    
                    If mcp_data.service_model = "" Then
                        mcp_data.service_model = arr(data_row, data_col + 1)
                        If mcp_data.service_model = "" Then GoTo Next_ws
                    End If
                    
                    mcp_data.service_model = check_service_model(mcp_data.service_model)
                    
                'supervisor
                ElseIf arr(data_row, data_col) = "Supervisor Name:" Or _
                    arr(data_row, data_col) = "Supervisor:" Then
                    
                    mcp_data.Supervisor_Name = arr(data_row, data_col + 2)
                    If mcp_data.Supervisor_Name = "" Then
                        mcp_data.Supervisor_Name = arr(data_row, data_col + 1)
                        
                    End If
                    
                    
                'wrong template
                ElseIf arr(data_row, data_col) = "Seq" Or _
                    arr(data_row, data_col + 1) = "Seq" Then
                        
                    mcp_seq = True
                    startData = data_row
                ElseIf data_row = 10 Then
                    If mcp_seq = False Then
                        If ws.Name = wb.Worksheets(wb.Worksheets.count).Name And IsArrayAllocated(temp) = False Then
                            Exit Function
                        Else
                            GoTo Next_ws
                        End If
                    Else
                        GoTo Next_seq
                    End If
                End If
                    
                
                
            Next data_col
        Next data_row

Next_seq:
        seq = 0
        
        For data_row = 1 To 9
            For data_col = 1 To UBound(arr, 2)
            
                'route day
                If arr(data_row, data_col) = "Day:" Then
                
                    'route day should not be empty
                    If arr(data_row, data_col + 1) <> "" Then
                    
                        seq = seq + 1
                        
                        ReDim Preserve Route_Name(1 To 2, 1 To seq)
                        ReDim Preserve Route(1 To seq)
                        ReDim Preserve Week1(1 To seq)
                        ReDim Preserve Week2(1 To seq)
                        ReDim Preserve Week3(1 To seq)
                        ReDim Preserve Week4(1 To seq)
                        ReDim Preserve Segment(1 To seq)
                        ReDim Preserve Address(1 To seq)
                        ReDim Preserve route_day(1 To seq)
                        
                        route_day(seq) = arr(data_row, data_col + 1)
                        
                        For x = 0 To 5
                            
                            If data_col + x <= UBound(arr, 2) Then
                                Select Case UCase(arr(10, data_col + x))
                                    Case Is = "WK1": Week1(seq) = data_col + x
                                    Case Is = "WK2": Week2(seq) = data_col + x
                                    Case Is = "WK3": Week3(seq) = data_col + x
                                    Case Is = "WK4": Week4(seq) = data_col + x
                                End Select
                            End If
                            
                            If data_col - x > 0 Then
                                Select Case Left(arr(10, data_col - x), 3)
                                    Case Is = "Seq": Route(seq) = data_col - x
                                    Case Is = "Cus": Segment(seq) = data_col - x
                                    Case Is = "No.": Address(seq) = data_col - x
                                End Select
                            End If
                            
                        Next
                        
                        Route_Name(1, seq) = UCase(mcp_data.salesman_code) '& "_" & Left(Route_Day(seq), 3)) 'route_code
                        Route_Name(2, seq) = UCase(mcp_data.salesman_code) '& "_" & Route_Day(seq))           'route_name
                        
                        'sidamo mcp template
                        If Week1(1) = 0 Then
                        
                            If arr(data_row + 2, data_col + 1) = "F2" Then
                                If arr(data_row + 3, data_col + 1) = "Weeks 1 & 3" _
                                    Or arr(data_row + 3, data_col + 1) = "" Then
                                        WeeksCover = "2F1"
        
                                ElseIf arr(data_row + 3, data_col + 1) = "Weeks 2 & 4" Then
                                    WeeksCover = "2F2"
        
                                End If
        
                                Route_Name(1, seq) = UCase(Route_Name(1, seq) & "_" & WeeksCover)  'route_code
                                Route_Name(2, seq) = UCase(Route_Name(2, seq) & "_" & WeeksCover)  'route_name
        
                            End If
                        
                        End If
                   
                    End If
                End If
            Next data_col
        Next data_row
                
        For data_col = LBound(Route) To UBound(Route)
            For data_row = startData To UBound(arr, 1)
            
                If Trim(CStr(arr(data_row, Route(data_col) + 1))) <> "" And _
                    arr(data_row, Route(data_col) + 1) <> "Account Code" Then
                    
                    cust_code = CStr(arr(data_row, Route(data_col) + 1))
                    
                    'check for duplicate retailer codes
                    If duplicates.Exists(cust_code) Then
                        
                        For a = 1 To UBound(temp, 2)
                            
                            'find the duplicate
                            If cust_code = temp(2, a) Then
                                
                                'check if same salesman
                                If temp(11, a) = mcp_data.salesman_code Then
                                    
                                    'add the coverage day
                                    Select Case UCase(route_day(data_col))
                                        Case "MONDAY"
                                            temp(17, a) = 1
                                        Case "TUESDAY"
                                            temp(18, a) = 1
                                        Case "WEDNESDAY"
                                            temp(19, a) = 1
                                        Case "THURSDAY"
                                            temp(20, a) = 1
                                        Case "FRIDAY"
                                            temp(21, a) = 1
                                        Case "SATURDAY"
                                            temp(22, a) = 1
                                    End Select
                                    
                                    'add the coverage week
                                    If arr(data_row, Week1(data_col)) = 1 Then temp(24, a) = 1
                                    If arr(data_row, Week2(data_col)) = 1 Then temp(25, a) = 1
                                    If arr(data_row, Week3(data_col)) = 1 Then temp(26, a) = 1
                                    If arr(data_row, Week4(data_col)) = 1 Then temp(27, a) = 1
                                   
                                Else
                                    dup = temp(1, a) & temp(2, a) & temp(11, a)
                                    dup2 = mcp_data.Distributor_Code & cust_code & mcp_data.salesman_code
                                    
                                    If duplicated.Exists(dup) = False Then
                                        
                                        duplicated.Add dup, 0
                                        
                                        y = y + 1
                                        
                                        ReDim Preserve dup_mcp(1 To 18, 1 To y)
                                        dup_mcp(1, y) = temp(1, a)
                                        dup_mcp(2, y) = temp(2, a)
                                        dup_mcp(3, y) = temp(3, a)
                                        dup_mcp(4, y) = temp(11, a)
                                        dup_mcp(5, y) = temp(12, a)
                                        dup_mcp(6, y) = temp(13, a)
                                        dup_mcp(7, y) = temp(17, a)
                                        dup_mcp(8, y) = temp(18, a)
                                        dup_mcp(9, y) = temp(19, a)
                                        dup_mcp(10, y) = temp(20, a)
                                        dup_mcp(11, y) = temp(21, a)
                                        dup_mcp(12, y) = temp(22, a)
                                        dup_mcp(13, y) = temp(23, a)
                                        dup_mcp(14, y) = temp(24, a)
                                        dup_mcp(15, y) = temp(25, a)
                                        dup_mcp(16, y) = temp(26, a)
                                        dup_mcp(17, y) = temp(27, a)
                                        dup_mcp(18, y) = "File: " & wb.Name & ", Sheet: " & ws.Name & _
                                                         ", Cell: " & Cells(data_row, Route(data_col) + 5).Address(False, False)

                                    End If
                                    
                                    If duplicated.Exists(dup2) = False Then
                                        
                                        duplicated.Add dup2, 0
                                        
                                        y = y + 1
                                        
                                        ReDim Preserve dup_mcp(1 To 18, 1 To y)
                                        dup_mcp(1, y) = mcp_data.Distributor_Code
                                        dup_mcp(2, y) = arr(data_row, Route(data_col) + 1)
                                        dup_mcp(3, y) = arr(data_row, Route(data_col) + 2)
                                        dup_mcp(4, y) = mcp_data.salesman_code
                                        dup_mcp(5, y) = mcp_data.Salesman_Name
                                        dup_mcp(6, y) = mcp_data.service_model
                                        dup_mcp(7, y) = IIf(route_day(data_col) = "Monday", 1, "")
                                        dup_mcp(8, y) = IIf(route_day(data_col) = "Tuesday", 1, "")
                                        dup_mcp(9, y) = IIf(route_day(data_col) = "Wednesday", 1, "")
                                        dup_mcp(10, y) = IIf(route_day(data_col) = "Thursday", 1, "")
                                        dup_mcp(11, y) = IIf(route_day(data_col) = "Friday", 1, "")
                                        dup_mcp(12, y) = IIf(route_day(data_col) = "Saturday", 1, "")
                                        
                                        dup_mcp(14, y) = arr(data_row, Route(data_col) + 6)
                                        dup_mcp(15, y) = arr(data_row, Route(data_col) + 7)
                                        dup_mcp(16, y) = arr(data_row, Route(data_col) + 8)
                                        dup_mcp(17, y) = arr(data_row, Route(data_col) + 9)
                                        dup_mcp(18, y) = "File: " & wb.Name & ", Sheet: " & ws.Name & _
                                                         ", Cell: " & Cells(data_row, Route(data_col) + 5).Address(False, False)
                                        
                                    End If
                                    
                                End If
                                
                                GoTo Next_row
                            
                            End If
                        Next
                    Else
                        duplicates.Add cust_code, 0 'CStr(arr(data_row, Route(data_col) + 1)), 0
                    End If
                    
                    k = k + 1

                    ReDim Preserve temp(1 To 30, 1 To k)
                    temp(1, k) = mcp_data.Distributor_Code
                    temp(2, k) = arr(data_row, Route(data_col) + 1)  'Customer Code
                    temp(3, k) = arr(data_row, Route(data_col) + 2)  'Customer Name
                    temp(3, k) = Replace(CStr(temp(3, k)), Chr(34), "")
'                    temp(4, k) = d                     'Province care of CML
'                    temp(5, k) = d                     'Town/City care of CML
'                    temp(6, k) = d                     'Barangay care of CML
                    temp(7, k) = WorksheetFunction.IfError(arr(data_row, Address(data_col)), "") 'Street address
                    temp(8, k) = "=IFERROR(VLOOKUP([@Province]&[@[Town/City]]&[@Barangay],ref_brgy!C18:C20,3,FALSE),"""")"

                    If Segment(data_col) = Empty Then
                        temp(9, k) = Empty
                    Else
                        temp(9, k) = arr(data_row, Segment(data_col))  'Segmentation code
                    End If

'                    temp(10, k) = d                    'Credit Term care of CML
                    temp(11, k) = mcp_data.salesman_code
                    temp(12, k) = mcp_data.Salesman_Name
                    temp(13, k) = mcp_data.service_model
'                    temp(14,k) = d                     'SalesSup_Code
                    temp(15, k) = mcp_data.Supervisor_Name
                    temp(16, k) = "=IF(COUNTA(tblMCP[@[Mon]:[Sun]])*COUNTA(tblMCP[@[WK1]:[WK4]])=0,"""",""F""&COUNTA(tblMCP[@[Mon]:[Sun]])*COUNTA(tblMCP[@[WK1]:[WK4]]))"
                    temp(28, k) = "=IFERROR(CHOOSE([@Day],COUNTIFS(R6C11:RC11,RC11,R6C17:RC17,1),COUNTIFS(R6C11:RC11,RC11,R6C18:RC18,1),COUNTIFS(R6C11:RC11,RC11,R6C19:RC19,1),COUNTIFS(R6C11:RC11,RC11,R6C20:RC20,1),COUNTIFS(R6C11:RC11,RC11,R6C21:RC21,1),COUNTIFS(R6C11:RC11,RC11,R6C22:RC22,1),COUNTIFS(R6C11:RC11,RC11,R6C23:RC23,1)),"""")"
                    temp(29, k) = "=IFERROR(MATCH(1,tblMCP[@[Mon]:[Sun]],0),"""")"

                    For n = 17 To 27            'reset days
                        temp(n, k) = Empty
                    Next n

                    'set Y for day
                    Select Case UCase(route_day(data_col))
                        Case "MONDAY"
                            temp(17, k) = 1
                        Case "TUESDAY"
                            temp(18, k) = 1
                        Case "WEDNESDAY"
                            temp(19, k) = 1
                        Case "THURSDAY"
                            temp(20, k) = 1
                        Case "FRIDAY"
                            temp(21, k) = 1
                        Case "SATURDAY"
                            temp(22, k) = 1
                    End Select

                    If Week1(1) = 0 Then
                        'week 1, 2, 3, 4
                        Select Case Right(Route_Name(1, data_col), 3)
                            Case "2F1"
                                temp(24, k) = 1
                                temp(26, k) = 1
                            Case "2F2"
                                temp(25, k) = 1
                                temp(27, k) = 1
                            Case Else
                                temp(24, k) = 1
                                temp(25, k) = 1
                                temp(26, k) = 1
                                temp(27, k) = 1
                        End Select

                    Else
                        temp(24, k) = arr(data_row, Week1(data_col))
                        temp(25, k) = arr(data_row, Week2(data_col))
                        temp(26, k) = arr(data_row, Week3(data_col))
                        temp(27, k) = arr(data_row, Week4(data_col))

                    End If
                    
                    If temp(24, k) = "" And temp(25, k) = "" _
                        And temp(26, k) = "" And temp(27, k) = "" Then
                        
                        miss_count = miss_count + 1
                        ReDim Preserve mcp_missed(1 To 18, 1 To miss_count)
                        
                        mcp_missed(1, miss_count) = temp(1, k) 'Distributor Code
                        mcp_missed(2, miss_count) = temp(2, k) 'Customer Code
                        mcp_missed(3, miss_count) = Replace(CStr(temp(3, k)), Chr(34), "") 'Customer Name
                        mcp_missed(4, miss_count) = temp(11, k) 'Salesman Code
                        mcp_missed(5, miss_count) = temp(12, k) 'Salesman Name
                        mcp_missed(6, miss_count) = temp(13, k) 'Service Model
                        mcp_missed(7, miss_count) = temp(17, k) 'Mon
                        mcp_missed(8, miss_count) = temp(18, k)
                        mcp_missed(9, miss_count) = temp(19, k)
                        mcp_missed(10, miss_count) = temp(20, k)
                        mcp_missed(11, miss_count) = temp(21, k)
                        mcp_missed(12, miss_count) = temp(22, k)
                        mcp_missed(13, miss_count) = temp(23, k)
                        mcp_missed(14, miss_count) = temp(24, k) 'Wk1
                        mcp_missed(15, miss_count) = temp(25, k)
                        mcp_missed(16, miss_count) = temp(26, k)
                        mcp_missed(17, miss_count) = temp(27, k)
                        mcp_missed(18, miss_count) = "File: " & wb.Name & ", Sheet: " & ws.Name & _
                                                         ", Cell: " & Cells(data_row, Route(data_col) + 5).Address(False, False)

'''                        mcp_m.Distributor_Code = temp(1, k)
                        mcp_m.Customer_Code = temp(2, k) 'UCase(temp(2, k))
'''                        mcp_m.Customer_Name = temp(3, k)
'''                        mcp_m.service_model = temp(4, k)
'''                        mcp_m.Province = temp(5, k)
'''                        mcp_m.Town = temp(6, k)
'''                        mcp_m.Street = temp(7, k)
'''                        mcp_m.Barangay = temp(8, k)
'''                        mcp_m.Customer_Class = temp(9, k)
'''                        mcp_m.Credit_Term = temp(10, k)
                        mcp_m.salesman_code = temp(11, k)
'''                        mcp_m.Salesman_Name = temp(12, k)
'''                        mcp_m.service_model = temp(13, k)
'''                        mcp_m.Supervisor_Code = temp(14, k)
'''                        mcp_m.Supervisor_Name = temp(15, k)
'''                        mcp_m.Mon = temp(17, k)
'''                        mcp_m.Tue = temp(18, k)
'''                        mcp_m.Wed = temp(19, k)
'''                        mcp_m.Thu = temp(20, k)
'''                        mcp_m.Fri = temp(21, k)
'''                        mcp_m.Sat = temp(22, k)
'''                        mcp_m.Sun = temp(23, k)
'''                        mcp_m.WK1 = IIf(temp(24, k) = "", 0, temp(24, k))
'''                        mcp_m.WK2 = IIf(temp(25, k) = "", 0, temp(25, k))
'''                        mcp_m.WK3 = IIf(temp(26, k) = "", 0, temp(26, k))
'''                        mcp_m.WK4 = IIf(temp(27, k) = "", 0, temp(27, k))
                        
                        If Not mcp_miss.Exists(mcp_m.Customer_Code & mcp_m.salesman_code) Then
                            mcp_miss.Add mcp_m.Customer_Code & mcp_m.salesman_code, 1
                        End If
                        
                    End If

                End If

Next_row:
            Next data_row
        Next data_col

Next_ws:
    Next
    
    If VarType(temp) = (vbVariant + vbArray) Then
        iterate_worksheets = Application.WorksheetFunction.Transpose(temp)
    End If
    
    Exit Function

mcp_error:
    MsgBox "Error in " & vbCrLf & _
            "File: " & wb.Name & vbCrLf & _
            "Sheet: " & ws.Name & vbCrLf & _
            "Row: " & data_row
    SwitchOff False   'turn these features back on
    
End Function

Function Copy_Paste(arr As Variant, start_row As Long, Optional template As Byte = 0) As Variant
'copy contents if the same consolidator template (old MCP template)

    Dim temp() As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim PSGC As String
    Dim freq  As String
    Dim sequ As String
    Dim dayt As String
    Dim serv As Variant
    Dim LstRow As Long
    
    'service model reference for vlookup
    serv = ThisWorkbook.Worksheets("ref_service_model").Range("serv_mod_name").value
    
    start_row = start_row + 1
    k = 1 - start_row
        
    'formula columns
    PSGC = "=IFERROR(VLOOKUP([@Province]&[@[Town/City]]&[@Barangay],ref_brgy!C18:C20,3,FALSE),"""")"
    freq = "=IF(COUNTA(tblMCP[@[Mon]:[Sun]])*COUNTA(tblMCP[@[WK1]:[WK4]])=0,"""",""F""&COUNTA(tblMCP[@[Mon]:[Sun]])*COUNTA(tblMCP[@[WK1]:[WK4]]))"
    sequ = "=IFERROR(CHOOSE([@Day],COUNTIFS(R6C11:RC11,RC11,R6C17:RC17,1),COUNTIFS(R6C11:RC11,RC11,R6C18:RC18,1),COUNTIFS(R6C11:RC11,RC11,R6C19:RC19,1),COUNTIFS(R6C11:RC11,RC11,R6C20:RC20,1),COUNTIFS(R6C11:RC11,RC11,R6C21:RC21,1),COUNTIFS(R6C11:RC11,RC11,R6C22:RC22,1),COUNTIFS(R6C11:RC11,RC11,R6C23:RC23,1)),"""")"
    dayt = "=IFERROR(MATCH(1,tblMCP[@[Mon]:[Sun]],0),"""")"
    
    'loop through contents
    For i = start_row To UBound(arr, 1)
    
        If arr(i, 2) <> "" Then
            ReDim Preserve temp(1 To 30, 1 To (k + i))
            If template = 1 Then
                temp(1, i + k) = arr(i, 1)
                temp(2, i + k) = arr(i, 4)
                temp(8, i + k) = PSGC
                temp(11, i + k) = arr(i, 5)
                temp(16, i + k) = freq
                temp(17, i + k) = IIf(WorksheetFunction.And(arr(i, 7) <> "", arr(i, 7) <> "N"), 1, "")
                temp(18, i + k) = IIf(WorksheetFunction.And(arr(i, 8) <> "", arr(i, 8) <> "N"), 1, "")
                temp(19, i + k) = IIf(WorksheetFunction.And(arr(i, 9) <> "", arr(i, 9) <> "N"), 1, "")
                temp(20, i + k) = IIf(WorksheetFunction.And(arr(i, 10) <> "", arr(i, 10) <> "N"), 1, "")
                temp(21, i + k) = IIf(WorksheetFunction.And(arr(i, 11) <> "", arr(i, 11) <> "N"), 1, "")
                temp(22, i + k) = IIf(WorksheetFunction.And(arr(i, 12) <> "", arr(i, 12) <> "N"), 1, "")
                temp(23, i + k) = IIf(WorksheetFunction.And(arr(i, 13) <> "", arr(i, 13) <> "N"), 1, "")
                temp(24, i + k) = IIf(WorksheetFunction.And(arr(i, 14) <> "", arr(i, 14) <> "N"), 1, "")
                temp(25, i + k) = IIf(WorksheetFunction.And(arr(i, 15) <> "", arr(i, 15) <> "N"), 1, "")
                temp(26, i + k) = IIf(WorksheetFunction.And(arr(i, 16) <> "", arr(i, 16) <> "N"), 1, "")
                temp(27, i + k) = IIf(WorksheetFunction.And(arr(i, 17) <> "", arr(i, 17) <> "N"), 1, "")
                temp(28, i + k) = sequ
                temp(29, i + k) = dayt
            Else
                For j = 1 To 30
                    If j = 8 Then
                        temp(j, i + k) = PSGC
                    ElseIf j = 16 Then
                        temp(j, i + k) = freq
                    ElseIf j = 28 Then
                        temp(j, i + k) = sequ
                    ElseIf j = 29 Then
                        temp(j, i + k) = dayt
                    ElseIf j = 13 Then
                        temp(j, i + k) = check_service_model(CStr(arr(i, j))) 'check if service model is updated
                    Else
                        temp(j, i + k) = arr(i, j)
                    End If
                Next
            End If
        End If
    Next
    
    'return
    Copy_Paste = WorksheetFunction.Transpose(temp)
    
End Function

Function iterate_cml(wb As Workbook, cml As Variant) As Variant
'returns single or combined CML

    Dim ws As Worksheet
    Dim arr As Variant
    Dim temp() As Variant
    Dim col_b1 As Variant
    Dim col_te As Variant
    Dim obj_b1 As Object
    Dim obj_te As Object
    Dim obj_temp As Object
    Dim b1 As Long
    Dim te As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim tem As Long
    Dim col_temp As Variant
    Dim Lst_Col As Long
    
    Set obj_b1 = CreateObject("Scripting.Dictionary")
    Set obj_te = CreateObject("Scripting.Dictionary")
    Set obj_temp = CreateObject("Scripting.Dictionary")
    
    'SAP B1 columns
    col_b1 = Array("Company Name", "Code", "Trade Name", "RTM Model Code", "RTM Model Name", _
                "Customer Class Code", "Customer Class Name", "Customer Class Group Code", _
                "Customer Class Group Name", "Channel Type Code", "Channel Type Name", _
                "Phone No.", "BIR TIN", "Area", "Owner", "Address", "Zip Code", "Province", _
                "Town", "Barangay", "Cluster", "RD Salesman Code", "RD Salesman Name", _
                "Credit Line", "Credit Terms", "Status", "Lead Account", "Tax Acct.", _
                "Date of 1st Purchase", "Freezer")
    
    'Trade Edge columns
    col_te = Array("CUSTOMER_NUMBER", "CUSTOMER_NAME", "CSTNUMBER", "TINNUMBER", "KEY_ACCOUNT", _
            "STATUS", "BRANCH_NAME", "FIRST_NAME", "LAST_NAME", "CONTACT_NO", "ADDRESS1", _
            "ADDRESS2", "CITY", "STATE_PROVINCE", "POSTAL_CODE", "CHANNEL", "CREDIT_LIMIT", _
            "CREDIT_LIMIT_DAYS", "RETAIL_ENVIRONMENT", "PARTY_CLASSIFICATION_DESCRIPTION", _
            "TAX_CODE", "CREATED_DATE", "SALES_REP_ID", "SALES_REP_NAME")
    
    'Trade Edge template
    col_temp = Array("*Customer Offline Number", "*Customer Name(Remove all special characters)", _
            "CST Number(Mark Up)", "*TIN Number", "Key Account(Channel Type)", "*Status(Always the same)", _
            "Servicing Branch Name", "*Contact Person First Name", "*Contact Person Last Name", _
            "*Contact Person Phone Number(No special Character)", "*Address Line 1", "Address Line 2(Barangay)", _
            "*City/Town", "*State(Province)", "*Postal Code(Zip Code)", "*Customer Channel(Customer Type)", _
            "*Credit Limit", "*Credit Limit (Days-Only numerical)", "*Retail Environment(Segmentation)", _
            "*Customer Class", "Tax Codes", "Sales Rep ID(Mandatory)", "Salesman Name")
            
    For i = LBound(col_b1) To UBound(col_b1)
        obj_b1.Add col_b1(i), 0
    Next
    
    For i = LBound(col_te) To UBound(col_te)
        obj_te.Add col_te(i), 0
    Next
    
    For i = LBound(col_temp) To UBound(col_temp)
        obj_temp.Add col_temp(i), 0
    Next
    
    b1 = 0
    te = 0
    tem = 0
    arr = wb.Worksheets(1).UsedRange
    
    'loop through column headers
    Lst_Col = IIf(UBound(arr, 2) > 100, 74, UBound(arr, 2))
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To Lst_Col
            If arr(i, j) <> "" Then
                
                'test column headers
                If obj_b1.Exists(arr(i, j)) Then
                    b1 = b1 + 1
                End If
                
                If obj_te.Exists(arr(i, j)) Then
                    te = te + 1
                End If
                
                If obj_temp.Exists(arr(i, j)) Then
                    tem = tem + 1
                End If
                
                '
                If j >= Lst_Col Or _
                (j = 31 And b1 > 20) Or _
                (te = UBound(col_te)) Or _
                (tem = UBound(col_temp)) Then
                
                    If VarType(cml) = (vbVariant + vbArray) Then
                    
                        If te = UBound(col_te) Or tem = UBound(col_temp) Then
                            If UBound(cml, 2) = 31 Then
                                iterate_cml = MergeArrays(cml, convert_TE_cml(arr, i))
                            Else
                                iterate_cml = MergeArrays(convert_TE_cml(arr, i), cml)
                            End If
                        ElseIf b1 = UBound(col_b1) + 1 Then
                            iterate_cml = MergeArrays(arr, cml)
                        Else
                            iterate_cml = MergeArrays(cml, arr)
                        End If
                    ElseIf te = UBound(col_te) Or tem = UBound(col_temp) Then 'te
                        iterate_cml = convert_TE_cml(arr, i)
'                    ElseIf tem = UBound(col_temp) + 1 Then
'                        iterate_cml = convert_TE_cml(arr, i) 'te
                    Else
                        iterate_cml = arr   'sap b1
                    End If
                    
                    iterate_cml = remove_special_characters(iterate_cml)
                    
                    Exit Function
                        
                End If
            Else
                Exit For
            End If
                    
        Next j
    Next i
    
End Function

Function remove_special_characters(cml) As Variant
'returns cml array stripped of special characters
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    k = 0
    'ReDim arr(LBound(cml, 1) To UBound(cml, 1), LBound(cml, 2) To UBound(cml, 2))
    
    For i = LBound(cml, 1) To UBound(cml, 1)
    
        If j = 0 Then
            For k = LBound(cml, 2) To UBound(cml, 2)
                If cml(i, k) = Empty Then Exit For
                If cml(i, k) = "Trade Name" Then
                    j = k
                    Exit For
                End If
            Next
        Else
            cml(i, j) = Replace(CStr(cml(i, j)), Chr(34), "") 'remove " character
        End If
    Next
    
    remove_special_characters = cml
    
End Function

Function convert_TE_cml(ws_data As Variant, header_row As Long) As Variant
'returns CML from TE

    Dim temp() As Variant
    Dim col_b1 As Variant   'sap b1 columns
    Dim col_te As Variant   'trade edge columns
    Dim header As Object
    Dim i As Long           'row
    Dim j As Long           'column
    Dim k As Long
    Dim n As Long
    Dim col As New clsColumnCML
    
    'sap b1 columns
    col_b1 = Array("Company Name", "Code", "Trade Name", "RTM Model Code", "RTM Model Name", _
            "Customer Class Code", "Customer Class Name", "Customer Class Group Code", _
            "Customer Class Group Name", "Channel Type Code", "Channel Type Name", _
            "Phone No.", "BIR TIN", "Area", "Owner", "Address", "Zip Code", "Province", _
            "Town", "Barangay", "Cluster", "RD Salesman Code", "RD Salesman Name", _
            "Discount/" & vbLf & "Markup", _
            "Credit Line", "Credit Terms", "Status", "Lead Account", "Tax Acct.", _
            "Date of 1st Purchase", "Freezer")
    
    For j = 1 To UBound(ws_data, 2)
            
        'change column names
        Select Case ws_data(header_row, j)
            Case Is = "CUSTOMER_ID"
                col.ID = j
            Case Is = "CUSTOMER_NUMBER", "*Customer Offline Number", "NEXT_UP_NUMBER"  '"CUSTOMER_ID"
                col.Code = j
            Case Is = "CUSTOMER_NAME", "*Customer Name(Remove all special characters)":
                col.Trade_Name = j
            Case Is = "CSTNUMBER", "CST Number(Mark Up)"
                col.discount = j
            Case Is = "TINNUMBER", "*TIN Number"
                col.TIN = j
            Case Is = "KEY_ACCOUNT", "Key Account(Channel Type)"
                col.Channel_Name = j
            Case Is = "STATUS", "*Status(Always the same)"
                col.Status = j
            Case Is = "BRANCH_NAME", "Servicing Branch Name"
                col.Company_Name = j
            Case Is = "FIRST_NAME", "*Contact Person First Name"
                col.Owner = j
            Case Is = "LAST_NAME", "*Contact Person Last Name"
                col.LastName = j
            Case Is = "CONTACT_NO", "*Contact Person Phone Number(No special Character)"
                col.Phone = j
            Case Is = "ADDRESS1", "*Address Line 1"
                col.Address = j
            Case Is = "GEO_LOCATION_HIERARCHYDESCRIPTION", "Address Line 2(Barangay)"
            'Case Is = "ADDRESS2", "Address Line 2(Barangay)"
                col.Barangay = j
            Case Is = "CITY", "*City/Town"
                col.Town = j
            Case Is = "STATE_PROVINCE", "*State(Province)"
                col.Province = j
            Case Is = "POSTAL_CODE", "*Postal Code(Zip Code)"
                col.Zip_Code = j
            Case Is = "CHANNEL", "*Customer Channel(Customer Type)"
                col.RTM_Model = j 'customer type
            Case Is = "CREDIT_LIMIT", "*Credit Limit"
                col.Credit_Line = j
            Case Is = "CREDIT_LIMIT_DAYS", "*Credit Limit (Days-Only numerical)"
                col.Credit_Terms = j
            Case Is = "RETAIL_ENVIRONMENT", "*Retail Environment(Segmentation)"
                col.Segmentation_Name = j
            Case Is = "PARTY_CLASSIFICATION_DESCRIPTION", "*Customer Class"
                col.Customer_Class_Name = j
            Case Is = "TAX_CODE"
                col.Tax = j
            Case Is = "SALES_REP_ID", "Sales Rep ID(Mandatory)"
                col.DSP_Code = j
            Case Is = "SALES_REP_NAME", "Salesman Name"
                col.DSP_Name = j
        End Select
    Next
    
    n = 1
    For i = header_row + 1 To UBound(ws_data, 1)
        If ws_data(i, col.Code) <> "" And ws_data(i, col.Code) <> "CUSTOMER_NUMBER" Then
            n = n + 1
            ReDim Preserve temp(1 To 32, 1 To n)
            
            If n = 2 Then
                For j = 0 To UBound(col_b1)
                    temp(j + 1, 1) = col_b1(j)
                Next
            End If
            
            temp(1, n) = ws_data(i, col.Company_Name)
            temp(2, n) = ws_data(i, col.Code)
            temp(3, n) = ws_data(i, col.Trade_Name)
            temp(4, n) = get_TE_rtm(CStr(ws_data(i, col.RTM_Model)))
            temp(5, n) = ws_data(i, col.RTM_Model)
            temp(6, n) = get_TE_class_code(CStr(ws_data(i, col.Customer_Class_Name)))
            temp(7, n) = ws_data(i, col.Customer_Class_Name)
            temp(8, n) = get_TE_seg_code(CStr(ws_data(i, col.Segmentation_Name)))
            temp(9, n) = ws_data(i, col.Segmentation_Name)
            temp(10, n) = get_TE_ch_code(CStr(ws_data(i, col.Channel_Name)))
            temp(11, n) = ws_data(i, col.Channel_Name)
            temp(12, n) = ws_data(i, col.Phone)
            temp(13, n) = ws_data(i, col.TIN)
            temp(14, n) = "" 'Area
            temp(15, n) = CStr(ws_data(i, col.Owner)) & " " & CStr(ws_data(i, col.LastName))
            temp(16, n) = ws_data(i, col.Address)
            temp(17, n) = ws_data(i, col.Zip_Code)
            temp(18, n) = ws_data(i, col.Province)
            temp(19, n) = ws_data(i, col.Town)
            temp(20, n) = ws_data(i, col.Barangay)
            temp(21, n) = "" 'Cluster
            temp(22, n) = ws_data(i, col.DSP_Code)
            temp(23, n) = ws_data(i, col.DSP_Name)
            temp(24, n) = get_discount(ws_data(i, col.discount), CStr(temp(4, n)))
'            temp(24, n) = IIf(WorksheetFunction.Or(ws_data(i, col.discount) = "M2", ws_data(i, col.discount) = -2), "2% markup", _
'                          IIf(WorksheetFunction.Or(ws_data(i, col.discount) = "M0", ws_data(i, col.discount) = 0), "0% markup", ws_data(i, col.discount)))
            temp(25, n) = ws_data(i, col.Credit_Line)
            temp(26, n) = ws_data(i, col.Credit_Terms)
            temp(27, n) = IIf(ws_data(i, col.Status) = "Active/Approved", "ACT", "IAC")
            temp(28, n) = "" 'Lead Account
            temp(29, n) = 12 'Tax
            temp(30, n) = "" 'Date of first purchase
            temp(31, n) = "" 'Freezer
            temp(32, n) = ws_data(i, col.ID)
        End If
    Next
    
    convert_TE_cml = Application.WorksheetFunction.Transpose(temp)

End Function

Function get_discount(discount As Variant, rtm As String) As Variant
    
    If UCase(Left(rtm, 3)) = "VAN" Or UCase(Left(rtm, 2)) = "EX" Then
        get_discount = "2% markup"
        Exit Function
    ElseIf UCase(Left(rtm, 4)) = "BOOK" Then
        get_discount = "0% markup"
        Exit Function
    End If
    
    Select Case discount
    Case Is = "M2", -2
        get_discount = "2% markup"
    Case Is = "M0", 0
        get_discount = "0% markup"
    Case Is = ""
        If UCase(Left(rtm, 3)) = "VAN" Or UCase(Left(rtm, 2)) = "EX" Then
            get_discount = "2% markup"
        Else
            get_discount = "0% markup"
        End If
    Case Else
        get_discount = discount
    End Select
    
End Function

Function MergeArrays(arr1 As Variant, arr2 As Variant) As Variant
'returns combined CML arrays
'provides the "Province" column to old SAP B1 CML lacking it

    Dim i As Long       'row
    Dim j As Long       'column
    Dim k As Long       'counter
    Dim m As Long       'header row for arr1
    Dim n As Long       'header row for arr2
    Dim arr As Variant
    Dim dict As Object
    Dim temp As Variant
    Dim temp2 As Variant
    
    Set dict = CreateObject("Scripting.Dictionary")
    
check_arr1:
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        For j = LBound(arr1, 2) To UBound(arr1, 2)
            If arr1(i, j) <> "" Then
                dict.Add arr1(i, j), 0
                If j = UBound(arr1, 2) Then
                    m = i
                    GoTo check_arr2
                End If
            Else
                dict.Remove arr1(i, j - 1)
                Exit For
            End If
        Next j
    Next i
    
check_arr2:
    k = 0
    For i = LBound(arr2, 1) To UBound(arr2, 1)
        For j = LBound(arr2, 2) To UBound(arr2, 2)
            If arr2(i, j) <> "" Then
                If dict.Exists(arr2(i, j)) Then
                    k = k + 1
                End If
                
                If k >= dict.count Or j = UBound(arr2, 2) Then
                    n = i
                    GoTo merge_now
                End If
            Else
                Exit For
            End If
        Next j
    Next i
            
merge_now:
    
    'ReDim arr(LBound(arr1, 1) To UBound(arr1, 1) + UBound(arr2, 1) - m - n, LBound(arr1, 2) To dict.count)
    ReDim arr(LBound(arr1, 1) To UBound(arr1, 1) + UBound(arr2, 1) - m - n + 1, LBound(arr1, 2) To dict.count) '10-28-2021
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If i <= (UBound(arr1, 1) - m + 1) Then
                arr(i, j) = arr1(m + i - 1, j)
            ElseIf UBound(arr2, 2) < UBound(arr, 2) Then
                If j = UBound(arr2, 2) Then
                    If arr(1, j + 1) = arr2(n, j) Then
                        arr(i, j) = arr2(i - (UBound(arr1, 1) - m + 1) + n + 1, j - 1)
                        arr(i, j + 1) = arr2(i - (UBound(arr1, 1) - m + 1) + n + 1, j)
                        Exit For
                    End If
                ElseIf arr(1, j) = arr2(n, j) Then
                    arr(i, j) = arr2(i - (UBound(arr1, 1) - m + 1) + n + 1, j)
                ElseIf Right(arr(1, j), 10) = Right(arr2(n, j), 10) Then
                    arr(i, j) = arr2(i - (UBound(arr1, 1) - m + 1) + n + 1, j)
                ElseIf arr(1, j) = arr2(n, j - 1) Then
                    arr(i, j) = arr2(i - (UBound(arr1, 1) - m + 1) + n + 1, j - 1)
                ElseIf arr(1, j) = "Province" And arr2(n, j) = "Town" Then
                    arr(i, j) = getProvince(arr2(i - (UBound(arr1, 1) - m + 1) + n + 1, j - 1))
                End If
            Else
                If arr(1, j) = arr2(n, j) Then
                    'arr(i, j) = arr2(i - (UBound(arr1, 1) - m + 1) + n + 1, j)
                    arr(i, j) = arr2(i - (UBound(arr1, 1) - m + 1) + n, j)      '10-28-2021
                End If
            End If
        Next j
    Next i
    
    MergeArrays = arr
    
End Function

Function consolidate(mcp As Variant, cml As Variant) As Variant
'returns mcp sheet data with alignment checking vs CML data

    Dim i As Long, j As Long, k As Long
    Dim m As Long, n As Long, p As Long
    Dim x As Long, r As Long, y As Long
    Dim arr As Variant
    Dim no_cml() As Variant, no_cml2() As Variant, iac() As Variant, iac2() As Variant
    Dim no_mcp() As Variant, no_mcp2() As Variant
    Dim wrong_dsp() As Variant, wrong_dsp2() As Variant
    Dim wrong_rtm() As Variant, wrong_rtm2() As Variant
    Dim inMCP As Boolean
    Dim hdr As Variant
    Dim s As Long
    Dim rtm As String
    Dim col_rtm As Long
    Dim dict As Object
    Dim w As Long
    Dim col As New clsColumnCML
    Dim Lst_Col As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    arr = mcp
    
    i = 0
    j = 0
    k = 0
    m = 1
    n = 1
    p = 1
    r = 1
    s = 1
    x = 0
    w = 0
    
    'find column numbers from cml file
    Lst_Col = IIf(UBound(cml, 2) > 31, 31, UBound(cml, 2))
    
    If cml(1, Lst_Col) <> "" Then
        w = 1
    ElseIf cml(2, Lst_Col) <> "" Then
        w = 2
    End If
    
    For j = 1 To UBound(cml, 2)
        
        Select Case cml(w, j)
            Case Is = "Code"
                col.Code = j
            Case Is = "Trade Name"
                col.Trade_Name = j
            Case Is = "RTM Model Code"
                col.RTM_Model = j
            Case Is = "Service Model Code"
                col.RTM_Model = j
            Case Is = "Customer Class Code"
                col.Customer_Class_Code = j
            Case Is = "Address"
                col.Address = j
            Case Is = "Zip Code"
                col.Zip_Code = j
            Case Is = "Province"
                col.Province = j
            Case Is = "Town"
                col.Town = j
            Case Is = "Barangay"
                col.Barangay = j
            Case Is = "RD Salesman Code"
                col.RD_Salesman_Code = j
            Case Is = "Credit Terms"
                col.Credit_Terms = j
            Case Is = "Status"
                col.Status = j
        End Select
    Next
    
    'merge selected cml data to mcp consolidator template on customer code
    For k = 2 To UBound(cml, 1)
            
        inMCP = False
        
        If cml(k, col.Code) <> "" Then 'And cml(k, col.Status) = "ACT" Then
            
            For i = LBound(arr, 1) To UBound(arr, 1)
            
                If CStr(arr(i, 2)) = CStr(cml(k, col.Code)) Then
                    
                    'Customer Name
                    arr(i, 3) = Replace(CStr(cml(k, col.Trade_Name)), Chr(34), "")
                    
                    'Province
                    If col.Province > 0 Then
                        arr(i, 4) = UCase(cml(k, col.Province))
                    Else
                        arr(i, 4) = getProvince(cml(k, col.Zip_Code))
                    End If
                    
                    'Town
                    arr(i, 5) = UCase(cml(k, col.Town))
                    
                    'Barangay
                    arr(i, 6) = UCase(cml(k, col.Barangay))
                    
'                    arr(i, 5) = Replace(cml(k, 19), "Ã‘", "Ñ")
'                    arr(i, 6) = Replace(cml(k, 20), "Ã‘", "Ñ")
                    
                    'Address
                    arr(i, 7) = Replace(cml(k, col.Address), arr(i, 6) & ", " & arr(i, 5) & ", " & arr(i, 4), "")
                    
                    'Customer Class
                    arr(i, 9) = IIf(cml(k, col.Customer_Class_Code) = "", arr(i, 9), cml(k, col.Customer_Class_Code))
                    
                    'Credit Term
                    arr(i, 10) = cml(k, col.Credit_Terms)
                    
                    
                    If arr(i, 11) <> cml(k, col.RD_Salesman_Code) Then
                        r = r + 1
                        ReDim Preserve wrong_dsp(1 To 7, 1 To r)
                        wrong_dsp(1, r) = cml(k, 1) 'Distributor
                        wrong_dsp(2, r) = cml(k, col.Code) 'Customer Code
                        wrong_dsp(3, r) = Replace(CStr(cml(k, col.Trade_Name)), Chr(34), "") 'Customer Name
                        wrong_dsp(4, r) = cml(k, col.RD_Salesman_Code) 'RD Salesman Code
                        wrong_dsp(5, r) = cml(k, col.RD_Salesman_Code + 1) 'RD Salesman Name
                        wrong_dsp(6, r) = arr(i, 11) 'Salesman_Code
                        wrong_dsp(7, r) = arr(i, 12) 'Salesman_Name
                        
                    End If
                    
                    rtm = Mid(arr(i, 13), WorksheetFunction.Find("(", arr(i, 13), 1) + 1, _
                        WorksheetFunction.Find(")", arr(i, 13), 1) - WorksheetFunction.Find("(", arr(i, 13), 1) - 1)
                        
                    If rtm <> cml(k, col.RTM_Model) Then
                        If (cml(k, col.RTM_Model) = "KCM" Or _
                            cml(k, col.RTM_Model) = "SEC" Or _
                            cml(k, col.RTM_Model) = "BTDT" Or _
                            cml(k, col.RTM_Model) = "OPM" Or _
                            cml(k, col.RTM_Model) = "EXT") _
                            Or _
                            (InStr(1, UCase(cml(k, col.RTM_Model)), "BOOK") And _
                            rtm = "EXT") _
                            Or _
                            (InStr(1, UCase(cml(k, col.RTM_Model)), "VAN") And _
                            rtm <> "EXT") Then
                            
                            s = s + 1
                            ReDim Preserve wrong_rtm(1 To 11, 1 To s)
                            wrong_rtm(1, s) = cml(k, 1) 'Distributor
                            wrong_rtm(2, s) = cml(k, col.Code) 'Customer Code
                            wrong_rtm(3, s) = Replace(CStr(cml(k, col.Trade_Name)), Chr(34), "") 'Customer Name
                            wrong_rtm(4, s) = cml(k, col.RTM_Model) 'RTM Model Code
                            wrong_rtm(5, s) = cml(k, col.RTM_Model + 1) 'RTM Model Name
                            wrong_rtm(6, s) = IIf(InStr(1, UCase(wrong_rtm(4, s)), "EXT"), "VAN(Extruck)", "BOOK(Booking)")
                            wrong_rtm(7, s) = arr(i, 11) 'Salesman_Code
                            wrong_rtm(8, s) = arr(i, 12) 'Salesman_Name
                            wrong_rtm(9, s) = rtm
                            wrong_rtm(10, s) = arr(i, 13) 'RTM Model Name
                            wrong_rtm(11, s) = IIf(rtm = "EXT", "VAN(Extruck)", "BOOK(Booking)")
                        
                        End If
                    End If
                    
                    inMCP = True
                    
                    Exit For
                    
                End If
            Next
        End If
        
        If cml(k, col.Code) <> "Code" And _
            cml(k, col.Code) <> "" And _
            cml(k, col.Status) = "ACT" And _
            inMCP = False Then
            
            m = m + 1
            ReDim Preserve no_mcp(1 To UBound(cml, 2), 1 To m)
            For x = 1 To UBound(cml, 2)
                On Error Resume Next
                no_mcp(x, m) = cml(k, x)
            Next
        End If
        
    Next
    
    For i = 1 To UBound(arr, 1)
        
        y = IsInArray(CStr(arr(i, 2)), cml)
        
        If y < 1 Then
            n = n + 1
            ReDim Preserve no_cml(1 To 30, 1 To n)
            For x = 1 To 30
                no_cml(x, n) = arr(i, x)
            Next
        ElseIf cml(y, col.Status) = "IAC" Then
            p = p + 1
            ReDim Preserve iac(1 To 30, 1 To p)
            For x = 1 To 30
                iac(x, p) = arr(i, x)
            Next
        End If
        
    Next
            
    If m > 1 Then
        ReDim Preserve no_mcp2(1 To UBound(no_mcp, 2), 1 To UBound(no_mcp, 1))
        
        For i = 1 To 10
            If cml(i, UBound(cml, 2)) <> "" Then
                w = i
                Exit For
            End If
        Next
        
        For i = 1 To UBound(no_mcp, 2)
            For j = 1 To UBound(no_mcp, 1)
                If i = 1 Then
                    no_mcp2(i, j) = cml(w, j)
                Else
                    no_mcp2(i, j) = no_mcp(j, i)
                End If
            Next
        Next
        
        AddNew no_mcp2, "NOT_IN_MCP_" & UCase(no_mcp2(3, 1)), "notinMCP"
        
        without_mcp = m - 1
    End If
    
    If n > 1 Then
        ReDim Preserve no_cml2(1 To UBound(no_cml, 2), 1 To UBound(no_cml, 1))
        
        hdr = Array("Distributor_Code", "Customer_Code", "Customer_Name", _
            "Province", "Town/Cty", "Barangay", "No./Bldg./Street/Landmark", _
            "Barangay Code", "Customer_Class", "Credit_Term", "Salesman_Code", _
            "SalesmanName", "Service_Model_Name", "Supervisor_Code", "Supervisor_Name", _
            "Frequency", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", _
            "WK1", "WK2", "WK3", "WK4", "Sequence", "Day", "REMARKS", "Flag")
        
        For i = 1 To UBound(no_cml, 2)
            For j = 1 To UBound(no_cml, 1)
                If i = 1 Then
                    no_cml2(i, j) = hdr(j - 1)
                Else
                    On Error Resume Next
                    no_cml2(i, j) = IIf(Left(no_cml(j, i), 1) = "=", "", no_cml(j, i))
                End If
            Next
        Next
        
        AddNew no_cml2, "NOT_IN_CML_" & no_cml2(2, 1), "notinCML"
        
        without_cml = n - 1
    End If
    
    If p > 1 Then
        ReDim Preserve iac2(1 To UBound(iac, 2), 1 To UBound(iac, 1))
        
        hdr = Array("Distributor_Code", "Customer_Code", "Customer_Name", _
            "Province", "Town/Cty", "Barangay", "No./Bldg./Street/Landmark", _
            "Barangay Code", "Customer_Class", "Credit_Term", "Salesman_Code", _
            "SalesmanName", "Service_Model_Name", "Supervisor_Code", "Supervisor_Name", _
            "Frequency", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", _
            "WK1", "WK2", "WK3", "WK4", "Sequence", "Day", "REMARKS", "Flag")
        
        For i = 1 To UBound(iac, 2)
            For j = 1 To UBound(iac, 1)
                If i = 1 Then
                    iac2(i, j) = hdr(j - 1)
                Else
                    On Error Resume Next
                    iac2(i, j) = IIf(Left(iac(j, i), 1) = "=", "", iac(j, i))
                End If
            Next
        Next
        
        AddNew iac2, "INACTIVE_IN_CML_" & iac2(2, 1), "inactive"
        
        inactive_cml = p - 1
    End If
    
    If r > 1 Then
        ReDim Preserve wrong_dsp2(1 To UBound(wrong_dsp, 2), 1 To UBound(wrong_dsp, 1))
        hdr = Array("Distributor", "Customer Code", "Customer Name", "DMS Salesman Code", _
                    "DMS Salesman Name", "MCP Salesman_Code", "MCP Salesman_Name")
        
        For i = 1 To UBound(wrong_dsp, 2)
            For j = 1 To UBound(wrong_dsp, 1)
                 If i = 1 Then
                    wrong_dsp2(i, j) = hdr(j - 1)
                Else
                    wrong_dsp2(i, j) = wrong_dsp(j, i)
                End If
            Next
        Next
        
        AddNew wrong_dsp2, "WRONG_SALESMAN_" & UCase(wrong_dsp2(2, 1)), "wrongSalesman"
        
        different_dsp = r - 1
    End If
        
    If s > 1 Then
        
        ReDim Preserve wrong_rtm2(1 To UBound(wrong_rtm, 2), 1 To UBound(wrong_rtm, 1))
        hdr = Array("Distributor", "Customer Code", "Customer Name", "DMS RTM Model Code", _
                    "DMS RTM Model Name", "DMS Customer Type", _
                    "Salesman_Code in MCP", "Salesman_Name in MCP", _
                    "Should be RTM Model Code", "Should be RTM Model Name", "Should be Customer Type")
        
        For i = 1 To UBound(wrong_rtm, 2)
            For j = 1 To UBound(wrong_rtm, 1)
                 If i = 1 Then
                    wrong_rtm2(i, j) = hdr(j - 1)
                Else
                    wrong_rtm2(i, j) = wrong_rtm(j, i)
                End If
            Next
        Next
        
        AddNew wrong_rtm2, "WRONG_RTM_IN_DMS_" & UCase(wrong_rtm2(2, 1)), "wrongRTM"
        
        different_rtm = s - 1
        
    End If
            
    consolidate = arr
        
End Function

Sub SFA_Uploader()
'creates route master (MCP) and retailer master (CML) upload files for SFA
'shortcut: Ctrl + j

    Dim mcp As Variant
    Dim mcp_head As Variant
    Dim route_plan() As Variant
    Dim route_temp() As Variant
    Dim route_temp2() As Variant
    Dim route_map() As Variant
    Dim route_map2() As Variant
    Dim temp() As Variant
    Dim temp2() As Variant
    Dim LstRow As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim rowNum As Long
    Dim key As Variant
    Dim x As Variant
    Dim cml_sfa As Variant
    Dim route_day As String
    Dim frequency As String
    Dim salesman_code As String
    Dim recurrence As String
    Dim route_code As Object
    Dim cml_files As Variant, cml_file As Variant, cml_data As Variant
    Dim file As Variant
    Dim mcp_wkbk As Workbook, cml_wkbk As Workbook
    Dim startTime As Single, endTime As Single, totalTime As Single
    Dim msg As String
    Dim cml_dict As Object

    SwitchOff True
    On Error GoTo Err_handler
    
    i = 0
    j = 0
    k = 1
    
    errors_found = ""
    
    mod_names = ThisWorkbook.Worksheets("ref_service_model").Range("serv_mod_name").value
    serv_model = ThisWorkbook.Worksheets("ref_service_model").Range("serv_model").value
    
    Set mcp_miss = CreateObject("Scripting.Dictionary")
    Set route_code = CreateObject("Scripting.Dictionary")
    Set cml_dict = CreateObject("Scripting.Dictionary")
    
    If MsgBox("Generate Route Plan for Trade Edge SFA?", vbYesNoCancel) = vbYes Then
        
        With ThisWorkbook
            mcp = .Worksheets("MCP").Range("tblMCP").value
            mcp_head = .Worksheets("MCP").Range("tblMCPhead").value
        End With
        
        If IsArrayAllocated(mcp) Then
            If MsgBox("Add customer master.", vbYesNo, Title:=ver) = vbYes Then
                cml_files = select_file(True)
            End If
        End If
        
        '5. Process CML workbook
        If IsArrayAllocated(cml_files) = False Then
            GoTo skip_cml
        Else
            z = 0
            For Each file In cml_files
                               
                z = z + 1
                startTime = Timer
                
                cml_file = CStr(file)
                Set cml_wkbk = open_workbook(cml_file)
                
                endTime = Timer
                totalTime = totalTime + (endTime - startTime)
                msg = msg & vbCrLf & "Opening the CML #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."
                
                startTime = Timer
                
                cml_data = iterate_cml(cml_wkbk, cml_data)
                cml_wkbk.Close
                
                endTime = Timer
                totalTime = totalTime + (endTime - startTime)
                msg = msg & vbCrLf & "Getting the CML #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."
    
            Next
        
        End If
        
        Set cml_wkbk = Nothing
        
        For i = LBound(cml_data, 1) + 1 To UBound(cml_data, 1)
            If cml_dict.Exists(cml_data(i, 2)) Then                 'check for duplicates
                If cml_data(i, 27) = "ACT" Then
                    cml_dict(cml_data(i, 2)) = cml_data(i, 32)      'replace with active cml
                End If
            Else
                cml_dict.Add cml_data(i, 2), cml_data(i, 32)        'add to dictionary if no duplicates
            End If
        Next
        
skip_cml:
        For i = LBound(mcp, 1) To UBound(mcp, 1)
        
            'branch code + retailer code + salesman code + service model
            If mcp(i, 1) <> "" And mcp(i, 2) <> "" And mcp(i, 11) <> "" Then
                If mcp(i, 16) <> "" Then
                    k = k + 1
                    frequency = ""
                    recurrence = ""
                    route_day = ""
                    
                    ReDim Preserve route_temp(1 To 16, 1 To k)
                    
                    salesman_code = mcp(i, 11)                     'salesrep_code
                    
                    For j = 7 To 13                                 'route day
                        If mcp(i, 10 + j) = 1 Then
                            If route_day = "" Then
                                route_day = UCase(mcp_head(1, 10 + j))
                            Else
                                route_day = route_day & "_" & UCase(mcp_head(1, 10 + j))
                            End If
                        End If
                    Next
                    
                    
                    If Replace(mcp(i, 16), "F", "") >= 4 Then                       'frequency
                        frequency = "WKLY"
                        recurrence = "FIRST;SECOND;THIRD;FOURTH"
                    Else
                        For j = 14 To 17
                            If mcp(i, 10 + j) = 1 Then
                                If frequency = "" Then
                                    frequency = mcp_head(1, 10 + j)
                                    recurrence = Application.WorksheetFunction.Choose(Replace(mcp_head(1, 10 + j), "WK", ""), "FIRST", "SECOND", "THIRD", "FOURTH")
                                Else
                                    frequency = frequency & "_" & mcp_head(1, 10 + j)
                                    recurrence = recurrence & ";" & Application.WorksheetFunction.Choose(Replace(mcp_head(1, 10 + j), "WK", ""), "FIRST", "SECOND", "THIRD", "FOURTH")
                                End If
                            End If
                        Next
                    End If
                    
                    route_temp(1, k) = mcp(i, 1)                      'rd_code // should be branch code
                    route_temp(2, k) = salesman_code & "_" & _
                                       route_day & "_" & frequency    'route code = dspcode_day_freq TSM0002_FRI_WKLY
                    route_temp(3, k) = route_temp(2, k)               'route name = dspcode_day_freq TSM0002_FRI_WKLY
                    route_temp(4, k) = "Active"                       'route status = default Active
                    
                    If mcp(i, 13) = "Extruck Accounts (EXT)" Then     'route type = depends on service model
                        route_temp(5, k) = "Van Sales"
                    Else
                        route_temp(5, k) = "Order Booking"
                    End If
                    
                    route_temp(6, k) = "High"                         'route priority = default High
                    
                    route_temp(7, k) = IIf(IsArrayAllocated(cml_files) = False, CStr(mcp(i, 2)), cml_dict(CStr(mcp(i, 2))))           'retailer_code = lookup customer offline number from CML
                    
                    route_temp(8, k) = ""                             'facility code = lookup from separate file
                    
                    route_temp(9, k) = Date                          'start date =
                    
                    route_temp(10, k) = "Weekly"                      'frequency
                    
                    route_temp(11, k) = ""                            'weekly recurrence, not mandatory - none
                    route_temp(12, k) = Replace(route_day, "_", ";")                   'Day Of Week
                    route_temp(13, k) = ""                            'Day of Every Month - none
                    route_temp(14, k) = ""                            'Monthly recurrence - none
                    route_temp(15, k) = ""                            'Monthly recurrence day - none
                    route_temp(16, k) = recurrence                    'Weekly Recurrence Week
                    
                    If Not route_code.Exists(route_temp(2, k)) Then
                        route_code.Add route_temp(2, k), salesman_code
                    End If
                    
                Else
                    If Not mcp_miss.Exists(mcp(i, 2)) Then
                        mcp_miss.Add mcp(i, 2), mcp(i, 3)
                    End If
                End If
            End If
        Next
        
        If k = 1 Then
            MsgBox "No data"
            GoTo cml_data
        End If
        
        If VarType(route_temp) = (vbVariant + vbArray) Then
            ReDim route_temp2(LBound(route_temp, 2) To UBound(route_temp, 2), LBound(route_temp, 1) To UBound(route_temp, 1))
            route_temp2 = Application.WorksheetFunction.Transpose(route_temp)
            k = UBound(route_temp, 1)
        Else
            k = 0
        End If
        
        ' 16 columns te route plan
        x = Array("*Branch Code", "*Route Code", "*Route Name", "*Route Status", "*Route Type", _
            "*Route Priority", "*Customer Offline No.", "Facility Code", "*Start Date", "*Frequency ", _
            "Weekly Recurrence", "Day Of Week", "Day Of Every Month", "Monthly Recurrence", _
            "Monthly Recurrence Day", "Weekly Recurrence Week")
        
        For j = LBound(x) To UBound(x)
            route_temp2(1, j + 1) = x(j)
        Next
        
        AddNew route_temp2, "Route Plan for " & UCase(CStr(route_temp2(2, 1)))
        
        ' create route mapping
        ReDim Preserve route_map(1 To route_code.count + 1, 1 To 3)
        
        ' Write headers
        route_map(1, 1) = "*Route Code"
        route_map(1, 2) = "*Route Name"
        route_map(1, 3) = "*Employees"
    
        ' Start writing from row 2
        rowNum = 2
        For Each key In route_code.Keys
            route_map(rowNum, 1) = key
            route_map(rowNum, 2) = key
            route_map(rowNum, 3) = route_code(key)
            rowNum = rowNum + 1
        Next key
        
'        If VarType(route_map) = (vbVariant + vbArray) Then
'            ReDim route_map2(LBound(route_map, 2) To UBound(route_map, 2), LBound(route_map, 1) To UBound(route_map, 1))
'            route_map2 = Application.WorksheetFunction.Transpose(route_map)
'            k = UBound(route_map, 1)
'        Else
'            k = 0
'        End If
    
        AddNew route_map, "Route Mapping for " & UCase(CStr(route_temp2(2, 1)))
    
    End If
    
    'for ivy SFA
'    If MsgBox("Generate Route Master for IvyMobility SFA?", vbYesNoCancel) = vbYes Then
'
'        With ThisWorkbook
'            mcp = .Worksheets("MCP").Range("tblMCP").value
'            rtm_mod = .Worksheets("ref").Range("rtm_model").value
'        End With
'
'        old_codes = Range("old_codes").value
'        new_codes = Range("new_codes").value
'
'        For i = LBound(mcp, 1) To UBound(mcp, 1)
'
'            'rd code + retailer code + salesman code
'            If mcp(i, 1) <> "" And mcp(i, 2) <> "" And mcp(i, 11) <> "" Then
'                If mcp(i, 16) <> "" Then
'                    k = k + 1
'                    ReDim Preserve temp(1 To 19, 1 To k)
'                    temp(1, k) = mcp(i, 1)                      'rd_code
'                    temp(4, k) = mcp(i, 2)                      'retailer_code
'                    temp(5, k) = mcp(i, 11)                     'salesrep_code
'                    temp(6, k) = Replace(mcp(i, 16), "F", "")   'frequency
'
'                    For j = 7 To 17                             'days
'                        If mcp(i, 10 + j) = "" Then
'                            temp(j, k) = "N"
'                        Else
'                            temp(j, k) = "Y"
'                            If j < 14 Then
'                                temp(2, k) = mcp(i, 11) '& "_" & UCase(Format(j - 5, "ddd"))
'                                temp(3, k) = mcp(i, 11) '& "_" & UCase(Format(j - 5, "dddd"))
'                            End If
'                        End If
'                    Next
'                    temp(18, k) = mcp(i, 28)                    'sequence
'                Else
'                    If Not mcp_miss.exists(mcp(i, 2)) Then
'                        mcp_miss.Add mcp(i, 2), mcp(i, 3)
'                    End If
'                End If
'            End If
'        Next
'
'        If k = 1 Then
'            MsgBox "No data"
'            GoTo cml_data
'        End If
'
'
'            ReDim temp2(1 To UBound(temp, 2), 1 To UBound(temp, 1))
'
'        For i = 1 To UBound(temp, 2)
'            For j = 1 To UBound(temp, 1)
'                temp2(i, j) = temp(j, i)
'            Next
'        Next
'
'        x = Array("Distributor_Code", "Route_Code", "Route_Name", "Retailer_Code", "DSR_Code", _
'            "Visit_Frequency", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", _
'            "Sunday", "WK1", "WK2", "WK3", "WK4", "Sequence", "Flag")
'
'        For j = LBound(x) To UBound(x)
'            temp2(1, j + 1) = x(j)
'        Next
'
'        AddNew temp2, "Route_Master_for_SFA_" & UCase(CStr(temp2(2, 1)))
'
'    End If
'
cml_data:
'
'    If MsgBox("Generate Retailer Master (CML) for SFA?", vbYesNoCancel) = vbYes Then
'
'        cml_sfa = retailer_master(mcp)
'
'        If VarType(cml_sfa) = 0 Then GoTo Exit_If
'
'        AddNew cml_sfa, "Retailer_Master_for_SFA_" & UCase(CStr(cml_sfa(2, 1)))
'
'    End If

Exit_If:
    
    If mcp_miss.count > 0 Then
        errors_found = errors_found & mcp_miss.count & " no week schedule in MCP."
    End If
    
    
'    MsgBox "Created successfully!" & vbCrLf & _
'            vbCrLf & _
'           "Generated upload files (Route Master and Retailer Master)" & vbCrLf & _
'           "will update or add to existing master in SFA. However, " & vbCrLf & _
'           "these will not remove or de-activate old customer " & vbCrLf & _
'           "existing in SFA device. " & vbCrLf & vbCrLf & _
'           "To remove old customer in SFA device: " & vbCrLf & _
'           "1. Extract the Route Plan from SFA portal into Excel, " & vbCrLf & _
'           "2. Add 'Flag' column and put 'D' per line item to deactivate, and " & vbCrLf & _
'           "3. Send as 'Deactivate_file.xlsx' along with this " & vbCrLf & _
'           "   Route Master and Retailer Master. " & vbCrLf & _
'           vbCrLf & _
'           errors_found, _
'           Title:="MCP Consolidator v" & ver

    MsgBox "Done!" & vbCrLf & _
            vbCrLf & _
            vbCrLf & _
            errors_found, _
            Title:="MCP Consolidator v" & ver
    
Exit_sub:
    SwitchOff False
    Exit Sub
Err_handler:
    MsgBox Err.Number & " - " & Err.Description
    GoTo Exit_sub
End Sub
    
Function cml_dictionary(cml_files As Variant) As Object
'5. Process CML workbook
    Dim z As Long
    Dim cml_file As String
    Dim cml_wkbk As Workbook
    Dim msg As String
    Dim cml As Object
    
    Set cml = CreateObject("Scripting.Dictionary")
    
    If IsArrayAllocated(cml_files) = False Then
        'GoTo Paste_Values
    Else
        z = 0
        For Each file In cml_files
                           
            z = z + 1
            startTime = Timer
            
            cml_file = CStr(file)
            Set cml_wkbk = open_workbook(cml_file)
            
            endTime = Timer
            totalTime = totalTime + (endTime - startTime)
            msg = msg & vbCrLf & "Opening the CML #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."
            
            startTime = Timer
            
            cml_data = iterate_cml(cml_wkbk, cml_data)
            cml_wkbk.Close
            
            endTime = Timer
            totalTime = totalTime + (endTime - startTime)
            msg = msg & vbCrLf & "Getting the CML #" & z & ": " & vbTab & Round(endTime - startTime, 4) & " sec."

    
End Function


Function retailer_master(mcp As Variant) As Variant
'returns customer master template for SFA

    Dim cml_names As Variant
    Dim cml_name As Variant
    Dim file As Variant
    Dim cml_book As Workbook
    Dim cml_data As Variant
    Dim i As Long
    
    '1. Select CML file
    MsgBox "Open CML file", vbOKOnly
    cml_names = select_file(True)
    
    If IsArrayAllocated(cml_names) = False Then
        MsgBox "No file selected"
        Exit Function
    Else
        For Each file In cml_names
                           
            i = i + 1
            
            cml_name = CStr(file)
            
            '2. Open CML workbook
            Set cml_book = open_workbook(cml_name)
                        
            '3. Put CML into array
            cml_data = iterate_cml(cml_book, cml_data)
            
            cml_book.Close
        Next
    
    End If
    
    retailer_master = extract_cml(cml_data, mcp)
    
End Function

Function extract_cml(cml As Variant, mcp As Variant) As Variant
'returns aligned customer master to mcp

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim n As Long
    Dim x As Variant
    Dim temp() As Variant
    Dim temp2() As Variant
    Dim dsp_row As Long
    Dim cust_chan As Variant
    Dim cust_add As String
    Dim RTM_Model As Variant
    Dim def_model As Variant
    Dim ret_rd As Long
    Dim ret_code As Long, ret_name As Long, ret_type As Long
    Dim ret_clas As Long, ret_num As Long, ret_tin As Long
    Dim ret_own As Long, ret_add As Long, ret_zip As Long
    Dim ret_twn As Long, ret_bgy As Long, ret_dis As Long
    Dim ret_lin As Long, ret_trm As Long, ret_stat As Long
    Dim ret_dsp As Long, ret_man As Long, ret_prov As Long
    Dim dist_code As String
    Dim err_dsp As Object
    Dim err_d As clsDSP
    Dim err_c As clsChannel
    Dim err_chnl As Object
    Dim err_rtm As Object
    Dim cust_r As String
    Dim err_r As clsChannel
    Dim sBrgy As String
    Dim sTown As String
    Dim sProv As String
    Dim cml As Object
    
    i = 0
    j = 0
    k = 1
    no_address = 0
    Set err_d = New clsDSP
    Set err_c = New clsChannel
    Set err_r = New clsChannel
    
    Set err_dsp = CreateObject("Scripting.Dictionary")
    Set err_chnl = CreateObject("Scripting.Dictionary")
    Set err_rtm = CreateObject("Scripting.Dictionary")
    Set cml = CreateObject("Scripting.Dictionary")
    
    If VarType(cml) = 0 Then Exit Function
    
    Dim obj As Object
    Dim test As Variant
    Set obj = CreateObject("Scripting.Dictionary")
    
    test = Array("Company Name", "Code", "Trade Name", "RTM Model Code", "RTM Model Name", _
    "Customer Class Code", "Customer Class Name", "Customer Class Group Code", "Customer Class Group Name", _
    "Channel Type Code", "Channel Type Name", "Phone No.", "BIR TIN", "Area", "Owner", "Address", _
    "Zip Code", "Province", "Town", "Barangay", "Cluster", "RD Salesman Code", "RD Salesman Name", _
    "Credit Line", "Credit Terms", "Status", "Lead Account", "Tax Acct.", "Date of 1st Purchase", "Freezer", "Customer ID")
    
    For i = LBound(test) To UBound(test)
        obj.Add test(i), 0
    Next
    
    For i = LBound(cml, 1) To 10
        For j = LBound(cml, 2) To UBound(cml, 2)
            If obj.Exists(cml(i, j)) Then
                n = n + 1
                If n >= UBound(test) - 5 Then
                    m = i
                    GoTo Lets_go
                End If
            ElseIf cml(i, j) = "" Then
                Exit For
            End If
        Next
    Next

cml_dict:
    For j = 1 To UBound(cml, 2)
        

Lets_go:
    For j = 1 To UBound(cml, 2)
        Select Case cml(m, j)
            Case Is = "Company Name"
                ret_rd = j
            Case Is = "Code"
                ret_code = j
            Case Is = "Trade Name"
                ret_name = j
            Case Is = "Customer Type Code"
                ret_type = j
            Case Is = "Service Model Code"
                ret_type = j
            Case Is = "RTM Model Code"
                ret_type = j
            Case Is = "Customer Class Code"
                ret_clas = j
            Case Is = "Phone No."
                ret_num = j
            Case Is = "BIR TIN"
                ret_tin = j
            Case Is = "Owner"
                ret_own = j
            Case Is = "Address"
                ret_add = j
            Case Is = "Zip Code"
                ret_zip = j
            Case Is = "Province"
                ret_prov = j
            Case Is = "Town"
                ret_twn = j
            Case Is = "Barangay"
                ret_bgy = j
            Case Is = "Discount/" & vbLf & "Markup"
                ret_dis = j
            Case Is = "Credit Line"
                ret_lin = j
            Case Is = "Credit Terms"
                ret_trm = j
            Case Is = "Status"
                ret_stat = j
            Case Is = "RD Salesman Code"
                ret_dsp = j
            Case Is = "RD Salesman Name"
                ret_man = j
        End Select
    Next
    
    If VarType(mcp) <> vbArray + vbVariant Then
        dist_code = UCase(InputBox("Enter 3 character RD code", "RD Code", "TFM"))
    Else
'        For i = LBound(mcp, 1) To UBound(mcp, 1)
'            If in_mcp.Exists(mcp(i, 2)) Then
'
'            Else
'                in_mcp.Add mcp(i, 2), 1
'            End If
'        Next
    End If
    
    
    For i = (m + 1) To UBound(cml, 1)
        
        'add only active customer status
        If cml(i, ret_stat) = "ACT" Then
            
            k = k + 1
                                
            ReDim Preserve temp(1 To 37, 1 To k)
            
            If VarType(mcp) = vbArray + vbVariant Then
                dsp_row = IsInArray(CStr(cml(i, ret_code)), mcp) 'IsInArray(UCase(CStr(cml(i, ret_code))), mcp)
                
                If dsp_row < 0 Then
                    k = k - 1
                    GoTo Next_i 'skip
                End If
                
                If dsp_row > 0 Then
                    temp(1, k) = mcp(dsp_row, 1)              'Distributor_Code
                    temp(2, k) = mcp(dsp_row, 11)       'Seller_Code
                    temp(3, k) = mcp(dsp_row, 12)       'Seller_Name
                    
                    'default service model of DSP
                    def_model = extract_dsp_service_model(CStr(mcp(dsp_row, 13)))
                    temp(14, k) = def_model(1)      'DSP Service Model Code
                    temp(15, k) = def_model(0)      'DSP Service Model Name
                    
                    'count wrong salesman
                    If mcp(dsp_row, 11) <> cml(i, ret_dsp) Then
                        err_d.Distributor_Code = mcp(1, 1)
                        err_d.Customer_Code = CStr(cml(i, ret_code)) 'UCase(CStr(cml(i, ret_code)))
                        err_d.Customer_Name = cml(i, ret_name)
                        err_d.cml_Salesman_Code = cml(i, ret_dsp)
                        err_d.cml_Salesman_Name = cml(i, ret_man)
                        err_d.mcp_Salesman_Code = mcp(dsp_row, 11)
                        err_d.mcp_Salesman_Name = mcp(dsp_row, 12)
                        
                        If Not err_dsp.Exists(err_d.Customer_Code & err_d.mcp_Salesman_Code) Then
                            err_dsp.Add err_d.Customer_Code & err_d.mcp_Salesman_Code, err_d
                        End If
                        
                    End If
                    
                Else
                    temp(1, k) = mcp(1, 1)            'Distributor_Code
                    temp(2, k) = cml(i, ret_dsp)       'Seller_Code
                    temp(3, k) = cml(i, ret_man)       'Seller_Name
                End If
            Else
                temp(1, k) = IIf(Len(cml(i, ret_rd)) < 5, cml(i, ret_rd), dist_code)         'Distributor_Code
                temp(2, k) = cml(i, ret_dsp)       'Seller_Code
                temp(3, k) = cml(i, ret_man)       'Seller_Name
            End If

                
            temp(4, k) = CStr(cml(i, ret_code)) 'UCase(CStr(cml(i, ret_code))) 'Retailer_Code
            temp(5, k) = cml(i, ret_name)        'Retailer_Name
            temp(6, k) = "GT"                   'Global_Channel_Code
            temp(7, k) = "General Trade"        'Global_Channel_Name
            
            cust_chan = get_channel(CStr(cml(i, ret_clas)))
            
            temp(8, k) = cust_chan(0)           'Customer_Class_Code
            
            If cust_chan(0) <> cml(i, ret_clas) Then
                err_c.Distributor_Code = temp(1, k)
                err_c.Customer_Code = temp(4, k) 'UCase(temp(4, k))
                err_c.Customer_Name = temp(5, k)
                err_c.Customer_Class = cml(i, ret_clas)
                err_c.Correct_Class = cust_chan(0)
                
                If Not err_chnl.Exists(err_c.Customer_Code) Then
                    err_chnl.Add err_c.Customer_Code, err_c
                End If
            End If
            
            temp(9, k) = cust_chan(1)           'Customer_Class_Name
            temp(10, k) = cust_chan(2)          'Segmentation_Code
            temp(11, k) = cust_chan(3)          'Segmentation_Name
            temp(12, k) = cust_chan(4)          'Channel_Code
            temp(13, k) = cust_chan(5)          'Channel_Name
            
            RTM_Model = get_rtm(CStr(cml(i, ret_type)))
            If RTM_Model(0) <> "" Then temp(14, k) = RTM_Model(0)      'Service_Model_Code
            If RTM_Model(1) <> "" Then temp(15, k) = RTM_Model(1)      'Service_Model_Name

            If dsp_row > 0 Then

                cust_r = get_rtm_code(check_service_model(CStr(mcp(dsp_row, 13))))

                If cust_r <> CStr(RTM_Model(0)) And _
                   (CStr(RTM_Model(0)) = "KCM" Or _
                   CStr(RTM_Model(0)) = "SEC" Or _
                   CStr(RTM_Model(0)) = "BTDT" Or _
                   CStr(RTM_Model(0)) = "EXT" Or _
                   CStr(RTM_Model(0)) = "OPM") Then
                    err_r.Distributor_Code = temp(1, k)
                    err_r.Customer_Code = temp(4, k) 'UCase(temp(4, k))
                    err_r.Customer_Name = temp(5, k)
                    err_r.Customer_RTM = RTM_Model(0)
                    err_r.Correct_RTM = cust_r
                    
                    If Not err_rtm.Exists(err_r.Customer_Code & err_r.Customer_RTM) Then
                        err_rtm.Add err_r.Customer_Code & err_r.Customer_RTM, err_r
                    End If
                End If
            End If
            
            cust_add = CStr(cml(i, ret_add))
            
            If cust_add = Empty Then
                no_address = no_address + 1
                cust_add = CStr(cml(i, ret_bgy)) & ", " & CStr(cml(i, ret_twn)) & ", " & CStr(cml(i, ret_prov))
            End If
            
                        sBrgy = UCase(CStr(cml(i, ret_bgy)))
            sTown = UCase(CStr(cml(i, ret_twn)))
            
            temp(16, k) = cust_add              'Address_Line1
            temp(17, k) = sBrgy                 'Address_Line2
            temp(18, k) = sTown                 'Address_Line3
            temp(19, k) = get_phone(CStr(cml(i, ret_num))) 'Contact_Phone_No
            temp(20, k) = IIf(cml(i, ret_own) = "", 0, cml(i, ret_own))       'Contact_Person
            temp(21, k) = ""                     'Lat_Position
            temp(22, k) = ""                     'Long_Position
            temp(23, k) = get_purtype(CStr(cml(i, ret_trm)), CStr(cml(i, ret_lin)))   'Purchase_Type
            
            If ret_prov > 0 Then
                sProv = UCase(cml(i, ret_prov))
                temp(24, k) = get_loc(sProv, sTown, sBrgy)
            Else
                sProv = UCase(CStr(getProvince(cml(i, ret_zip))))
                temp(24, k) = get_loc(sProv, sTown, sBrgy)
            End If
            
'            sBrgy = UCase(CStr(cml(i, ret_bgy)))
'            sTown = UCase(CStr(cml(i, ret_twn)))
'
'            temp(16, k) = cust_add 'cml(i, ret_add)            'Address_Line1
'            temp(17, k) = sTown 'cml(i, ret_twn)            'Address_Line2
'            'temp(18, k) = cml(i, ret_prov)            'Address_Line3
'            temp(19, k) = get_phone(CStr(cml(i, ret_num))) 'Contact_Phone_No
'            temp(20, k) = IIf(cml(i, ret_own) = "", 0, cml(i, ret_own))       'Contact_Person
'            temp(21, k) = ""                     'Lat_Position
'            temp(22, k) = ""                     'Long_Position
'            temp(23, k) = get_purtype(CStr(cml(i, ret_trm)))   'Purchase_Type
'
'            If ret_prov > 0 Then
'                sProv = UCase(cml(i, ret_prov))
'                temp(18, k) = sProv
'                temp(24, k) = get_loc(sProv, sTown, sBrgy)
'            Else
'                sProv = UCase(CStr(getProvince(cml(i, ret_zip))))
'                temp(18, k) = sProv
'                temp(24, k) = get_loc(sProv, sTown, sBrgy)
'            End If
'
            
            temp(25, k) = IIf(cml(i, ret_lin) = Empty, IIf(cml(i, ret_trm) = "Credit", 1, 0), cml(i, ret_lin)) 'Credit_Limit
            temp(26, k) = cml(i, ret_tin)            'TinNo
            'temp(27, k) = none                 'Discount_Type
            temp(28, k) = get_markup(CStr(cml(i, ret_dis)))    'Mark_up_Percentage check
            'temp(29, k) = none                 'Account_Code
            'temp(30, k) = none                 'Account_Name
            'temp(31, k) = none                 'Banner_Code
            'temp(32, k) = none                 'Banner_Name
            'temp(33, k) = none                 'Customer_Group_Code
            'temp(34, k) = none                 'Customer_Group_Name
            temp(35, k) = "REGULAR"             'Tax_Type
            temp(36, k) = 12                    'Tax_Rate
            'temp(37, k) = none                 'Flag
            
        End If

Next_i:
    Next
    
    ReDim Preserve temp2(1 To UBound(temp, 2), 1 To UBound(temp, 1))
    For i = LBound(temp, 2) To UBound(temp, 2)
        For j = LBound(temp, 1) To UBound(temp, 1)
            temp2(i, j) = temp(j, i)
        Next
    Next
    
    x = Array("Distributor_Code", "Seller_Code", "Seller_Name", "Retailer_Code", _
        "Retailer_Name", "Global_Channel_Code", "Global_Channel_Name", "Customer_Class_Code", _
        "Customer_Class_Name", "Segmentation_Code", "Segmentation_Name", "Channel_Code", _
        "Channel_Name", "Service_Model_Code", "Service_Model_Name", "Address_Line1", _
        "Address_Line2", "Address_Line3", "Contact_Phone_No", "Contact_Person", _
        "Lat_Position", "Long_Position", "Purchase_Type", "Location_Code", _
        "Credit_Limit", "TinNo", "Discount_Type", "Mark_up_Percentage", _
        "Account_Code", "Account_Name", "Banner_Code", "Banner_name", _
        "Customer_Group_Code", "Customer_Group_Name", "Tax_Type", _
        "Tax_Rate", "Flag")
        
    For j = LBound(x) To UBound(x)
        temp2(1, j + 1) = x(j)
    Next
    
    extract_cml = temp2
    
    'error text, if any
    If err_dsp.count > 0 Then errors_found = errors_found & vbCrLf & err_dsp.count & " wrong DSP."
    If err_chnl.count > 0 Then errors_found = errors_found & vbCrLf & err_chnl.count & " wrong Customer Class."
    If err_rtm.count > 0 Then errors_found = errors_found & vbCrLf & err_rtm.count & " wrong RTM model."
    If no_address > 0 Then errors_found = errors_found & vbCrLf & no_address & " no address (forced with location)."
    
End Function

Function clean_template()
'resets MCP sheet contents
    
    Dim LstRow As Long
    Dim FstRow As Long
    Dim brgy_fmla As String
    Dim freq_fmla As String
    Dim seq_fmla As String
    Dim day_fmla As String
    Dim this_wsh As Worksheet
    
    brgy_fmla = "=IFERROR(VLOOKUP([@Province]&[@[Town/City]]&[@Barangay],ref_brgy!C18:C20,3,FALSE),"""")"
    freq_fmla = "=IF(COUNTA(tblMCP[@[Mon]:[Sun]])*COUNTA(tblMCP[@[WK1]:[WK4]])=0,"""",""F""&COUNTA(tblMCP[@[Mon]:[Sun]])*COUNTA(tblMCP[@[WK1]:[WK4]]))"
    seq_fmla = "=IFERROR(CHOOSE([@Day],COUNTIFS(R6C11:RC11,RC11,R6C17:RC17,1),COUNTIFS(R6C11:RC11,RC11,R6C18:RC18,1),COUNTIFS(R6C11:RC11,RC11,R6C19:RC19,1),COUNTIFS(R6C11:RC11,RC11,R6C20:RC20,1),COUNTIFS(R6C11:RC11,RC11,R6C21:RC21,1),COUNTIFS(R6C11:RC11,RC11,R6C22:RC22,1),COUNTIFS(R6C11:RC11,RC11,R6C23:RC23,1)),"""")"
    day_fmla = "=IFERROR(MATCH(1,tblMCP[@[Mon]:[Sun]],0),"""")"
    
    Set this_wsh = ThisWorkbook.Sheets("MCP")
    
    With this_wsh
        
        If sheetProtect(this_wsh, 0) Then
        
            LstRow = .Cells(.Rows.count, 1).End(xlUp).Row
            FstRow = .Cells(LstRow, 1).End(xlUp).Row + 1
            
            .Range(.Cells(6, 1), .Cells(LstRow, 30)).ClearContents
            
            
            .Range(.Cells(6, 8), .Cells(LstRow, 8)).FormulaR1C1 = brgy_fmla
            .Range(.Cells(6, 16), .Cells(LstRow, 16)).FormulaR1C1 = freq_fmla
            .Range(.Cells(6, 28), .Cells(LstRow, 28)).FormulaR1C1 = seq_fmla
            .Range(.Cells(6, 29), .Cells(LstRow, 29)).FormulaR1C1 = day_fmla
            
            sheetProtect this_wsh, 1
            'xEnableOutlining this_wsh
            
        Else
            MsgBox "Please unprotect worksheet."
        End If
        
    End With
    
    Set this_wsh = Nothing
    
End Function

Sub run_clean()
'resets MCP consolidator
    Call clean_template
End Sub

Function check_error(v As Variant) As String
'returns error message

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim wrong_repcode As Long
    Dim wrong_add As Long
    Dim ref_brgy As Object
    Dim ref As Variant
    Dim w_add() As Variant
    Dim w_add2() As Variant
    Dim hdr As Variant
    
    RD_code = UCase(RD_code)
    
    Set ref_brgy = CreateObject("Scripting.Dictionary")
    
    With ThisWorkbook.Sheets("ref_brgy")
        ref = .Range("ref_brgy").value
    End With
    
    For i = LBound(ref, 1) To UBound(ref, 1)
        ref_brgy.Add ref(i, 1), 1
    Next
    
    For i = LBound(v, 1) To UBound(v, 1)
                        
        If RD_code = "TFM" Then
            If Left(v(i, 11), 3) <> "TFO" Then
                wrong_repcode = wrong_repcode + 1
            End If
        ElseIf RD_code = "SFOR" Then
            If Left(v(i, 11), 3) <> "SFR" Then
                wrong_repcode = wrong_repcode + 1
            End If
        ElseIf RD_code = "SMEI" Then
            If Left(v(i, 11), 3) <> "SME" Then
                wrong_repcode = wrong_repcode + 1
            End If
        ElseIf Left(v(i, 11), Len(RD_code)) <> RD_code Then
            wrong_repcode = wrong_repcode + 1
        End If
        
        
        'If v(i, 5) = Empty Or v(i, 6) = Empty Then
        If Not ref_brgy.Exists(UCase(v(i, 4) & v(i, 5) & v(i, 6))) Then
            wrong_add = wrong_add + 1
            ReDim Preserve w_add(1 To 10, 1 To wrong_add)
            w_add(1, wrong_add) = v(i, 1)
            w_add(2, wrong_add) = v(i, 2)
            w_add(3, wrong_add) = v(i, 3)
            w_add(4, wrong_add) = v(i, 4)
            w_add(5, wrong_add) = v(i, 5)
            w_add(6, wrong_add) = v(i, 6)
            w_add(7, wrong_add) = v(i, 7)
            w_add(8, wrong_add) = Empty
            w_add(9, wrong_add) = v(i, 11)
            w_add(10, wrong_add) = v(i, 12)
        End If
        
    Next
    
    If IsArrayAllocated(w_add) Then
        ReDim w_add2(LBound(w_add, 2) To UBound(w_add, 2) + 1, LBound(w_add, 1) To UBound(w_add, 1))
        hdr = Array("Distributor_Code", "Customer_Code", "Customer_Name", _
                    "Province", "Town", "Barangay", "Street", _
                    "Barangay Code", "Salesman_Code", "Salesman_Name")
        For i = 1 To UBound(w_add, 2) + 1
            For j = 1 To UBound(w_add, 1)
                 If i = 1 Then
                    w_add2(i, j) = hdr(j - 1)
                Else
                    w_add2(i, j) = w_add(j, i - 1)
                End If
            Next
        Next
        
        AddNew w_add2, "INCORRECT_ADDRESSES_IN_CML_" & w_add2(2, 1), "wrongAddress"
        
    End If
    
    check_error = "Found " & vbCrLf & _
        error_semantics(without_cml, "in MCP not in CML") & _
        error_semantics(inactive_cml, "in MCP inactive in CML") & _
        error_semantics(without_mcp, "in CML not in MCP") & _
        error_semantics(different_dsp, "with wrong assigned salesman in CML") & _
        error_semantics(different_rtm, "with wrong RTM Model in CML") & _
        error_semantics(z_dup, "are duplicates.") & _
        error_semantics(wrong_add, " incorrect customer address.", 0) & _
        error_semantics(wrong_repcode, "wrong salesman code.", 0) & _
        error_semantics(mcp_miss.count, "have no week schedule in MCP.") & _
        IIf((without_cml + inactive_cml + without_mcp + different_dsp + different_rtm + z_dup _
        + mcp_miss.count + wrong_add + wrong_repcode) = 0, "No error.", "")
    
End Function

Function check_service_model(serv As String) As String
'returns correct service model name

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim r As Long
    Dim x As Variant
    Dim y As Variant
    Dim d As Variant
    Dim newString As String
    
    k = 0
    
    For i = LBound(mod_names) To UBound(mod_names)
        If InStr(1, serv, mod_names(i, 1)) Then
            check_service_model = serv
            Exit Function
        ElseIf mod_names(i, 1) = serv Then
            check_service_model = serv
            Exit Function
        ElseIf LCase(Left(mod_names(i, 1), Len(serv))) = LCase(serv) Then
            check_service_model = mod_names(i, 1)
            Exit Function
        ElseIf LCase(Left(mod_names(i, 1), Len(serv))) = LCase(Replace(serv, "-", "")) Then
            check_service_model = mod_names(i, 1)
            Exit Function
        ElseIf LCase(Left(serv, 2)) = "ex" Then
            check_service_model = "Extruck Accounts (EXT)"
            Exit Function
        End If
    Next
    
    newString = Replace(serv, "-", " ")
    x = Split(newString, " ")
    For Each d In x
        For i = LBound(mod_names) To UBound(mod_names)
            If IsInStr(CStr(d), CStr(mod_names(i, 1))) Then
                k = k + 1
                If k > 1 Then
                    check_service_model = mod_names(i, 1)
                    Exit Function
                End If
            End If
        Next
    Next
    
    check_service_model = serv

End Function
