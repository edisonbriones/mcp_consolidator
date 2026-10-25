Attribute VB_Name = "Module2"
Option Explicit
'reusable functions

Function xEnableOutlining(ws As Worksheet) As Boolean
'enables grouped cells to expand and collapse
    Dim sh As Worksheet
    
    If ws Is Nothing Then
        Set sh = ThisWorkbook.Sheets("Summary")
    Else
        Set sh = ws
    End If
    
    With sh
        .Protect _
                AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, _
                AllowInsertingRows:=True, _
                AllowDeletingRows:=True, _
                AllowSorting:=True, _
                AllowUsingPivotTables:=True, _
                DrawingObjects:=True, _
                Contents:=True, _
                Scenarios:=True, _
                AllowFiltering:=True, Password:="", Userinterfaceonly:=True
        .EnableOutlining = True
    End With
End Function

Function PickSomeFiles() As Variant
'asks for user file input

    Dim fdFileDialog As FileDialog
    Dim SelectedItemsCount As Long
    Dim msg As String
    Dim r As Object
    Dim i As Long

    Set r = CreateObject("Scripting.Dictionary")
    
    Set fdFileDialog = Application.FileDialog(msoFileDialogOpen)
    With fdFileDialog
        .Filters.Clear
        .Filters.Add "Microsoft Excel Files (*.xls*), *.xls*", "*.xls*"
        .FilterIndex = 1
        .InitialView = msoFileDialogViewDetails
        .InitialFileName = ThisWorkbook.Path
        .Title = "Please choose files"
        .ButtonName = "open"
        .AllowMultiSelect = True
        .Show
        If .SelectedItems.count = 0 Then
            MsgBox "No files selected"
            Exit Function
        End If
        SelectedItemsCount = .SelectedItems.count
        For i = 1 To SelectedItemsCount
            msg = msg & vbCrLf & i & " - " & .SelectedItems(i)
            r.Add .SelectedItems(i), 1
        Next i
        
        MsgBox "You selected: " & msg
        
        Set PickSomeFiles = r
    End With
    
End Function

Function select_file(mult As Boolean) As Variant
'asks for user file input

    Dim f_open As Variant
    
    f_open = Application _
        .GetOpenFilename("Microsoft Excel Files (*.xls*), *.xls*", 1, _
        "Microsoft Excel File to Import", , mult)
    
    If VarType(f_open) = vbBoolean Then
        MsgBox "No file was selected."
        select_file = ""
    Else
        select_file = f_open
    End If
       
End Function

Function open_workbook(d_file As Variant) As Workbook
'returns opened workbook
    Set open_workbook = Workbooks.Open(d_file, UpdateLinks:=0)
End Function

Function IsArrayAllocated(arr As Variant) As Boolean
'returns True if array has value

        On Error Resume Next
        IsArrayAllocated = IsArray(arr) And _
                           Not IsError(LBound(arr, 1)) And _
                           LBound(arr, 1) <= UBound(arr, 1)
End Function

Function SwitchOff(bSwitchOff As Boolean)
'switches background processes on or off for faster execution time

    Dim Scrn_Status As Boolean
    Dim Evnt_Status As Boolean
    Dim Alrt_Status As Boolean
    
    With Application
        If bSwitchOff Then
            
            'save
            'Calc_Status = .Calculation

            'Turn off
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableAnimations = False
            .EnableEvents = False
            .DisplayAlerts = False
        Else
        
             'Turn on
             .Calculation = xlAutomatic 'If .Calculation <> Calc_Status And Calc_Status <> 0 Then .Calculation = Calc_Status
             .ScreenUpdating = True
             .EnableAnimations = True
             .EnableEvents = True
             .DisplayAlerts = True
        End If
  End With
End Function

Function sheetProtect(ws As Worksheet, n As Byte) As Boolean
'protect sheet

    With ws
        If n = 0 Then
            If .ProtectContents Then
                On Error GoTo errorLabel
                .Unprotect ""
                
            End If
        ElseIf n = 1 Then
            .Protect _
                AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, _
                AllowInsertingRows:=True, _
                AllowDeletingRows:=True, _
                AllowSorting:=True, _
                AllowUsingPivotTables:=True, _
                DrawingObjects:=True, _
                Contents:=True, _
                Scenarios:=True, _
                AllowFiltering:=True
            .EnableOutlining = True
        End If

    sheetProtect = True
    
    End With
    
    Exit Function
    
errorLabel:
    If Err.Number = 1004 Then
        MsgBox "Please unprotect worksheet."
        sheetProtect = False
    End If
End Function

Function GetNumericValue(some_text As String) As String
'returns numbers from some text

    Dim RE As Object
    Dim x As String
    
    Set RE = CreateObject("vbscript.regexp")
    RE.IgnoreCase = True
    RE.Global = True
    RE.Pattern = "[^.\d]"

    GetNumericValue = RE.Replace(some_text, "")
    
End Function

Function IsInArrayStr(stringToBeFound As String, arr As Variant) As Boolean
'returns True if string is in array

    IsInArrayStr = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Function IsInStr(stringToBeFound As String, txt As String) As Boolean
'returns True if character is in string

    Dim x As Variant
    stringToBeFound = LCase(stringToBeFound)
    txt = LCase(txt)
    x = Split(txt)
    IsInStr = (UBound(Filter(x, stringToBeFound)) > -1)
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Long
'returns row location of string in array

  Dim i As Long
  Dim j As Long
  
  ' default return value if value not found in array
  IsInArray = -1

  For i = LBound(arr, 1) To UBound(arr, 1)
    'For j = LBound(arr, 2) To UBound(arr, 2)
        'If arr(i, j) = stringToBeFound Then
        j = 2
        If arr(i, j) = stringToBeFound Then
            IsInArray = i
            Exit Function
        End If
    'Next j
  Next i
End Function

Function AddNew(data_table As Variant, fname As String, Optional fname2 As String)
'returns new workbook with the data

    Dim NewBook As Workbook
    Dim NewSheet As Worksheet
    Dim rng As Range
    Dim relativePath As String
    
    Set NewBook = Workbooks.Add
    
    With NewBook
        
        With .Sheets(1)
            Set rng = .Range(.Cells(2, 1), .Cells(UBound(data_table, 1) + 1, UBound(data_table, 2)))
            rng.NumberFormat = "@"
            .Cells(1, 1).Value2 = fname
            rng.Value2 = data_table
        End With
        
        
        If fname2 = "" Then fname2 = fname
        
        'save
        .Title = fname2 'You can modify this value.
        .Subject = fname2 'You can modify this value.
        
        relativePath = ThisWorkbook.Path & "\" & fname2 & ".xlsx"
        
        On Error Resume Next
        .SaveAs Filename:=relativePath
        
    End With
    
End Function

Function getProvince(zipcode As Variant) As String
'returns the province from postal code input

    Dim arr As Variant
    Dim rng As Range
    Dim zip As Long
    Dim prov As String
     
    On Error Resume Next
    If VarType(zipcode) = 0 Then Exit Function
       
    zip = CLng(zipcode)
    
    If IsArrayAllocated(zip_ref) = False Then zip_ref = ThisWorkbook.Worksheets("ref_brgy").Range("zip").value
    
    prov = WorksheetFunction.VLookup(zip, zip_ref, 3, True)
    
    getProvince = prov
    
End Function

Function extract_dsp_service_model(txt As String) As Variant
'returns salesman service model code and name from text

    Dim a As Variant
    a = Split(txt, " (")
    a(1) = Left(CStr(a(1)), Len(CStr(a(1))) - 1)
    
    If a(0) = "EXT" Then
        a(1) = "Extruck Accounts"
    ElseIf a(0) = "OPM" Then
        a(1) = "On Premise Management"
    End If
    
    extract_dsp_service_model = a
End Function

Function get_rtm_code(txt As String) As String
'returns rtm model code
    On Error Resume Next
    get_rtm_code = Application.WorksheetFunction.VLookup(txt, serv_model, 2, False)
End Function

Function get_loc(prov As String, Town As String, brgy As String) As String
'returns barangay code

    Dim rng As Range
    Dim lcode As String
    Dim loc_ As String
    Dim i As Long
    Dim j As Long
    
    loc_ = UCase(prov) & UCase(Town) & UCase(brgy)
    
    If IsArrayAllocated(loc_codes) = False Then
        
        Set dict = CreateObject("Scripting.Dictionary")
        loc_codes = ThisWorkbook.Worksheets("ref_brgy").Range(["loc"]).value
        For i = LBound(loc_codes, 1) To UBound(loc_codes, 1)
            dict(loc_codes(i, 1)) = loc_codes(i, 2)
        Next
        
    End If
    
    If dict.Exists(loc_) Then lcode = dict(loc_) Else lcode = loc_codes(1, 2)
        
    get_loc = lcode
    
End Function

Function get_rtm(v As String) As Variant
'returns rtm model code and name of lookup value

    Dim x(1) As String
    
    If IsArrayAllocated(rtm_mod) = False Then rtm_mod = Range("rtm_model").value
    
    On Error Resume Next
    x(0) = WorksheetFunction.VLookup(v, rtm_mod, 2, False)
    x(1) = WorksheetFunction.VLookup(v, rtm_mod, 3, False)
    
    get_rtm = x
    
End Function

Function get_channel(v As String) As Variant
'returns correct customer classification

    Dim x(5) As String
    
    With ThisWorkbook.Worksheets("ref")
        If VarType(old_codes) <> vbArray + vbVariant Then old_codes = .Range("old_codes").value
        If VarType(new_codes) <> vbArray + vbVariant Then new_codes = .Range("new_codes").value
    End With
    
    On Error Resume Next
    x(0) = WorksheetFunction.VLookup(v, old_codes, 2, False)
    x(1) = WorksheetFunction.VLookup(x(0), new_codes, 2, False)
    x(2) = WorksheetFunction.VLookup(x(0), new_codes, 3, False)
    x(3) = WorksheetFunction.VLookup(x(0), new_codes, 4, False)
    x(4) = WorksheetFunction.VLookup(x(0), new_codes, 5, False)
    x(5) = WorksheetFunction.VLookup(x(0), new_codes, 6, False)
    
    get_channel = x
    
End Function

Function get_purtype(v As String, line As String) As String
'returns Cash or Credit
        
    If InStr("COD", v) Then
        get_purtype = "Cash"
    ElseIf InStr("PDC", v) Then
        get_purtype = "Credit"
    ElseIf InStr("CIA", v) Then
        get_purtype = "Cash"
    ElseIf InStr("CWO", v) Then
        get_purtype = "Cash"
    ElseIf line * 1 > 0 Then
        get_purtype = "Credit"
    Else
        get_purtype = "Cash"
    End If

End Function

Function get_markup(v As String) As Double
'returns mark up or discount percentage

    Dim x As Variant
    
    On Error Resume Next
    If v = "" Then get_markup = 0: Exit Function
    
    x = Split(v, " ")
    If IsInArrayStr("discount", x) = True Then
        get_markup = GetNumericValue(v)
    Else
        get_markup = -(CDbl(GetNumericValue(v)))
    End If

End Function

Function get_phone(v As String) As String
'returns valid phone number

    If v = Empty Then
        get_phone = 0
        Exit Function
    ElseIf IsNumeric(v) = False Then
        get_phone = 0
    Else
        get_phone = Left(GetNumericValue(v), 15)
    End If
End Function

Function error_semantics(count As Long, msg As String, Optional cust As Boolean = True) As String
'returns grammatically corrected sentence

    If count = 0 Then
        error_semantics = ""
    ElseIf cust = False Then
        error_semantics = count & msg & vbCrLf
    ElseIf count > 1 Then
        error_semantics = count & " customers " & msg & vbCrLf
    Else
        error_semantics = count & " customer " & msg & vbCrLf
    End If
End Function


Function get_TE_seg_code(txt As String) As String
'returns segmentation from TE "RETAIL ENVIRONMENT"

    If IsArrayAllocated(seg_name) = False Then seg_name = ThisWorkbook.Worksheets("ref").Range("$Q$2:$Q$26")
    If IsArrayAllocated(seg_code) = False Then seg_code = ThisWorkbook.Worksheets("ref").Range("$P$2:$P$26")
    
    On Error Resume Next
    get_TE_seg_code = WorksheetFunction.Index(seg_code, _
                        WorksheetFunction.Match(txt, seg_name, 0), 1)

End Function

Function get_TE_ch_code(txt As String) As String
'returns Channel Type from TE "KEY ACCOUNT"

    If IsArrayAllocated(ch_name) = False Then ch_name = ThisWorkbook.Worksheets("ref").Range("$S$2:$S$26")
    If IsArrayAllocated(ch_code) = False Then ch_code = ThisWorkbook.Worksheets("ref").Range("$R$2:$R$26")
    
    On Error Resume Next
    get_TE_ch_code = WorksheetFunction.Index(ch_code, _
                        WorksheetFunction.Match(txt, ch_name, 0), 1)

End Function

Function get_TE_rtm(txt As String) As String
'returns customer assigned Salesman Service Model Code and Name

    Dim a As Variant
    a = Split(txt, "(")
    get_TE_rtm = a(0)

End Function

Function get_TE_class_code(txt As String) As String
'returns customer class code of TE "PARTY_CLASSIFICATION_DESCRIPTION"

    If IsArrayAllocated(cust_name) = False Then cust_name = ThisWorkbook.Worksheets("ref").Range("$O$2:$O$26")
    If IsArrayAllocated(cust_code) = False Then cust_code = ThisWorkbook.Worksheets("ref").Range("$N$2:$N$26")
    
    On Error Resume Next
    get_TE_class_code = WorksheetFunction.Index(cust_code, _
                        WorksheetFunction.Match(txt, cust_name, 0), 1)

End Function

Function get_rd_code() As String
'asks user for RD code

    If RD_code = Empty Then
        RD_code = InputBox("Kindly put the distributor code", "Distributor Code", "TFM")
    End If
    
    get_rd_code = UCase(RD_code)
    
End Function


