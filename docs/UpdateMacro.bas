' BORPATHAR BLF — Update button macro (direct pull, no Power Query).
' Reads data.xlsx (next to this master) and drops each tab's values into
' sheets named data_<TabName> in this workbook. Create those landing sheets
' once (data_KPIs, data_GreenLeaf, data_Auction, data_Production,
' data_Expenses, data_Economics, data_YoY_Expenses). Save master as .xlsm.
' Assign UpdateData to a button. Mac: allow file access if prompted.

Sub UpdateData()
    Dim src As Workbook, dst As Workbook, s As Worksheet, d As Worksheet
    Dim p As String, names As Variant, i As Integer, nm As String
    Application.ScreenUpdating = False
    Set dst = ThisWorkbook
    p = dst.Path & Application.PathSeparator & "data.xlsx"
    If Dir(p) = "" Then
        MsgBox "data.xlsx not found next to this file. Run RUN.command first.", vbExclamation
        Application.ScreenUpdating = True: Exit Sub
    End If
    ' Mac file-access (ignored on Windows)
    On Error Resume Next
    GrantAccessToMultipleFiles Array(p)
    On Error GoTo 0
    Set src = Workbooks.Open(Filename:=p, ReadOnly:=True)
    names = Array("KPIs", "GreenLeaf", "Auction", "Production", "Expenses", "Economics", "YoY_Expenses")
    For i = LBound(names) To UBound(names)
        nm = names(i)
        Set s = Nothing: Set d = Nothing
        On Error Resume Next
        Set s = src.Worksheets(nm)
        Set d = dst.Worksheets("data_" & nm)
        On Error GoTo 0
        If Not s Is Nothing And Not d Is Nothing Then
            d.Cells.Clear
            s.UsedRange.Copy
            d.Range("A1").PasteSpecial xlPasteValues
        End If
    Next i
    Application.CutCopyMode = False
    src.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Dashboard updated.", vbInformation
End Sub
