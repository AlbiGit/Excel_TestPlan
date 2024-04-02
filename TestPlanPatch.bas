Attribute VB_Name = "Module1"
Sub test_plan_format()
Dim names
Dim Row As Integer
Dim Column As Integer
Dim tfs() As String
Dim TfsRows As String
Dim dt As String

On Error GoTo exit_sub
Set Rng = Application.InputBox(Title:="Test plan format", Prompt:="Select the TFSs", Type:=8)

On Error GoTo 0

If Rng(1) = "" Then Exit Sub
Set Move = Rng
ActiveSheet.Name = "Summary"

names = "/"
For N = 2 To Worksheets.Count
    names = names + CStr(Worksheets(N).Name) + "/"
Next N



For i = Rng(1).Row To Rng(1).Cells(Rng.Rows.Count, Rng.Columns.Count).Row
    
    sh_name = Cells(i, Rng(1).Column).Value
    If InStr(names, CStr(sh_name)) = 0 Then
        tfs = Split(CStr(sh_name), ":")
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = tfs(0)
        Sheets("Summary").Activate
        Sheets("Summary").Range(Cells(i, Rng(1).Column), Cells(i, Rng(1).Cells(Rng.Rows.Count, Rng.Columns.Count).Column)).Copy
        Sheets(tfs(0)).Activate
        Range("D1").Select
        ActiveSheet.Paste
        Columns("D").ColumnWidth = 100
        ActiveSheet.Range("A1").Value = tfs(0)
        ActiveSheet.Range("B1").Value = tfs(1)
        ActiveSheet.Range("C1").Select
        ActiveSheet.Range("C1").Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:="'Summary'!C1", TextToDisplay:="Top"
        ActiveSheet.Range("A1:B1").Borders.LineStyle = xlNone
        ActiveSheet.Range("A1:B1").WrapText = False
        ActiveSheet.Range("B4:B12").WrapText = True
        For Row = 3 To 12
            ActiveSheet.Range("A" & Row & ":D" & Row).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        Next
        ActiveSheet.Range("A3:A12").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range("B3:B12").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range("C3:C12").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range("D3:D12").BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range("A3:D12").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        ActiveSheet.Range("A3").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        ActiveSheet.Range("B3").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        ActiveSheet.Range("C3").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        ActiveSheet.Range("D3").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        ActiveSheet.Range("A3").Value = "No."
        ActiveSheet.Range("B3").Value = "Test "
        ActiveSheet.Range("C3").Value = "P/F"
        ActiveSheet.Range("D3").Value = "Notes"
        Columns("A").ColumnWidth = 15
        Columns("B").ColumnWidth = 75
        Rows("1").RowHeight = 15.75
        ActiveSheet.Range("A3:D12").HorizontalAlignment = xlCenter
        ActiveSheet.Range("A3:D12").VerticalAlignment = xlCenter
        Sheets("Summary").Activate
        ActiveSheet.Hyperlinks.Add Range(Cells(i, Rng(1).Column), Cells(i, Rng(1).Column)), Address:="", SubAddress:="'" & tfs(0) & "'!A2"
        
    End If

Next i
    Move.Cut
    Sheets("Summary").Activate
    Set PasteCell = ActiveSheet.Range("A10")
    If IsEmpty(PasteCell.Value) Then
        ActiveSheet.Range("A10").Select
        ActiveSheet.Paste
    Else
    i = 10
    While (IsEmpty(PasteCell)) = False
        i = i + 1
        Set PasteCell = ActiveSheet.Range("A" & CStr(i))
    Wend
    ActiveSheet.Range("A" + CStr(i)).Select
    ActiveSheet.Paste
    End If
    
    Application.CutCopyMode = False
    ActiveSheet.Range("A9").Value = "TFS"
    ActiveSheet.Range("B9").Value = "Status"
    ActiveSheet.Range("C9").Value = "Notes"
    ActiveSheet.Range("D9").Value = "Branch"
    TfsRows = "A" & CStr(Worksheets.Count + 10)
    SummeryRows = "D" & CStr(Worksheets.Count + 10)
    For i = 9 To (Worksheets.Count + 10)
        CellA = "A" & CStr(i)
        CellB = "B" & CStr(i)
        CellC = "C" & CStr(i)
        CellD = "D" & CStr(i)
        ActiveSheet.Range(CellA).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range(CellB).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range(CellC).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range(CellD).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    Next i
    
    For i = 2 To 7
        CellA = "A" & CStr(i)
        CellB = "B" & CStr(i)
        ActiveSheet.Range(CellA).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        ActiveSheet.Range(CellB).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    
    Next i
    ActiveSheet.Range("A9" + ":" + TfsRows).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    ActiveSheet.Range("A9" + ":" + SummeryRows).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    ActiveSheet.Range("A2:B7").BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    ActiveSheet.Range("A2").Value = "Sprint"
    ActiveSheet.Range("A3").Value = "Fabware"
    ActiveSheet.Range("A4").Value = "PTO"
    ActiveSheet.Range("A5").Value = "PLC"
    ActiveSheet.Range("A6").Value = "FabXRF"
    ActiveSheet.Range("A7").Value = "JVXRR"
    dt = Format(Date, "mm-yy")
    ActiveSheet.Range("B2").Value = dt
    Columns("A").ColumnWidth = 75
    ActiveSheet.Range("A2:B7").HorizontalAlignment = xlCenter
    ActiveSheet.Range("A2:B7").VerticalAlignment = xlCenter
Exit Sub

exit_sub:
    Exit Sub
    
End Sub
