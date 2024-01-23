Sub Autofit_Listbox1()
Dim WS As Worksheet
Dim LS, LastColumn, i As Long
Dim objek As String

Set WS = ThisWorkbook.Sheets("Data")
LS = WS.Range("A" & Rows.Count).End(xlUp).Row
objek = "userform1.listbox1_"
WS.Cells.EntireColumn.AutoFit

    With UserForm1.ListBox1
        .ColumnCount = 13
        .ColumnWidths = ""
        For i = 1 To 15
            If i > 20 Then
                .ColumnWidths = .ColumnWidths & IIf(i > 1, ";", "") & 0
            Else
                .ColumnWidths = .ColumnWidths & IIf(i > 1, ";", "") & WS.Cells(1, i).Width
            End If
        Next i
    End With

UserForm1.ListBox1.RowSource = WS.Range("A2:E" & LS + 1).Address(External:=True)
End Sub
