' This macro pulls data from multiple selected Excel files and processes it
Option Explicit

Sub pulldata()

    ' Disable alerts and screen updating for performance improvement
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim ofile As Variant
    Dim wk As Workbook
    Dim cnt As Long

    ' Open file dialog to select multiple Excel files
    ofile = Application.GetOpenFilename(filefilter:=".xlsx,(.xlsx)", MultiSelect:=True)
    
    On Error GoTo handle:
    
    ' Loop through selected files and copy data
    For cnt = 1 To UBound(ofile)
        Set wk = Workbooks.Open(ofile(cnt))
        
        ' Clear existing data in the destination sheet
        shraw.Cells.Clear
        
        ' Copy data from "sheet1" of the opened workbook
        wk.Sheets("sheet1").Cells.Copy
        shraw.Cells.PasteSpecial (xlPasteValuesAndNumberFormats)
        
        ' Close the workbook after copying data
        wk.Close
    Next cnt
    
    ' Select the test sheet after processing
    shtest.Select
    
handle:
    ' Restore alerts and screen updating
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub

End Sub

' This macro creates a table by filtering and organizing data based on predefined conditions
Sub createtable()

    ' Disable alerts and screen updating for performance improvement
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim arow As Long       ' Last row for EdgeID sheet
    Dim a As Long
    Dim lacolumn As Long   ' Last column for raw data sheet
    Dim nrow As Long       ' Last row for raw data sheet
    Dim trow As Long       ' Last row for new sheet
    Dim c As Long
    Dim supply1 As Variant
    Dim status1 As Variant
    Dim i As Long

    ' Determine the last row in different sheets
    arow = shEdgeID.Range("a" & Rows.Count).End(xlUp).Row
    nrow = shraw.Range("a" & Rows.Count).End(xlUp).Row
    trow = shtest.Range("a" & Rows.Count).End(xlUp).Row + 1
    
    ' Clear existing data in the test sheet
    shtest.Rows("3:" & trow).Delete

    ' Retrieve values from dashboard sheet
    supply1 = shDash.supplyx
    status1 = shDash.statusx

    lacolumn = shraw.Cells(1, Columns.Count).End(xlToLeft).Column

    i = 2
    
    ' Loop through raw data and match against EdgeID conditions
    For a = 2 To nrow
        For c = 2 To arow
            If shraw.Range("a" & a) = shEdgeID.Range("token").Cells(c, 1) Then
                If supply1 = shEdgeID.Range("supplier").Cells(c, 1) Then
                    If status1 = shEdgeID.Range("status").Cells(c, 1) Then
                        
                        ' Copy header and row data to the test sheet
                        shraw.Range("a1", "beb1").Copy
                        shtest.Range("b1").PasteSpecial xlPasteValuesAndNumberFormats
                        
                        shraw.Range("a" & a, "bdz" & a).Copy
                        shtest.Range("b" & i, "beb" & i).PasteSpecial xlPasteValuesAndNumberFormats
                        
                        i = i + 1
                    End If
                End If
            End If
        Next c
    Next a

    ' Notify user to verify data before proceeding
    MsgBox "Data created, please vlookup IDs in column A before creating file", vbCritical

    ' Select the test sheet
    shtest.Select
    
    ' Restore alerts and screen updating
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

' This macro generates a report by copying processed data into a new workbook
Sub suppreport()

    ' Disable alerts and screen updating for performance improvement
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wk As Workbook
    Dim wkpath As Variant

    ' Create a new workbook
    Set wk = Workbooks.Add
    
    ' Define the save path for the new report
    wkpath = ThisWorkbook.Path & "\" & shDash.supplyx & "_" & shEdgeID.Range("a2")
    
    ' Copy processed data to the new workbook
    shtest.Cells.Copy
    wk.Sheets("sheet1").PasteSpecial xlPasteValuesAndNumberFormats
    
    ' Remove column B from the new sheet
    wk.Sheets("sheet1").Columns(2).Delete
    
    ' Save and close the new workbook
    wk.SaveAs wkpath & ".xlsx"
    wk.Close
    
    ' Notify user of successful file creation
    MsgBox "File created in " & wkpath
    
    ' Restore alerts and screen updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
