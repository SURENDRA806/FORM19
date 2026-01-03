Sub Generate_FORM19()

    Dim xlApp As Object, wb As Object
    Dim wsH As Object, wsS As Object
    Dim wdDoc As Document
    Dim i As Long, lastRow As Long, r As Long

    Set xlApp = CreateObject("Excel.Application")
    Set wb = xlApp.Workbooks.Open(ThisDocument.Path & "\data\form19_data.xlsx")
    Set wsH = wb.Sheets("Header")
    Set wsS = wb.Sheets("Schedule")

    Set wdDoc = Documents.Open(ThisDocument.Path & "\template\FORM_19_Template.docx")

    ' Header bookmarks
    wdDoc.Bookmarks("NotificationNo").Range.Text = wsH.Range("C2").Value
    wdDoc.Bookmarks("Mandal").Range.Text = wsH.Range("A2").Value
    wdDoc.Bookmarks("Village").Range.Text = wsH.Range("B2").Value
    wdDoc.Bookmarks("SurveyDate").Range.Text = wsH.Range("D2").Value
    wdDoc.Bookmarks("SurveyTime").Range.Text = wsH.Range("E2").Value

    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(-4162).Row

    r = 2 ' Word table row (after header)
    For i = 2 To lastRow
        wdDoc.Tables(1).Rows.Add
        wdDoc.Tables(1).Cell(r, 1).Range.Text = wsS.Cells(i, 1).Value
        wdDoc.Tables(1).Cell(r, 2).Range.Text = wsS.Cells(i, 2).Value
        wdDoc.Tables(1).Cell(r, 3).Range.Text = wsS.Cells(i, 3).Value
        wdDoc.Tables(1).Cell(r, 4).Range.Text = wsS.Cells(i, 4).Value
        wdDoc.Tables(1).Cell(r, 5).Range.Text = wsS.Cells(i, 5).Value
        wdDoc.Tables(1).Cell(r, 6).Range.Text = wsS.Cells(i, 6).Value
        r = r + 1
    Next i

    wdDoc.SaveAs2 ThisDocument.Path & "\output\Form19_" & wsH.Range("B2").Value & ".docx"

    wb.Close False
    xlApp.Quit

    MsgBox "FORM-19 generated successfully"

End Sub
