Option Explicit

Sub obtain_document_properties()

Dim intCounter As Integer
Dim strName As String
Dim docProperties As DocumentProperty
Dim wkbSheet As Worksheet

strName = "doc_properties"
intCounter = 1

For Each wkbSheet In ThisWorkbook.Worksheets
  If Worksheets(Worksheets.Count).Name = strName Then
    Exit For
  ElseIf Not wkbSheet.Name = strName Then
  ' Add an extra WorkSheet for the document's properties
    Worksheets.Add after:=Worksheets(Worksheets.Count)
    ThisWorkbook.Worksheets(Worksheets.Count).Name = strName
    Exit For
  End If
Next wkbSheet

For Each docProperties In ActiveWorkbook.BuiltinDocumentProperties
  With ThisWorkbook.Worksheets(strName)
    .Cells(intCounter, 1).Value = docProperties.Name
    On Error Resume Next
    .Cells(intCounter, 2).Value = docProperties.Value
    intCounter = intCounter + 1
  End With
Next docProperties

ThisWorkbook.Worksheets(strName).Columns.AutoFit

End Sub
