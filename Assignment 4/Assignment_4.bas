Attribute VB_Name = "Module1"
Option Explicit

' NOTE: For highlighting, use .ColorIndex = 4
' For example, Range("A1").Interior.ColorIndex = 4 would color cell A1 green

Sub HighlightRows()
'Place your code here
Dim nr As Integer, i As Integer, id As String, idkey As Integer, ws As Worksheet

Set ws = ThisWorkbook.Sheets("Data")
nr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
id = ws.Range("identifier").Value
idkey = ws.Range("key").Value

For i = 2 To nr
    If Identifier(ws.Cells(i, 1).Value) = id And Key(ws.Cells(i, 1).Value) = idkey Then
        ws.Range(ws.Cells(i, 1), ws.Cells(i, 3)).Interior.ColorIndex = 4
    End If
    
Next i
    


End Sub

Sub Example()
' This is just to show how the Identifier and Key functions below can be utilized in VBA code
Dim id As String
id = "Y4-824X"
MsgBox "The identifier is " & Identifier(id) & " and the key is " & Key(id)
End Sub

Function Identifier(id As String) As String
Identifier = Left(id, 1)
End Function

Function Key(id As String) As Integer
Key = Left(Mid(id, 4, 4), 1)
End Function

Sub Reset()
' Obtained through a macro recording:
With Cells.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
End Sub
