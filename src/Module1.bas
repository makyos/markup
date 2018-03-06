Attribute VB_Name = "Module1"
Function tagTABLE(aStr As String)
    tagTABLE = "<table>" & aStr & "</table>"
End Function

Function tagTR(aStr As String)
    tagTR = "<tr>" & aStr & "</tr>"
End Function

Private Function tagTD(aStr As String)
    tagTD = "<td>" & aStr & "</td>"
End Function

Sub CopyHTML()

    ''' クリップボード操作のため、
    ''' 参照設定が必要 "c:\Windows\System32\FM20.DLL"
    
    Dim CB As New DataObject

    Dim retStr As String
    Dim bufStr As String
    
    Dim aRng As Range
    Dim lastRow As Integer

    lastRow = Selection(1).Row
    
    For Each aRng In Selection
        
        If aRng.Row > lastRow Then
            retStr = retStr & tagTR(bufStr) & vbCrLf
            bufStr = ""
        End If
        
        bufStr = bufStr & tagTD(aRng.Value)
        lastRow = aRng.Row
    
    Next
    
    retStr = retStr & tagTR(bufStr) & vbCrLf
    
    Debug.Print tagTABLE(vbCrLf & retStr) & vbCrLf
    With CB
        .SetText tagTABLE(vbCrLf & retStr) & vbCrLf
        .PutInClipboard
    End With

End Sub


Sub Auto_Open()

    Set bar = Application.CommandBars("Cell")
    bar.Reset
    
    Set contextmemu = bar.Controls.Add(Before:=1, temporary:=True)
    contextmemu.Caption = "CopyHTML"
    contextmemu.OnAction = "CopyHTML"
    contextmemu.FaceId = 351

End Sub


