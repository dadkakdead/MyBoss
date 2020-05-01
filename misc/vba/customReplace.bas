'2020-05-01 by NT: File added to repo to make GitHub show repo content as 100% VBA

Attribute VB_Name = "customReplace.bas"
' Substitute a string inside the textrange preserving the initial formating
Sub customReplace(shp As Shape, stringToSearch As String, stringToPaste As String)
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            Set trFoundText = shp.TextFrame.TextRange.Find(stringToSearch)
            If Not (trFoundText Is Nothing) Then
                m = shp.TextFrame.TextRange.Find(stringToSearch).Characters.Start
                shp.TextFrame.TextRange.Characters(m).InsertBefore (stringToPaste)
                shp.TextFrame.TextRange.Find(stringToSearch).Delete
            End If
        End If
    End If
End Sub
