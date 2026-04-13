Sub FixCaptionWithSpace()

    Dim fld As Field
    Dim rng As Range
    Dim txt As String
    
    For Each fld In ActiveDocument.Fields
        
        txt = fld.Code.Text
        
        If InStr(txt, "SEQ 图") > 0 Or InStr(txt, "SEQ 表") > 0 Then
            
            Set rng = fld.result
            rng.Collapse wdCollapseEnd
            
            ' 删除后面所有空白
            Do While rng.Characters.count > 0
                If rng.Characters.First.Text = " " _
                Or rng.Characters.First.Text = Chr(160) _
                Or rng.Characters.First.Text = vbTab Then
                    
                    rng.Characters.First.Delete
                Else
                    Exit Do
                End If
            Loop
            
         
            rng.InsertAfter " "
            
        End If
        
    Next fld

End Sub
