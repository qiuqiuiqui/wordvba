Sub ReplaceMathTypeByHyperlink()
    Dim rng As Range
    Dim fld As Field
    Dim codeText As String
    Dim labelName As String
    Dim pos As Long
    Dim i As Long
    Dim count As Long  ' 添加计数器
    
    count = 0  ' 初始化计数器
    
    For Each rng In ActiveDocument.StoryRanges
        Do
            For Each fld In rng.Fields
                
                ' 关键：只处理“类型为 REF 的域”
                If fld.Type = wdFieldRef Then
                    
                    codeText = fld.Code.Text
                    
                    ' 只处理 MathType 的 ZEqnNum
                    If InStr(codeText, "ZEqnNum") > 0 Then
                        
                        ' 提取 ZEqnNumXXXX
                        pos = InStr(codeText, "ZEqnNum")
                        labelName = Mid(codeText, pos)
                        
                        For i = 1 To Len(labelName)
                            If Mid(labelName, i, 1) = " " Then Exit For
                        Next i
                        labelName = Left(labelName, i - 1)
                        
                        ' 只重写 REF 域本身
                        fld.Code.Text = " REF \* Charformat \! " & labelName & " \h \* MERGEFORMAT "
                        
                        fld.Update
                        
                        count = count + 1  ' 计数加1
                        
                    End If
                    
                End If
                
            Next fld
            
            Set rng = rng.NextStoryRange
            
        Loop Until rng Is Nothing
    Next rng

    ' 修改后的提示信息，显示完成数量
    MsgBox "完成！共转换了 " & count & " 个MathType公式为超链接。"
End Sub