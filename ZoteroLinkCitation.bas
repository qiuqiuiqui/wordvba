Public Sub ZoteroLinkCitation()
    On Error GoTo ErrorHandler
    
    Dim originalScreenUpdating As Boolean
    Dim originalSelectionStart As Long
    Dim originalSelectionEnd As Long
    
    ' 保存原始光标位置（用于最后恢复）
    originalSelectionStart = Selection.Start
    originalSelectionEnd = Selection.End
    originalScreenUpdating = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    
    Dim title As String, titleAnchor As String
    Dim fieldCode As String
    Dim n1&, n2&
    Dim numOrYear As String
    Dim part As Variant, refParts As Variant, dashParts As Variant
    
    ' 先临时选中整个文档来查找参考文献（但不改变用户的选区意识）
    Dim docRange As Range
    Set docRange = ActiveDocument.Range
    
    ' 在文档末尾查找 Zotero 参考文献
    docRange.Collapse Direction:=wdCollapseEnd
    docRange.Find.ClearFormatting
    With docRange.Find
        .Text = "^d ADDIN ZOTERO_BIBL"
        .Forward = False  ' 从后往前找
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    
    If docRange.Find.Execute Then
        ' 找到后扩展到整个域范围
        docRange.Select
        ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:="Zotero_Bibliography"
    Else
        ' 如果没找到，尝试另一种方式：遍历所有域
        Dim bibField As Field
        Dim foundBib As Boolean
        foundBib = False
        For Each bibField In ActiveDocument.Fields
            If InStr(bibField.Code, "ZOTERO_BIBL") > 0 Then
                ActiveDocument.Bookmarks.Add Range:=bibField.result, Name:="Zotero_Bibliography"
                foundBib = True
                Exit For
            End If
        Next bibField
        
        If Not foundBib Then
            MsgBox "未找到 Zotero 参考文献，请先插入文献并刷新（按 F9）。", vbExclamation
            GoTo ErrorHandler
        End If
    End If
    
    ' 遍历文档中所有 Zotero 引用域
    Dim aField As Field
    For Each aField In ActiveDocument.Fields
        If InStr(aField.Code, "ADDIN ZOTERO_ITEM") > 0 Then
            fieldCode = aField.Code
            Do While InStr(fieldCode, """title"":""") > 0
                ' 解析标题
                n1 = InStr(fieldCode, """title"":""") + Len("""title"":""")
                n2 = InStr(Mid(fieldCode, n1), """,""") - 1 + n1
                If n2 < n1 Then Exit Do
                title = Mid(fieldCode, n1, n2 - n1)
                
                ' 生成合法书签名
                titleAnchor = CleanBookmarkName(title)
                
                ' 在文末匹配对应参考文献并加书签
                On Error Resume Next
                Selection.GoTo What:=wdGoToBookmark, Name:="Zotero_Bibliography"
                If Err.Number <> 0 Then GoTo NextField
                On Error GoTo 0
                
                ' 在参考文献区域查找标题
                Selection.Find.ClearFormatting
                With Selection.Find
                    .Text = Left(title, 255)
                    .Forward = True
                    .Wrap = wdFindStop
                    .MatchWildcards = False
                End With
                Selection.Find.Execute
                
                If Selection.Find.Found Then
                    Selection.Paragraphs(1).Range.Select
                    On Error Resume Next
                    ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:=titleAnchor
                    If Err.Number <> 0 Then
                        titleAnchor = titleAnchor & "_" & Format(Now, "hhmmss")
                        Err.Clear
                        ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:=titleAnchor
                    End If
                    On Error GoTo 0
                End If
                
                ' 获取文中引用编号字符串
                aField.Select
                numOrYear = Selection.Range.Text
                numOrYear = Replace(numOrYear, "[", "")
                numOrYear = Replace(numOrYear, "]", "")
                numOrYear = Trim(numOrYear)
                
                ' 按逗号拆分
                refParts = Split(numOrYear, ",")
                
                ' 遍历每个部分
                For Each part In refParts
                    part = Trim(part)
                    ' 判断是否是区间
                    If InStr(part, "-") > 0 Or InStr(part, "–") > 0 Or InStr(part, "—") > 0 Then
                        dashParts = Split(Replace(Replace(Replace(part, "–", "-"), "—", "-"), ChrW(&H2010), "-"), "-")
                        If UBound(dashParts) = 1 Then
                            dashParts(0) = Trim(dashParts(0))
                            dashParts(1) = Trim(dashParts(1))
                            InsertRefLink dashParts(0), titleAnchor, aField
                            InsertRefLink dashParts(1), titleAnchor, aField
                        End If
                    Else
                        InsertRefLink part, titleAnchor, aField
                    End If
                Next part
                
NextField:
                fieldCode = Mid(fieldCode, n2 + 1)
            Loop
        End If
    Next aField
    
ErrorHandler:
    ' 恢复屏幕更新
    Application.ScreenUpdating = originalScreenUpdating
    
    ' 恢复原始光标位置
    On Error Resume Next
    If originalSelectionStart >= 0 And originalSelectionStart <= ActiveDocument.Range.End Then
        Dim restoreRange As Range
        Set restoreRange = ActiveDocument.Range(originalSelectionStart, originalSelectionEnd)
        restoreRange.Select
    End If
    On Error GoTo 0
    
    If Err.Number <> 0 And Err.Number <> 0 Then
        MsgBox "执行过程中出现错误：" & Err.Description, vbInformation
    End If
End Sub

' 清理书签名称的辅助函数
Private Function CleanBookmarkName(ByVal title As String) As String
    Dim result As String
    Dim i As Integer, ch As String
    
    ' 基本替换
    result = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(title, " ", "_"), "#", "_"), "&", "_"), ":", "_"), ",", "_"), "-", "_"), "‐", "_"), "'", "_"), ".", "_"), "(", "_"), ")", "_"), "?", "_"), "!", "_")
    
    ' 只保留字母、数字、下划线
    Dim clean As String
    clean = ""
    For i = 1 To Len(result)
        ch = Mid(result, i, 1)
        If (ch Like "[A-Za-z0-9_]") Then
            clean = clean & ch
        Else
            clean = clean & "_"
        End If
    Next i
    result = clean
    
    ' 确保以字母开头
    If Len(result) > 0 Then
        If Not (Left(result, 1) Like "[A-Za-z]") Then
            result = "B_" & result
        End If
    Else
        result = "Bookmark_" & Format(Now, "yyyymmddhhnnss")
    End If
    
    ' 限制长度
    CleanBookmarkName = Left(result, 40)
End Function

Private Sub InsertRefLink(ByVal refNum As String, ByVal anchorName As String, ByVal aField As Field)
    On Error Resume Next
    Dim findRange As Range
    Set findRange = aField.result.Duplicate
    With findRange.Find
        .ClearFormatting
        .Text = refNum
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        If .Execute Then
            ActiveDocument.Hyperlinks.Add Anchor:=findRange, _
                Address:="", SubAddress:=anchorName, TextToDisplay:=refNum
        End If
    End With
    On Error GoTo 0
End Sub