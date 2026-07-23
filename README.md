Option Explicit

'========================================================
' 主程式
' 搜尋目前執行簡報所在資料夾及所有子資料夾，
' 將所有 PowerPoint 檔案合併到目前簡報。
'========================================================
Public Sub MergePowerPointFiles()

    Dim targetPres As Presentation
    Dim rootFolder As String
    Dim fileList As Collection
    Dim filePath As Variant
    Dim insertedCount As Long
    Dim fileCount As Long
    Dim startSlideCount As Long
    Dim resultPath As String
    
    On Error GoTo ErrorHandler
    
    Set targetPres = ActivePresentation
    
    ' 必須先儲存執行用簡報，才能取得所在資料夾
    If targetPres.Path = "" Then
        MsgBox "請先將此空白簡報另存為 .pptm 檔案，並放到需要合併的最上層資料夾。", _
               vbExclamation, "尚未儲存"
        Exit Sub
    End If
    
    rootFolder = targetPres.Path
    Set fileList = New Collection
    
    ' 遞迴搜尋所有 PowerPoint 檔案
    CollectPowerPointFiles rootFolder, fileList, targetPres.FullName
    
    If fileList.Count = 0 Then
        MsgBox "在以下資料夾及子資料夾中找不到 PowerPoint 檔案：" & _
               vbCrLf & vbCrLf & rootFolder, _
               vbInformation, "沒有可合併檔案"
        Exit Sub
    End If
    
    ' 將檔案路徑排序
    Set fileList = SortCollection(fileList)
    
    startSlideCount = targetPres.Slides.Count
    
    Application.StatusBar = "開始合併 PowerPoint 檔案..."
    
    For Each filePath In fileList
        
        fileCount = fileCount + 1
        
        Application.StatusBar = _
            "正在處理第 " & fileCount & " / " & fileList.Count & " 個檔案：" & _
            GetFileName(CStr(filePath))
        
        insertedCount = insertedCount + _
                        InsertPresentationSlides(targetPres, CStr(filePath))
        
        DoEvents
        
    Next filePath
    
    ' 如果執行簡報原本只有一張完全空白投影片，可選擇刪除
    If startSlideCount = 1 Then
        If IsBlankSlide(targetPres.Slides(1)) Then
            targetPres.Slides(1).Delete
        End If
    End If
    
    ' 另存合併結果，避免直接覆寫執行工具
    resultPath = rootFolder & "\" & _
                 "Merged_Presentation_" & Format(Now, "yyyymmdd_hhnnss") & ".pptx"
    
    targetPres.SaveAs resultPath, ppSaveAsOpenXMLPresentation
    
    Application.StatusBar = False
    
    MsgBox "PowerPoint 合併完成。" & vbCrLf & vbCrLf & _
           "合併檔案數：" & fileList.Count & vbCrLf & _
           "新增投影片數：" & insertedCount & vbCrLf & _
           "總投影片數：" & targetPres.Slides.Count & vbCrLf & vbCrLf & _
           "輸出位置：" & resultPath, _
           vbInformation, "合併完成"
    
    Exit Sub

ErrorHandler:

    Application.StatusBar = False
    
    MsgBox "執行時發生錯誤：" & vbCrLf & _
           "錯誤編號：" & Err.Number & vbCrLf & _
           "錯誤內容：" & Err.Description, _
           vbCritical, "PowerPoint 合併失敗"

End Sub


'========================================================
' 遞迴搜尋資料夾及所有子資料夾
'========================================================
Private Sub CollectPowerPointFiles( _
    ByVal folderPath As String, _
    ByRef fileList As Collection, _
    ByVal runnerFilePath As String)

    Dim fso As Object
    Dim currentFolder As Object
    Dim subFolder As Object
    Dim currentFile As Object
    Dim extensionName As String
    
    On Error GoTo FolderError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set currentFolder = fso.GetFolder(folderPath)
    
    For Each currentFile In currentFolder.Files
        
        extensionName = LCase$(fso.GetExtensionName(currentFile.Name))
        
        Select Case extensionName
            Case "ppt", "pptx", "pptm"
                
                ' 排除目前正在執行的空白工具簡報
                If LCase$(currentFile.Path) <> LCase$(runnerFilePath) Then
                    
                    ' 排除先前產生的合併結果
                    If LCase$(Left$(currentFile.Name, 20)) <> _
                       LCase$("Merged_Presentation_") Then
                        
                        fileList.Add currentFile.Path
                        
                    End If
                    
                End If
                
        End Select
        
    Next currentFile
    
    For Each subFolder In currentFolder.SubFolders
        CollectPowerPointFiles subFolder.Path, fileList, runnerFilePath
    Next subFolder
    
    Exit Sub

FolderError:

    ' 無權限或無法讀取的資料夾直接跳過
    Err.Clear

End Sub


'========================================================
' 將指定簡報的全部投影片插入目標簡報
'========================================================
Private Function InsertPresentationSlides( _
    ByVal targetPres As Presentation, _
    ByVal sourceFilePath As String) As Long

    Dim beforeCount As Long
    Dim afterCount As Long
    
    On Error GoTo InsertError
    
    beforeCount = targetPres.Slides.Count
    
    targetPres.Slides.InsertFromFile _
        FileName:=sourceFilePath, _
        Index:=targetPres.Slides.Count
    
    afterCount = targetPres.Slides.Count
    
    InsertPresentationSlides = afterCount - beforeCount
    
    Exit Function

InsertError:

    InsertPresentationSlides = 0
    
    Debug.Print "無法合併檔案：" & sourceFilePath
    Debug.Print "錯誤：" & Err.Number & " - " & Err.Description
    
    Err.Clear

End Function


'========================================================
' 依完整路徑排序
'========================================================
Private Function SortCollection( _
    ByVal sourceCollection As Collection) As Collection

    Dim resultCollection As New Collection
    Dim fileArray() As String
    Dim i As Long
    Dim j As Long
    Dim tempValue As String
    
    If sourceCollection.Count = 0 Then
        Set SortCollection = resultCollection
        Exit Function
    End If
    
    ReDim fileArray(1 To sourceCollection.Count)
    
    For i = 1 To sourceCollection.Count
        fileArray(i) = CStr(sourceCollection(i))
    Next i
    
    For i = LBound(fileArray) To UBound(fileArray) - 1
        
        For j = i + 1 To UBound(fileArray)
            
            If StrComp(fileArray(i), fileArray(j), vbTextCompare) > 0 Then
                
                tempValue = fileArray(i)
                fileArray(i) = fileArray(j)
                fileArray(j) = tempValue
                
            End If
            
        Next j
        
    Next i
    
    For i = LBound(fileArray) To UBound(fileArray)
        resultCollection.Add fileArray(i)
    Next i
    
    Set SortCollection = resultCollection

End Function


'========================================================
' 判斷投影片是否為空白
'========================================================
Private Function IsBlankSlide(ByVal targetSlide As Slide) As Boolean

    Dim currentShape As Shape
    Dim hasContent As Boolean
    
    hasContent = False
    
    For Each currentShape In targetSlide.Shapes
        
        If currentShape.Type = msoPlaceholder Then
            
            If currentShape.HasTextFrame Then
                
                If currentShape.TextFrame.HasText Then
                    
                    If Trim$(currentShape.TextFrame.TextRange.Text) <> "" Then
                        hasContent = True
                        Exit For
                    End If
                    
                End If
                
            End If
            
        Else
            
            hasContent = True
            Exit For
            
        End If
        
    Next currentShape
    
    IsBlankSlide = Not hasContent

End Function


'========================================================
' 從完整路徑取得檔案名稱
'========================================================
Private Function GetFileName(ByVal fullPath As String) As String

    Dim position As Long
    
    position = InStrRev(fullPath, "\")
    
    If position > 0 Then
        GetFileName = Mid$(fullPath, position + 1)
    Else
        GetFileName = fullPath
    End If

End Function
