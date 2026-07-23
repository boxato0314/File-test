Option Explicit

'============================================================
' PowerPoint 資料夾及子資料夾簡報自動合併工具
'
' 使用方式：
' 1. 將此 PowerPoint 儲存為 .pptm
' 2. 放在需要搜尋的最上層資料夾
' 3. 執行 MergeAllPowerPointFiles
'
' 功能：
' - 搜尋目前資料夾及所有子資料夾
' - 合併 PPT、PPTX、PPTM、PPS、PPSX、PPSM
' - 排除目前執行中的 PPTM
' - 排除先前輸出的 Merged_Presentation_ 檔案
' - 依完整路徑及檔名排序
' - 建立新的 PPTX 輸出檔
'============================================================


'============================================================
' 主程式
'============================================================
Public Sub MergeAllPowerPointFiles()

    Dim runnerPres As Presentation
    Dim outputPres As Presentation
    
    Dim rootFolder As String
    Dim outputPath As String
    
    Dim fileList As Collection
    Dim sortedFileList As Collection
    
    Dim filePath As Variant
    
    Dim totalFileCount As Long
    Dim processedFileCount As Long
    Dim successFileCount As Long
    Dim failedFileCount As Long
    Dim insertedSlideCount As Long
    Dim addedSlides As Long
    
    Dim failedFiles As String
    Dim resultMessage As String
    
    On Error GoTo MainError
    
    '--------------------------------------------------------
    ' 取得目前執行此巨集的簡報
    '--------------------------------------------------------
    Set runnerPres = ActivePresentation
    
    '--------------------------------------------------------
    ' 檢查執行簡報是否已儲存
    '--------------------------------------------------------
    If runnerPres.Path = "" Then
        
        MsgBox _
            "請先將此空白 PowerPoint 儲存為 .pptm 檔案，" & vbCrLf & _
            "並放在需要搜尋及合併的最上層資料夾。", _
            vbExclamation, _
            "請先儲存執行工具"
        
        Exit Sub
        
    End If
    
    '--------------------------------------------------------
    ' 建議執行工具使用 PPTM 格式
    '--------------------------------------------------------
    If LCase$(GetFileExtension(runnerPres.Name)) <> "pptm" Then
        
        If MsgBox( _
            "目前執行簡報不是 .pptm 格式。" & vbCrLf & vbCrLf & _
            "若關閉簡報，VBA 程式可能無法保留。" & vbCrLf & _
            "是否仍要繼續執行？", _
            vbQuestion + vbYesNo, _
            "格式確認") = vbNo Then
            
            Exit Sub
            
        End If
        
    End If
    
    rootFolder = runnerPres.Path
    
    '--------------------------------------------------------
    ' 建立檔案清單
    '--------------------------------------------------------
    Set fileList = New Collection
    
    CollectPowerPointFiles _
        folderPath:=rootFolder, _
        fileList:=fileList, _
        runnerFullName:=runnerPres.FullName
    
    '--------------------------------------------------------
    ' 檢查是否找到可合併檔案
    '--------------------------------------------------------
    If fileList.Count = 0 Then
        
        MsgBox _
            "在以下資料夾及其子資料夾中，" & vbCrLf & _
            "找不到可合併的 PowerPoint 檔案：" & vbCrLf & vbCrLf & _
            rootFolder, _
            vbInformation, _
            "沒有找到簡報"
        
        Exit Sub
        
    End If
    
    '--------------------------------------------------------
    ' 排序檔案清單
    '--------------------------------------------------------
    Set sortedFileList = SortFileCollection(fileList)
    
    totalFileCount = sortedFileList.Count
    
    '--------------------------------------------------------
    ' 執行前確認
    '--------------------------------------------------------
    If MsgBox( _
        "已找到 " & totalFileCount & " 個 PowerPoint 檔案。" & vbCrLf & vbCrLf & _
        "搜尋資料夾：" & vbCrLf & _
        rootFolder & vbCrLf & vbCrLf & _
        "檔案將依照完整路徑及檔名排序後合併。" & vbCrLf & _
        "是否開始執行？", _
        vbQuestion + vbYesNo, _
        "確認開始合併") = vbNo Then
        
        Exit Sub
        
    End If
    
    '--------------------------------------------------------
    ' 建立新的輸出簡報
    ' 不直接使用執行工具，避免破壞 PPTM
    '--------------------------------------------------------
    Set outputPres = Presentations.Add
    
    '--------------------------------------------------------
    ' 逐一插入簡報
    '--------------------------------------------------------
    For Each filePath In sortedFileList
        
        processedFileCount = processedFileCount + 1
        
        Debug.Print String(80, "-")
        Debug.Print "處理進度：" & processedFileCount & " / " & totalFileCount
        Debug.Print "檔案：" & CStr(filePath)
        
        addedSlides = InsertPresentationFile( _
            outputPres:=outputPres, _
            sourceFilePath:=CStr(filePath))
        
        If addedSlides >= 0 Then
            
            successFileCount = successFileCount + 1
            insertedSlideCount = insertedSlideCount + addedSlides
            
            Debug.Print "結果：成功"
            Debug.Print "新增投影片：" & addedSlides
            
        Else
            
            failedFileCount = failedFileCount + 1
            
            failedFiles = failedFiles & _
                          failedFileCount & ". " & _
                          CStr(filePath) & vbCrLf
            
            Debug.Print "結果：失敗"
            
        End If
        
        DoEvents
        
    Next filePath
    
    '--------------------------------------------------------
    ' 如果完全沒有成功插入投影片
    '--------------------------------------------------------
    If outputPres.Slides.Count = 0 Then
        
        outputPres.Close
        
        MsgBox _
            "沒有任何投影片成功合併。" & vbCrLf & vbCrLf & _
            "請確認來源簡報是否損壞、被加密、被其他程式鎖定，" & _
            "或目前帳號是否具有讀取權限。", _
            vbCritical, _
            "合併失敗"
        
        Exit Sub
        
    End If
    
    '--------------------------------------------------------
    ' 建立輸出檔案名稱
    '--------------------------------------------------------
    outputPath = rootFolder & "\" & _
                 "Merged_Presentation_" & _
                 Format(Now, "yyyymmdd_hhnnss") & ".pptx"
    
    '--------------------------------------------------------
    ' 儲存輸出簡報
    '--------------------------------------------------------
    outputPres.SaveAs _
        FileName:=outputPath, _
        FileFormat:=ppSaveAsOpenXMLPresentation
    
    '--------------------------------------------------------
    ' 顯示完成訊息
    '--------------------------------------------------------
    resultMessage = _
        "PowerPoint 合併完成。" & vbCrLf & vbCrLf & _
        "搜尋到的檔案數：" & totalFileCount & vbCrLf & _
        "成功合併檔案數：" & successFileCount & vbCrLf & _
        "失敗檔案數：" & failedFileCount & vbCrLf & _
        "合併投影片總數：" & insertedSlideCount & vbCrLf & vbCrLf & _
        "輸出檔案：" & vbCrLf & _
        outputPath
    
    If failedFileCount > 0 Then
        
        resultMessage = resultMessage & vbCrLf & vbCrLf & _
                        "以下檔案未能合併：" & vbCrLf & _
                        failedFiles
        
    End If
    
    MsgBox _
        resultMessage, _
        vbInformation, _
        "合併完成"
    
    Exit Sub


MainError:

    MsgBox _
        "主程式執行時發生錯誤。" & vbCrLf & vbCrLf & _
        "錯誤編號：" & Err.Number & vbCrLf & _
        "錯誤內容：" & Err.Description, _
        vbCritical, _
        "程式錯誤"

End Sub


'============================================================
' 遞迴搜尋指定資料夾及所有子資料夾
'============================================================
Private Sub CollectPowerPointFiles( _
    ByVal folderPath As String, _
    ByRef fileList As Collection, _
    ByVal runnerFullName As String)

    Dim fileSystem As Object
    Dim currentFolder As Object
    Dim subFolder As Object
    Dim currentFile As Object
    
    Dim extensionName As String
    Dim fileNameLower As String
    Dim filePathLower As String
    Dim runnerPathLower As String
    
    On Error GoTo FolderReadError
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set currentFolder = fileSystem.GetFolder(folderPath)
    
    runnerPathLower = LCase$(runnerFullName)
    
    '--------------------------------------------------------
    ' 搜尋目前資料夾中的檔案
    '--------------------------------------------------------
    For Each currentFile In currentFolder.Files
        
        extensionName = LCase$( _
            fileSystem.GetExtensionName(currentFile.Name))
        
        fileNameLower = LCase$(currentFile.Name)
        filePathLower = LCase$(currentFile.Path)
        
        If IsPowerPointExtension(extensionName) Then
            
            ' 排除 Office 暫存檔，例如 ~$123.pptx
            If Left$(fileNameLower, 2) <> "~$" Then
                
                ' 排除目前執行巨集的簡報
                If filePathLower <> runnerPathLower Then
                    
                    ' 排除之前產生的合併結果
                    If Left$(fileNameLower, 20) <> _
                       LCase$("Merged_Presentation_") Then
                        
                        fileList.Add currentFile.Path
                        
                    End If
                    
                End If
                
            End If
            
        End If
        
    Next currentFile
    
    '--------------------------------------------------------
    ' 遞迴搜尋所有子資料夾
    '--------------------------------------------------------
    For Each subFolder In currentFolder.SubFolders
        
        CollectPowerPointFiles _
            folderPath:=subFolder.Path, _
            fileList:=fileList, _
            runnerFullName:=runnerFullName
        
    Next subFolder
    
    Exit Sub


FolderReadError:

    ' 遇到無權限或無法讀取的資料夾時跳過
    Debug.Print String(80, "-")
    Debug.Print "無法讀取資料夾：" & folderPath
    Debug.Print "錯誤編號：" & Err.Number
    Debug.Print "錯誤內容：" & Err.Description
    
    Err.Clear

End Sub


'============================================================
' 判斷副檔名是否為 PowerPoint 格式
'============================================================
Private Function IsPowerPointExtension( _
    ByVal extensionName As String) As Boolean

    Select Case LCase$(extensionName)
        
        Case "ppt", "pptx", "pptm", _
             "pps", "ppsx", "ppsm"
            
            IsPowerPointExtension = True
        
        Case Else
            
            IsPowerPointExtension = False
        
    End Select

End Function


'============================================================
' 插入單一 PowerPoint 檔案
'
' 回傳值：
' 0 以上：成功，數值為新增投影片數
' -1：失敗
'============================================================
Private Function InsertPresentationFile( _
    ByVal outputPres As Presentation, _
    ByVal sourceFilePath As String) As Long

    Dim beforeSlideCount As Long
    Dim afterSlideCount As Long
    Dim addedSlideCount As Long
    
    On Error GoTo InsertError
    
    beforeSlideCount = outputPres.Slides.Count
    
    ' Index 代表插入在哪一張投影片之後
    ' 0 表示插入到最前面
    outputPres.Slides.InsertFromFile _
        FileName:=sourceFilePath, _
        Index:=outputPres.Slides.Count
    
    afterSlideCount = outputPres.Slides.Count
    addedSlideCount = afterSlideCount - beforeSlideCount
    
    InsertPresentationFile = addedSlideCount
    
    Exit Function


InsertError:

    Debug.Print "無法插入簡報：" & sourceFilePath
    Debug.Print "錯誤編號：" & Err.Number
    Debug.Print "錯誤內容：" & Err.Description
    
    InsertPresentationFile = -1
    
    Err.Clear

End Function


'============================================================
' 將 Collection 中的完整檔案路徑排序
' 排序方式：不分大小寫，依完整路徑由小到大
'============================================================
Private Function SortFileCollection( _
    ByVal sourceCollection As Collection) As Collection

    Dim resultCollection As Collection
    Dim fileArray() As String
    
    Dim i As Long
    Dim j As Long
    Dim tempValue As String
    
    Set resultCollection = New Collection
    
    If sourceCollection Is Nothing Then
        
        Set SortFileCollection = resultCollection
        Exit Function
        
    End If
    
    If sourceCollection.Count = 0 Then
        
        Set SortFileCollection = resultCollection
        Exit Function
        
    End If
    
    ReDim fileArray(1 To sourceCollection.Count)
    
    '--------------------------------------------------------
    ' Collection 轉成陣列
    '--------------------------------------------------------
    For i = 1 To sourceCollection.Count
        
        fileArray(i) = CStr(sourceCollection(i))
        
    Next i
    
    '--------------------------------------------------------
    ' 排序
    '--------------------------------------------------------
    For i = LBound(fileArray) To UBound(fileArray) - 1
        
        For j = i + 1 To UBound(fileArray)
            
            If StrComp( _
                fileArray(i), _
                fileArray(j), _
                vbTextCompare) > 0 Then
                
                tempValue = fileArray(i)
                fileArray(i) = fileArray(j)
                fileArray(j) = tempValue
                
            End If
            
        Next j
        
    Next i
    
    '--------------------------------------------------------
    ' 陣列轉回 Collection
    '--------------------------------------------------------
    For i = LBound(fileArray) To UBound(fileArray)
        
        resultCollection.Add fileArray(i)
        
    Next i
    
    Set SortFileCollection = resultCollection

End Function


'============================================================
' 取得副檔名
' 例如：
' ABC.pptm → pptm
'============================================================
Private Function GetFileExtension( _
    ByVal fileName As String) As String

    Dim dotPosition As Long
    
    dotPosition = InStrRev(fileName, ".")
    
    If dotPosition > 0 Then
        
        GetFileExtension = Mid$(fileName, dotPosition + 1)
        
    Else
        
        GetFileExtension = ""
        
    End If

End Function
