Sub 批量简繁转换()  '此代码为指定文件夹中所有选取的WORD文件的进行格式设置
    'On Error Resume Next'忽略错误
    '定义一个文件夹选取对话框
    Set MyDialog = Application.FileDialog(msoFileDialogFilePicker)
    With MyDialog
        .Title = "请选择要处理的文档（可多选）"
        .Filters.Clear    '清除所有文件筛选器中的项目
        .Filters.Add "所有 WORD 文件", "*.MSG", 1    '增加筛选器的项目为所有WORD文件
        .AllowMultiSelect = True    '允许多项选择
        If .Show = -1 Then    '确定
            Application.ScreenUpdating = False
            For Each vrtSelectedItem In .SelectedItems    '在所有选取项目中循环
                Set Doc = Documents.Open(FileName:=vrtSelectedItem, Visible:=True)
                Selection.Range.TCSCConverter WdTCSCConverterDirection:= _
                    wdTCSCConverterDirectionTCSC, CommonTerms:=True, UseVariants:=False
                Doc.Save
                Doc.Close
            Next
            Application.ScreenUpdating = True
        End If
    End With
    MsgBox "批量设置完毕!", vbInformation
End Sub
