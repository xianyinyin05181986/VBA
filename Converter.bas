Attribute VB_Name = "Module1"

Sub ConvertToHyperlinks()

Dim WorkRng As Range
On Error Resume Next

Dim fldr As FileDialog
Dim sItem As String
Dim DirFile As String
Dim tmpArr() As String
    

Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
With fldr
.Title = "Select a Folder"
.AllowMultiSelect = False
.InitialFileName = "."
If .Show <> -1 Then GoTo NextCode
sItem = .SelectedItems(1)
End With
NextCode:
 If Len(sItem) Then
'    MsgBox sItem
    Else
        MsgBox "Cancel was pressed"
        Exit Sub
    End If
GetFolder = sItem
Set fldr = Nothing

tmpArr = Split(sItem, "\")
'MsgBox tmpArr(UBound(tmpArr) - LBound(tmpArr))
'Exit Sub
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each r In WorkRng
       DirFile = tmpArr(UBound(tmpArr) - LBound(tmpArr)) & "\\" & Cells(r.Row, 1).Value & ".JPG"
       If Dir(DirFile) = "" Then
          MsgBox DirFile & " does not exist"
          Exit Sub
       Else
'           MsgBox DirFile
          Application.ActiveSheet.Hyperlinks.Add Cells(r.Row, 1), DirFile
       End If
    
Next
End Sub
