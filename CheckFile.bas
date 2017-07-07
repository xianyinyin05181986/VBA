Attribute VB_Name = "Module2"

Sub CheckWhetherExist()

Dim WorkRng As Range
On Error Resume Next

Dim currentFolder As String
currentFolder = ".\\Photos"

If Dir(currentFolder) = "" Then
'Create Folder
fso.createfolder (currentFolder)
Else
End If



Dim fldr As FileDialog
Dim sItem As String
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
        MsgBox sItem
    Else
        MsgBox "Cancel was pressed"
        Exit Sub
    End If

GetFolder = sItem
Set fldr = Nothing

xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
For Each r In WorkRng
    DirFile = "C:\Documents and Settings\Administrator\Desktop\" & File
    If Dir(DirFile) = "" Then
        MsgBox "File does not exist"
        Application.ActiveSheet.Hyperlinks.Add Cells(r.Row, 2), "Kulin\\" & Cells(r.Row, 1).Value & ".JPG"
    Else
    
    End If
    
Next
End Sub

