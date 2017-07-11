Attribute VB_Name = "Module3"
Sub FindWhetherThePhotoExist()

Dim currentFolder As String
currentFolder = ActiveWorkbook.Path & "\Photos\"
'MsgBox currentFolder
If Dir(currentFolder, vbDirectory) = Empty Then
MsgBox "Create Folder " & currentFolder
'Create Folder
MkDir (currentFolder)
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
       Exit Sub
   End If

Dim WorkRng As Range

  
Set WorkRng = Application.Selection
'  MsgBox WorkRng.Column
'    Exit Sub
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)

Dim ValidationCount As Integer
ValidationCount = 0
For Each r In WorkRng
  
    If Trim(Cells(r.Row, WorkRng.Column).Value & "") = "" Then
        GoTo NextIteration
    End If
    
    DirFile = sItem & "\\" & Cells(r.Row, WorkRng.Column).Value & ".JPG"
    If Dir(DirFile) = "" Then
        Cells(r.Row, WorkRng.Column).Interior.Color = RGB(255, 0, 0)
        ValidationCount = ValidationCount + 1
    Else
       Cells(r.Row, WorkRng.Column).Interior.Color = RGB(0, 255, 0)
       
    End If
NextIteration:
Next

If ValidationCount > 0 Then
    MsgBox "Not All Data Validate"
End If

Dim fso As Object
Set fso = VBA.CreateObject("Scripting.FileSystemObject")

Dim destination As String
Dim versionNumber As Integer

For Each r In WorkRng

    If Trim(Cells(r.Row, WorkRng.Column).Value & "") = "" Then
        GoTo NextIteration_Two
    End If

    DirFile = sItem & "\\" & Cells(r.Row, WorkRng.Column).Value & ".JPG"
    If Dir(DirFile) = "" Then
         GoTo NextIteration_Two
    Else
'        versionNumber = WorkRng.Column - 2
'        If versionNumber = 0 Then
         destination = currentFolder & Cells(r.Row, 1).Value & ".JPG"
'        Else
'            destination = currentFolder & Cells(r.Row, 1).Value & "(" & versionNumber & ")" & ".JPG"
'        End If

        If Dir(destination) = "" Then
            Call fso.CopyFile(DirFile, destination)
            GoTo NextIteration_Two
        Else
            MsgBox destination & " already exist"
            Exit Sub
        End If
    End If
NextIteration_Two:
Next

End Sub

