Attribute VB_Name = "mDirectoryFunctions"
'@Folder "Utilities"
Option Compare Database
Option Explicit

Public Sub CreateDirTree(ByVal strPath As String)
    Dim elm As Variant
    Dim strCheckPath As String
    strPath = StripFilename(strPath)
    strCheckPath = vbNullString
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Sub


Public Function StripFilename(ByVal sPathFile As String) As String

  'given a full path and file, strip the filename off the end and return the path
  
  Dim filesystem As New FileSystemObject
  
  StripFilename = filesystem.GetParentFolderName(sPathFile) & "\"
  
  Exit Function

End Function
