Attribute VB_Name = "FileStuff"
Private Type BrowseInfo
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib "Shell32.dll" (bBrowse As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long

Public Declare Function ExtractAssociatedIcon Lib "Shell32.dll" Alias "ExtractAssociateIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long


Public Function GetFile(ByVal s As String) As String
   Dim i As Integer
   Dim j As Integer
   
   i = 0
   j = 0
   
   i = InStr(s, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, s, "\")
   Loop
   
   If j = 0 Then
      GetFile = ""
   Else
      GetFile = Right$(s, Len(s) - j)
   End If
End Function

'
'  Returns the path portion of a file + pathname
'
Public Function GetPath(ByVal s As String) As String
   Dim i As Integer
   Dim j As Integer
   
   i = 0
   j = 0
   
   i = InStr(s, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, s, "\")
   Loop
   
   If j = 0 Then
      GetPath = ""
   Else
      GetPath = Left$(s, j)
   End If
End Function

Public Function GetFileSize(flen As Long) As String
If flen < 1024 Then
GetFileSize = flen & " Bytes"
Exit Function
End If

If flen < 1048576 Then
GetFileSize = Format(flen / 1024, "#####.0#") & " Kb"
Exit Function
Else
GetFileSize = Format((flen / 1024) / 1024, "#######.0#") & " Mb"
End If

End Function


' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
Public Function BrowseForDirectory(ByVal Mesage As String) As String
Dim browse_info As BrowseInfo
Dim item As Long
Dim dir_name As String
   
   browse_info.hWndOwner = hwnd
   browse_info.pidlRoot = 0
   browse_info.sDisplayName = Space$(260)
   browse_info.sTitle = Mesage
   browse_info.ulFlags = 1 ' Return directory name.
   browse_info.lpfn = 0
   browse_info.lParam = 0
   browse_info.iImage = 0
   
   item = SHBrowseForFolder(browse_info)
   If item Then
       dir_name = Space$(260)
       If SHGetPathFromIDList(item, dir_name) Then
           BrowseForDirectory = Left(dir_name, InStr(dir_name, Chr$(0)) - 1)
       Else
           BrowseForDirectory = ""
       End If
   End If
End Function
