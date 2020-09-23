<div align="center">

## Folder Browser


</div>

### Description

This is a Function which displays the Browse directory dialogue and returns a path as a string. Unlike the usual shell command, this code works even on machines without the Active Desktop or IE 5 installed.
 
### More Info
 
lnghWndOwner as Long, the hWnd of the Owner form

I think you need IE 4 on the machine, but not 100% sure. Call it like this :

strSelectedFolder = BrowseForFolder(me.hwnd)

BrowseForFolder as String, the path to the selected folder


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[FluffyDave](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fluffydave.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fluffydave-folder-browser__1-29551/archive/master.zip)

### API Declarations

```
'API for BrowseForFolder
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Declare Function SHBrowseForFolder Lib "shell32" _
  (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
  (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
  (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Type BrowseInfo
  hWndOwner   As Long
  pIDLRoot    As Long
  pszDisplayName As Long
  lpszTitle   As Long
  ulFlags    As Long
  lpfnCallback  As Long
  lParam     As Long
  iImage     As Long
End Type
```


### Source Code

```
Public Function BrowseForFolder(lnghWndOwner As Long) As String
  ''Opens a Treeview control that displays the directories in a computer and Returns a String
  Dim lpIDList  As Long
  Dim sBuffer   As String
  Dim szTitle   As String
  Dim tBrowseInfo As BrowseInfo
  szTitle = "This is the title"
  With tBrowseInfo
    .hWndOwner = lnghWndOwner
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  End If
  BrowseForFolder = sBuffer
End Function
```

