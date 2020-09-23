<div align="center">

## Split any file into smaller files


</div>

### Description

This code will read any large file and split it into smaller chuncks so you can copy to stiffy,e-mail or ftp it. This code is for you out there playing with file management etc. This code is very basic but it does some cool things. It will leave the source file and will create a bunch of smaller files in the same directory.. This code can be modified to output directly to the stiffy drive if you want.
 
### More Info
 
Create a new form and drop four Command Buttons on it (Command1 to Command4). Also drop a Textbox on it (Text1) and a Combobox (Combo1). You can (if you want) place a label above the textbox and change it's caption to "Source File" and a label above the combobox and change it's caption to "Split File size".

Now copy the source into the form and the module. Run and have fun ;).

If you make a nice util with the code please send me a copy : riaana@hotmail.com

If checked the split files after I've Assembled them again with FC (FileCompare) in binary mode and it didn't find any differences. But you should know that you are playing with files so don't delete the origanal after you've checked that you can re-assemble it ok.

Split files with extensions from Myfile.000 to MyFile.999

None that I know of... This code can be a basis for a cool util (You have to e-mail me that cool util .. riaana@hotmail.com)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Riaan Aspeling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/riaan-aspeling.md)
**Level**          |Unknown
**User Rating**    |6.0 (605 globes from 101 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/riaan-aspeling-split-any-file-into-smaller-files__1-1662/archive/master.zip)

### API Declarations

```
'*************************************
'*** PASTE THIS CODE INTO A MODULE ***
'*************************************
Option Explicit
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Type OPENFILENAME
 lStructSize As Long
 hwndOwner As Long
 hInstance As Long
 lpstrFilter As String
 lpstrCustomFilter As String
 nMaxCustFilter As Long
 nFilterIndex As Long
 lpstrFile As String
 nMaxFile As Long
 lpstrFileTitle As String
 nMaxFileTitle As Long
 lpstrInitialDir As String
 lpstrTitle As String
 flags As Long
 nFileOffset As Integer
 nFileExtension As Integer
 lpstrDefExt As String
 lCustData As Long
 lpfnHook As Long
 lpTemplateName As String
End Type
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type SHITEMID
 mkidcb As Long
 abID As Byte
End Type
Public Type ITEMIDLIST
 idlmkid As SHITEMID
End Type
Public Type BROWSEINFO
 hOwner As Long
 pidlRoot As Long
 pszDisplayName As String
 lpszTitle As String
 ulFlags As Long
 lpfn As Long
 lParam As Long
 iImage As Long
End Type
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Const BIF_RETURNONLYFSDIRS = &H1
Function GetOpenFileNameDLG(Filter As String, Title As String, DefaultExt As String, WindowHnd As Long) As String
On Error GoTo handelopenfile
 Dim OpenFile As OPENFILENAME, Tempstr As String
 Dim Success As Long, FileTitleLength%
 Filter = Find_And_Replace(Filter, "|", Chr(0))
 If Right$(Filter, 1) <> Chr(0) Then Filter = Filter & Chr(0)
 OpenFile.lStructSize = Len(OpenFile)
 OpenFile.hwndOwner = WindowHnd
 OpenFile.hInstance = App.hInstance
 OpenFile.lpstrFilter = Filter
 OpenFile.nFilterIndex = 1
 OpenFile.lpstrFile = String(257, 0)
 OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
 OpenFile.lpstrFileTitle = OpenFile.lpstrFile
 OpenFile.nMaxFileTitle = OpenFile.nMaxFile
 OpenFile.lpstrTitle = Title
 OpenFile.lpstrDefExt = DefaultExt
 OpenFile.flags = 0
 Success = GetOpenFileName(OpenFile)
 If Success = 0 Then
 GetOpenFileNameDLG = ""
 Else
 Tempstr = OpenFile.lpstrFile
 GetOpenFileNameDLG = Trim(Tempstr)
 End If
 Exit Function
handelopenfile:
 MsgBox Err.Description, 16, "Error " & Err.Number
 Exit Function
End Function
Function Find_And_Replace(ByRef TextLine As String, ByRef SourceStr As String, ByRef ReplaceStr As String) As String
On Error GoTo handelfindandreplace
 Dim DoAnother As Boolean, PosFound As Integer, ReturnStr As String
 DoAnother = True
 ReturnStr = TextLine
 While DoAnother
 PosFound = InStr(1, ReturnStr, SourceStr)
 If PosFound > 0 Then
  ReturnStr = Mid$(ReturnStr, 1, PosFound - 1) & ReplaceStr & Mid$(ReturnStr, PosFound + Len(SourceStr))
  Else
  DoAnother = False
 End If
 Wend
 Find_And_Replace = ReturnStr
handelfindandreplace:
 Exit Function
End Function
```


### Source Code

```
'***********************************
'*** PASTE THIS CODE INTO A FORM ***
'***********************************
Option Explicit
Private Sub Command1_Click()
 Dim Ans As String
 Ans = GetOpenFileNameDLG("File to split *.*|*.*|File to combine *.000|*.000|", "Please select a file", "", Me.hwnd)
 If Ans <> "" Then
 Text1.Text = Ans
 End If
End Sub
Private Sub Command2_Click()
 'Check that somting is selected
 If Not CheckForFile Then Exit Sub
 'Ok split the file in the current directory
 If SplitFile(Text1.Text, Combo1.ItemData(Combo1.ListIndex)) Then
 MsgBox "File was split!"
 Else
 MsgBox "Error splitting file..."
 End If
End Sub
Private Sub Command3_Click()
 'Check that somting is selected
 If Not CheckForFile Then Exit Sub
 'Check to see if it is a Split file with extension "MYFILE.SP(x)"
 If (Right$(Text1.Text, 3)) <> "000" Then
 MsgBox "That's not the proper extension for a split file. It should be somthing like Myfile.000, the first file of the split files.", 16, "No go !"
 Exit Sub
 End If
 'Ok assemble the files in the current directory
 If AssembleFile(Text1.Text) Then
 MsgBox "File assembled!"
 Else
 MsgBox "Error assembeling file..."
 End If
End Sub
Private Sub Command4_Click()
 Unload Me
 End
End Sub
Private Sub Form_Load()
 Text1.Text = ""
 Combo1.AddItem "16 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 16
 Combo1.AddItem "32 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 32
 Combo1.AddItem "64 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 64
 Combo1.AddItem "128 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 128
 Combo1.AddItem "256 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 256
 Combo1.AddItem "512 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 512
 Combo1.AddItem "720 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 720
 Combo1.AddItem "1200 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 1200
 Combo1.AddItem "1440 Kb"
 Combo1.ItemData(Combo1.NewIndex) = 1440
 Combo1.ListIndex = Combo1.ListCount - 1
 Command1.Caption = "Browse"
 Command2.Caption = "Split File"
 Command3.Caption = "Assemble Files"
 Command4.Caption = "Cancel"
End Sub
Function CheckForFile() As Boolean
 'We don't want nasty spaces in the end
 Text1.Text = Trim(Text1.Text)
 CheckForFile = False
 'Check for text in textbox
 If Text1.Text = "" Then
 'Stop !! no text entered
 MsgBox "Please select a file first!", 16, "No file selected"
 Exit Function
 End If
 'Check if the file excists
 If Dir$(Text1.Text, vbNormal) = "" Then
 'Stop !! no file
 MsgBox "The file '" & Text1.Text & "' was not found!", 16, "File non excistend?!"
 Exit Function
 End If
 CheckForFile = True
End Function
Function SplitFile(Filename As String, Filesize As Long) As Boolean
On Error GoTo handelsplit
 Dim lSizeOfFile As Long, iCountFiles As Integer
 Dim iNumberOfFiles As Integer, lSizeOfCurrentFile As Long
 Dim sBuffer As String '10Kb buffer
 Dim sRemainBuffer As String, lEndPart As Long
 Dim lSizeToSplit As Long, sHeader As String * 16
 Dim iFileCounter As Integer, sNewFilename As String
 Dim lWhereInFileCounter As Long
 If MsgBox("Continue to split file?", 4 + 32 + 256, "Split?") = vbNo Then
 SplitFile = False
 Exit Function
 End If
 Open Filename For Binary As #1
 lSizeOfFile = LOF(1)
 lSizeToSplit = Filesize * 1024
 'Check if the file is actualy larger than the selected split size
 If lSizeOfFile <= lSizeToSplit Then
 Close #1
 SplitFile = False
 MsgBox "This file is smaller than the selected split size! Why split it ?", 16, "Duh!"
 Exit Function
 End If
 'Check if file isn't alread split
 sHeader = Input(16, #1)
 Close #1
 If Mid$(sHeader, 1, 7) = "SPLITIT" Then
 MsgBox "This file is alread split!"
 SplitFile = False
 Exit Function
 End If
 Open Filename For Binary As #1
 lSizeOfFile = LOF(1)
 lSizeToSplit = Filesize * 1024
 'Write the header of the split file
 ' Signature   = "SPLITIT" = Size 7
 ' Split Number  = "xxx" = Size 3
 ' Total Number of Split Files = "xxx" = Size 3
 ' Origanal file extension = "aaa" = Size 3
 'Total of 16 for header
 iCountFiles = 0
 iNumberOfFiles = (lSizeOfFile \ lSizeToSplit) + 1
 sHeader = "SPLITIT" & Format$(iFileCounter, "000") & Format$(iNumberOfFiles, "000") & Right$(Filename, 3)
 sNewFilename = Left$(Filename, Len(Filename) - 3) & Format$(iFileCounter, "000")
 Open sNewFilename For Binary As #2
 Put #2, , sHeader 'Write the header
 lSizeOfCurrentFile = Len(sHeader)
 While Not EOF(1)
 Me.Caption = "File Split : " & iFileCounter & " (" & Int(lSizeOfCurrentFile / 1024) & " Kb)"
 Me.Refresh
 sBuffer = Input(10240, #1)
 lSizeOfCurrentFile = lSizeOfCurrentFile + Len(sBuffer)
 If lSizeOfCurrentFile > lSizeToSplit Then
  'Write last bit
  lEndPart = Len(sBuffer) - (lSizeOfCurrentFile - lSizeToSplit) + Len(sHeader)
  Put #2, , Mid$(sBuffer, 1, lEndPart)
  Close #2
  'Make new file
  iFileCounter = iFileCounter + 1
  sHeader = "SPLITIT" & Format$(iFileCounter, "000") & Format$(iNumberOfFiles, "000") & Right$(Filename, 3)
  sNewFilename = Left$(Filename, Len(Filename) - 3) & Format$(iFileCounter, "000")
  Open sNewFilename For Binary As #2
  Put #2, , sHeader 'Write the header
  'Put Rest of buffer read
  Put #2, , Mid$(sBuffer, lEndPart + 1)
  lSizeOfCurrentFile = Len(sHeader) + (Len(sBuffer) - lEndPart)
  Else
  Put #2, , sBuffer
 End If
 Wend
 Me.Caption = "Finished"
 Close #2
 Close #1
 SplitFile = True
 Exit Function
handelsplit:
 SplitFile = False
 MsgBox Err.Description, 16, "Error #" & Err.Number
 Exit Function
End Function
Function AssembleFile(Filename As String) As Boolean
On Error GoTo handelassemble
 Dim sHeader As String * 16
 Dim sBuffer As String '10Kb buffer
 Dim sFileExt As String, iNumberOfFiles As Integer
 Dim iCurrentFileNumber As Integer
 Dim iCounter As Integer, sTempFilename As String
 Dim sNewFilename As String
 If MsgBox("Continue to assemble file?", 4 + 256 + 32, "Assemble?") = vbNo Then
 AssembleFile = False
 Exit Function
 End If
 Open Filename For Binary As #1
 sHeader = Input(Len(sHeader), #1)
 'Check if it's a split file !!!
 If Mid$(sHeader, 1, 7) <> "SPLITIT" Then
 MsgBox "This is not a split file ;) nice try!"
 AssembleFile = False
 Exit Function
 Else
 'The first file is a split file ok
 'Read the header values
 iCurrentFileNumber = Val(Mid$(sHeader, 8, 3))
 iNumberOfFiles = Val(Mid$(sHeader, 11, 3))
 sFileExt = Mid$(sHeader, 14, 3)
 If iCurrentFileNumber <> 0 Then
  MsgBox "This is not the first file in the sequence!!! AAAGGHH!"
  AssembleFile = False
  Exit Function
 End If
 End If
 Close #1
 sNewFilename = Left$(Filename, Len(Filename) - 3) & sFileExt
 'Create the assembled file
 Open sNewFilename For Binary As #2
 'Assemble files
 For iCounter = 0 To iNumberOfFiles - 1
 sTempFilename = Left$(Filename, Len(Filename) - 3) & Format$(iCounter, "000")
 Me.Caption = "File Assemble : " & sTempFilename
 Me.Refresh
 Open sTempFilename For Binary As #1
 sHeader = Input(Len(sHeader), #1)
 If Mid$(sHeader, 1, 7) <> "SPLITIT" Then
  MsgBox "This is not a split file ;) nice try! " & sTempFilename
  AssembleFile = False
  Exit Function
 End If
 iCurrentFileNumber = Val(Mid$(sHeader, 8, 3))
 If iCurrentFileNumber <> iCounter Then
  MsgBox "The file '" & sTempFilename & "' is out of sequence!! AARRGHH!"
  AssembleFile = False
  Close #2
  Close #1
  Exit Function
 End If
 While Not EOF(1)
  sBuffer = Input(10240, #1)
  Put #2, , sBuffer
 Wend
 Close #1
 Next iCounter
 Close #2
 Me.Caption = "Finished"
 AssembleFile = True
 Exit Function
handelassemble:
 AssembleFile = False
 MsgBox Err.Description, 16, "Error #" & Err.Number
 Exit Function
End Function
```

