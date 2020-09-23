VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.UserControl FileBox 
   ClientHeight    =   3210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   KeyPreview      =   -1  'True
   ScaleHeight     =   3210
   ScaleWidth      =   4950
   ToolboxBitmap   =   "FileBox.ctx":0000
   Begin ComctlLib.TreeView TV 
      Height          =   2895
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5106
      _Version        =   327682
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "IlIcons"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.FileListBox Files 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DirListBox Folder 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.DriveListBox Drives 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.ComboBox CboFiletype 
      Height          =   315
      ItemData        =   "FileBox.ctx":0532
      Left            =   0
      List            =   "FileBox.ctx":0534
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2880
      Width           =   4935
   End
   Begin ComctlLib.ImageList IlIcons 
      Left            =   0
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":0536
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":0648
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":075A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":086C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":097E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":0A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":0BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":0CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":0DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FileBox.ctx":0ED8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FileBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'FileBox
'a simple Fileselector by Scythe
'scythe@cablenet.de

Option Explicit

'Need this to get the DriveType
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'Events the Control has
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)


Dim StrFilter As String
Dim Patterns() As String
Dim HideFiles As Boolean


'Changed the Filter
Private Sub CboFiletype_Click()
 Files.Pattern = Patterns(CboFiletype.ListIndex)
 DriveList
End Sub

'Changed the Folder (used from tv_expand)
Private Sub Folder_Change()
 Files.Path = Folder.Path
End Sub

'List all Drives
Private Sub DriveList()

 Dim i As Integer
 Dim DrivePath As String
 Dim TVIcon As Integer

 'Clear all
 TV.Nodes.Clear

 For i = 0 To Drives.ListCount - 1
  DrivePath = UCase(Left(Drives.List(i), 1)) & ":\"
  Select Case GetDriveType(DrivePath)
  Case 2 'Removable
   If i < 2 Then 'Drive a: or b:
    TVIcon = 1
   Else
    TVIcon = 4
   End If
  Case 3 'Fixed
   TVIcon = 2
  Case 4 'Remote (Network)
   TVIcon = 5
  Case 5 'CD Rom
   TVIcon = 3
  Case 6 'Ram
   TVIcon = 6
  End Select
  TV.Nodes.Add , , DrivePath, Left$(DrivePath, 2), TVIcon
  TV.Nodes.Add DrivePath, tvwChild, ""
 Next
End Sub

'Expand the Directory Tree and search for new
Private Sub tv_Expand(ByVal Node As ComctlLib.Node)
 On Error GoTo ErrExp

 Dim i As Integer
 Dim Relative As String
 Dim FolderName As String
 Dim FolderPos As Integer
 Dim Icon As Integer
 Dim NewPath As String
 Dim Ext As String
 Dim ExtPos As Integer

 MousePointer = vbHourglass

 If Node.Child.Text = "" Then

  TV.Nodes.Remove Node.Child.Index
  Relative = Node.Key
  Folder.Path = Relative
  FolderPos = Len(Relative) + 1

  'Add folders
  For i = 0 To Folder.ListCount - 1
   FolderName = Mid(Folder.List(i), FolderPos)
   NewPath = Relative & FolderName & "\"
   TV.Nodes.Add Relative, tvwChild, NewPath, FolderName, 7
   Folder.Path = NewPath
   If (Files.ListCount > 0) Or (Folder.ListCount > 0) Then
    TV.Nodes.Add NewPath, tvwChild, , ""
    TV.Nodes(NewPath).ExpandedImage = 8
   End If
   Folder.Path = Relative
  Next

  'Add files
  If HideFiles = False Then
   For i = 0 To Files.ListCount - 1
    If Right$(UCase(Files.List(i)), 3) = "TXT" Then
     Icon = 9
    Else
     Icon = 10
    End If
    TV.Nodes.Add Relative, tvwChild, , Files.List(i), Icon
   Next
  End If
 End If
ErrExp:
 MousePointer = vbDefault
End Sub

'Set the Filters
Private Sub SetFilters()
 Dim X As Integer
 Dim Y As Integer
 Dim ctr As Integer
 Dim Filterlist() As String
 Dim i As Integer
 ReDim Filterlist(0)

 X = 1

'Seperate the Filterstring
Do
 Y = InStr(X, Filter, "|")
 If Y = 0 Then Y = Len(Filter) + 1
 Filterlist(ctr) = Mid$(Filter, X, Y - X)
 ctr = ctr + 1
 ReDim Preserve Filterlist(ctr)
 X = Y + 1
 If X > Len(Filter) Then Exit Do
Loop

'Clear the Combobox and Prepare the Pattern String
CboFiletype.Clear
ReDim Patterns(ctr - 1)

X = 0
For i = 0 To ctr - 1 Step 2
If Filterlist(i + 1) <> "" Then
 'Write the Name
 CboFiletype.AddItem Filterlist(i)
 'Get the Filter to the Name
 Patterns(X) = Filterlist(i + 1)
End If
X = X + 1
Next i

'Update the View
UserControl_Resize
'If we hav a Filter then show First
If CboFiletype.ListCount > 0 Then
 CboFiletype.ListIndex = 0
Else
 DriveList
End If

End Sub

Private Sub UserControl_Initialize()
 ReDim Patterns(0)
End Sub


Private Sub UserControl_Resize()
 TV.Width = UserControl.Width
 'If we have more Than 1 Filter
 'Show Filters
 If CboFiletype.ListCount > 1 And HideFiles = False Then
  TV.Height = UserControl.Height - CboFiletype.Height
  CboFiletype.Width = TV.Width
  CboFiletype.Top = TV.Height
 Else
  'Show only treeview
  TV.Height = UserControl.Height
 End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
 Call .WriteProperty("Filter", StrFilter)
 Call .WriteProperty("Enabled", UserControl.Enabled)
 Call .WriteProperty("NoFiles", HideFiles)
 End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 With PropBag
 StrFilter = .ReadProperty("Filter", "")
 UserControl.Enabled = .ReadProperty("Enabled", True)
 HideFiles = .ReadProperty("NoFiles", False)
 End With
 SetFilters
 TV.Enabled = UserControl.Enabled
 CboFiletype.Enabled = UserControl.Enabled
End Sub


Public Property Get Filter() As String
Attribute Filter.VB_Description = "A Filter looks like this: Pictures|*.bmp;*.jpg;*.gif|All Files|*.*\r\nFirst the visible text and after a | the Filter\r\nThe Filtebox is only visible if u use more than 1 Filter"
 Filter = StrFilter
End Property
Public Property Let Filter(ByVal NewFilter As String)
 StrFilter = NewFilter
 SetFilters
 PropertyChanged "Filter"
End Property

Public Property Get NoFiles() As Boolean
Attribute NoFiles.VB_Description = "True disables the FileView so u can use it as Directory Browser"
 NoFiles = HideFiles
End Property
Public Property Let NoFiles(ByVal NewFiles As Boolean)
 HideFiles = NewFiles
 PropertyChanged "NoFiles"
 UserControl_Resize
 DriveList
End Property

Public Property Get Enabled() As Boolean
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewEna As Boolean)
 UserControl.Enabled = NewEna
 TV.Enabled = UserControl.Enabled
 CboFiletype.Enabled = UserControl.Enabled
 PropertyChanged "Enabled"
End Property

Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub tv_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub tv_KeyUp(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub CboFiletype_KeyDown(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub CboFiletype_KeyPress(KeyAscii As Integer)
 RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub CboFiletype_KeyUp(KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub TV_Click()
 RaiseEvent Click
End Sub

Private Sub TV_DblClick()
 RaiseEvent DblClick
End Sub

Public Function Filename() As String
 On Error Resume Next
 If TV.SelectedItem.Key = "" Then
  Filename = TV.SelectedItem
 End If
End Function
Public Function Filepath() As String
 On Error Resume Next
 If HideFiles Then
  Filepath = TV.SelectedItem.FullPath & "\"
 Else
  If TV.SelectedItem.Key = "" Then
   Filepath = TV.SelectedItem.Parent.FullPath & "\"
  End If
 End If
End Function

Public Function FileFullName() As String
 On Error Resume Next
 If TV.SelectedItem.Key = "" Then
  FileFullName = TV.SelectedItem.FullPath
 End If
End Function



