VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "IExplorer Preview"
   ClientHeight    =   7620
   ClientLeft      =   690
   ClientTop       =   1815
   ClientWidth     =   12450
   Icon            =   "FST.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   12450
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   7305
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   26458
            MinWidth        =   26458
         EndProperty
      EndProperty
   End
   Begin FileSelect.Splitter Splitter1 
      Height          =   6015
      Left            =   180
      TabIndex        =   0
      Top             =   1080
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   10610
      SplitterBorder  =   0
      SplitterSize    =   5
      RatioFromTop    =   0.3
      Child1          =   "FileBox1"
      Child2          =   "web1"
      Begin FileSelect.FileBox FileBox1 
         Height          =   5955
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   10504
         Filter          =   ""
         Enabled         =   -1  'True
         NoFiles         =   0   'False
      End
      Begin SHDocVwCtl.WebBrowser web1 
         Height          =   5955
         Left            =   2205
         TabIndex        =   1
         Top             =   0
         Width           =   5010
         ExtentX         =   8837
         ExtentY         =   10504
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12450
      _ExtentX        =   21960
      _ExtentY        =   1429
      ButtonWidth     =   1720
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "H&ome"
            Key             =   "home"
            Object.ToolTipText     =   "Home"
            ImageKey        =   "home"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "B&ack"
            Key             =   "prev"
            Object.ToolTipText     =   "Back in history"
            ImageKey        =   "prev"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "For&ward"
            Key             =   "next"
            Object.ToolTipText     =   "Forward in history"
            ImageKey        =   "next"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reload"
            Key             =   "refresh"
            Object.ToolTipText     =   "Reload current page"
            ImageKey        =   "refresh"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Stop"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop current download"
            ImageKey        =   "stop"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sa&ve As..."
            Key             =   "save"
            Object.ToolTipText     =   "Save currently viewed file"
            ImageKey        =   "save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FST.frx":27A2
            Key             =   "home"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FST.frx":347C
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FST.frx":4156
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FST.frx":4E30
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FST.frx":5B0A
            Key             =   "next"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FST.frx":67E4
            Key             =   "save"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HTMLpage As HTMLDocument

Function ExtractFileName(Filename As String) As String
    
'Extract the File title from a full file name


    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(Filename, "\")
    If pos = 0 Then
            ExtractFileName = Filename
            Exit Function
    End If
    
    Do While pos <> 0
            PrevPos = pos
            pos = InStr(pos + 1, Filename, "\")
    Loop

    ExtractFileName = Right(Filename, Len(Filename) - PrevPos)

End Function

Function ExtractFileExtension(ByVal Filename As String) As String

Dim ThePos As Integer

'In case the path contains a dot
Filename = ExtractFileName(Filename)

ThePos = InStrRev(Filename, ".")
If ThePos = 0 Then
    ExtractFileExtension = ""
Else
    ExtractFileExtension = Right$(Filename, Len(Filename) - ThePos)
End If


End Function

Private Sub Command2_Click()

End Sub

Private Sub cmdSaveAs_Click()


End Sub
Private Sub FileBox1_Click()
Dim sFile As String

sFile = LCase$(FileBox1.FileFullName)

Select Case ExtractFileExtension(sFile)

    Case "htm", "html", "txt", "jpg", "jpeg", "jif", "gif", "bmp"
        web1.Navigate sFile

    Case Else
         web1.Navigate "about:blank"

End Select


End Sub
Private Sub Form_Load()

web1.Offline = True
web1.Navigate "about:blank"

FileBox1.Filter = "All Files|*.*|HTML Files|*.htm;*.html"

End Sub
Private Sub Form_Resize()
On Error Resume Next

Splitter1.Top = Toolbar.Height
Splitter1.Left = 0
Splitter1.Width = ScaleWidth
Splitter1.Height = ScaleHeight - Toolbar.Height - StatusBar.Height

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
    
Select Case Button.Key
    Case "home"
    web1.GoHome
    
    Case "prev"
    web1.GoBack
    
    Case "next"
    web1.GoForward
    
    Case "refresh"
    web1.Refresh
    
    Case "stop"
    web1.Stop
    
    Case "save"
    web1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT

    Case Else

    
End Select

End Sub
Private Sub web1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Set HTMLpage = Nothing

End Sub

Private Sub web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    If pDisp Is web1.object Then
         'MsgBox "your document is ready"
        Set HTMLpage = web1.Document
    End If

End Sub


Private Sub web1_StatusTextChange(ByVal Text As String)
'Changes the text in the txtAddress to the current URL of page
'Changes the caption of the Browser Form to the current value of <TITLE>

'If Text <> "Done" Then
    StatusBar.Panels(1).Text = Text
'End If

'strLocName = wbBrowser.LocationName
'Browser.Caption = strLocName & conBrowser
    
'txtAddress.Text = wbBrowser.LocationURL
    


End Sub
