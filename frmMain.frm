VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Internet Backround"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Locate Your Picture File"
      Height          =   4335
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   5295
      Begin VB.FileListBox filList 
         Height          =   3795
         Left            =   2760
         Pattern         =   "*.bmp"
         TabIndex        =   8
         Top             =   360
         Width           =   2415
      End
      Begin VB.DirListBox DirList 
         Height          =   3465
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.DriveListBox drvList 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Location"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set Image"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtLocation 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   5040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   240
      X2              =   5040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmMain.frx":0ABA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    If DirList.Path <> DirList.List(DirList.ListIndex) Then
       DirList.Path = DirList.List(DirList.ListIndex)
       Exit Sub
    End If
    
    
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdClear_Click()
txtLocation.Text = ""
End Sub

Private Sub cmdDefault_Click()
On Error Resume Next
txtLocation.Text = "(Default)"
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", txtLocation.Text
txtLocation.Text = ""
MsgBox "Default Backround Set", vbInformation, "Info"
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdSet_Click()
On Error Resume Next
SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", txtLocation.Text
txtLocation.Text = ""
End Sub

Private Sub dirList_Change()
 On Error Resume Next
    DirList.Path = drvList.Drive
    ChDir DirList.Path
End Sub

Private Sub dirList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    filList.Path = DirList.Path
End Sub

Private Sub drvList_Change()
    On Error GoTo Drivehandler
    DirList.Path = drvList.Drive
    Exit Sub
Drivehandler:
    drvList.Drive = DirList.Path
    Exit Sub
End Sub

Private Sub filList_Change()
    
    On Error GoTo Drivehandler
    filList.Path = DirList.Path
    Exit Sub
Drivehandler:
    DirList.Path = filList.Path
    Exit Sub
    filList.Pattern = "*.bmp"
End Sub



Private Sub filList_Click()
On Error Resume Next

txtLocation.Text = filList.Path
txtLocation = txtLocation + "\" & filList.FileName


End Sub

Private Sub filList_DblClick()
On Error Resume Next

SetStringValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap", txtLocation.Text
txtLocation.Text = ""
End Sub

Private Sub filList_PathChange()
On Error Resume Next

txtLocation.Text = filList.Path
End Sub

Private Sub Form_Load()
On Error Resume Next
'On Error GoTo Err
txtLocation.Text = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar", "BackBitmap")
'Err: txtLocation.Text = "No Value assigned!"
End Sub
