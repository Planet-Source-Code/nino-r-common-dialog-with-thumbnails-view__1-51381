VERSION 5.00
Begin VB.Form frmMainT 
   Caption         =   "Start Common Dialog with thumnails view by NR (demo for PSC)"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOSVersion 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtFilePath 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3240
      Width           =   5535
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   5535
   End
   Begin VB.ListBox lstViewMenu 
      Height          =   1035
      ItemData        =   "frmMainT.frx":0000
      Left            =   120
      List            =   "frmMainT.frx":0013
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdOFN 
      Caption         =   "Start "
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblOSVersion 
      Caption         =   "OS - Windows ?"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblViewMenu 
      Caption         =   "View menu list"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblFileName 
      Caption         =   "FileName"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblFilePath 
      Caption         =   "FilePath"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image imgShowPicture 
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmMainT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'starting Common Dialog with thumbnaisl view
'it works both on Win2000 and WinXP
'by Nino R.  http://ca.geocities.com/ninek_zg/
'this demo was made for PSC, 31.01.2004

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOFN_Click()
    Dim aPath As String
    Dim sDirActual As String
    Dim sFileName As String
    Dim sFilePath As String

On Error GoTo errorhandler

    sDirActual = App.Path
    'second parameter: if false will not show preview window
    'aPath = OpenPictureDialog(Me.hwnd, False)

    Select Case lstViewMenu.ListIndex
        Case 4
                meDesiredView = vsThumbnails
                aPath = OpenPictureDialog(Me.hwnd, False)
        Case 0
                meDesiredView = vsLargeIcons
                aPath = OpenPictureDialog(Me.hwnd, True)
        Case 1
                meDesiredView = vsSmallIcons
                aPath = OpenPictureDialog(Me.hwnd, True)
        Case 2
                meDesiredView = vsList
                aPath = OpenPictureDialog(Me.hwnd, True)
        Case 3
                meDesiredView = vsDetails
                aPath = OpenPictureDialog(Me.hwnd, True)
        Case Else   'this should never be
            MsgBox "error"
            Exit Sub
        End Select

    If aPath <> "" Then
        sFileName = ParsePath(aPath, FILE_ONLY)
        sFilePath = ParsePath(aPath, PATH_ONLY)

        txtFileName.Text = sFileName
        txtFilePath.Text = sFilePath
        'Text1.Text = aPath
        imgShowPicture.Picture = LoadPicture(aPath)
    End If
Exit Sub
errorhandler:
MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & _
    "source: cmdOFN_Click"

End Sub

Private Sub Form_Load()
    lstViewMenu.Selected(4) = True
    FindWindowsVersion
End Sub
