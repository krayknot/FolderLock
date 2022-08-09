VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Folder Lock"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "lockfolder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1680
      Top             =   5400
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   3360
   End
   Begin VB.ComboBox Combo1 
      BeginProperty DataFormat 
         Type            =   4
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   8
      EndProperty
      Height          =   315
      ItemData        =   "lockfolder.frx":2CFA
      Left            =   240
      List            =   "lockfolder.frx":2D10
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   4920
      Left            =   5520
      TabIndex        =   6
      Top             =   -50
      Width           =   1215
      Begin VB.CommandButton Command2 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   1000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   1000
      End
      Begin VB.CommandButton cmdhelp 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1000
      End
      Begin VB.CommandButton cmdunlock 
         Caption         =   "&Unlock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1000
      End
      Begin VB.CommandButton cmdlock 
         Caption         =   "&Lock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   435
         Width           =   1000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4920
      Left            =   20
      TabIndex        =   0
      Top             =   -50
      Width           =   5535
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   1680
         OleObjectBlob   =   "lockfolder.frx":2D71
         Top             =   3240
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   " "
         Top             =   4440
         Width           =   5295
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   2640
         ScaleHeight     =   3945
         ScaleWidth      =   2745
         TabIndex        =   12
         Top             =   360
         Width           =   2775
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   0
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   1215
            TabIndex        =   4
            Top             =   1080
            Width           =   1215
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   1
            Left            =   1440
            ScaleHeight     =   855
            ScaleWidth      =   1215
            TabIndex        =   17
            Top             =   1080
            Width           =   1215
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   2
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   1215
            TabIndex        =   16
            Top             =   2040
            Width           =   1215
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   3
            Left            =   1440
            ScaleHeight     =   855
            ScaleWidth      =   1215
            TabIndex        =   15
            Top             =   2040
            Width           =   1215
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   4
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   1215
            TabIndex        =   14
            Top             =   3000
            Width           =   1215
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   5
            Left            =   1440
            ScaleHeight     =   855
            ScaleWidth      =   1215
            TabIndex        =   13
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   975
            Index           =   1
            Left            =   120
            Picture         =   "lockfolder.frx":5C436
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.DriveListBox drvDrive 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.DirListBox dirDir 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3240
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   5
      Left            =   6840
      Picture         =   "lockfolder.frx":5EA53
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   4
      Left            =   6840
      Picture         =   "lockfolder.frx":6174D
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   3
      Left            =   6840
      Picture         =   "lockfolder.frx":62017
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   2
      Left            =   6840
      Picture         =   "lockfolder.frx":628E1
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   1
      Left            =   6840
      Picture         =   "lockfolder.frx":645DB
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Index           =   0
      Left            =   6840
      Picture         =   "lockfolder.frx":672D5
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim WindowsDirectory As String
Dim strlock As String
Dim counttext As String
Dim lockcount As Integer

Dim Bldown As Boolean, BlUp As Boolean
Dim i As Integer

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim lpPoint As POINTAPI, mHwnd As Long, lHwnd As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Dim IntIndex As Integer
Dim IntSelect As Integer

Private Sub cmdhelp_Click()
frmmsg.Show vbModal
End Sub

Private Sub cmdlock_Click()
On Error GoTo Err
    Dim Path As String
    Dim Data As String
    Dim File As String
    Dim Ext As String
    Dim filename As String

    
    'My computer
If IntSelect = 0 Then
    Ext = ".{21EC2020-3AEA-1069-A2DD-08002B30309C}"
'    Ext = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    Path = dirDir.Path
    Data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(Data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(Data) = UCase("desktop") Then
        
        filename = File & Data & Ext
        Name dirDir.Path As filename
        dirDir.Path = File
        frmmsgbox.SkinLabel1.Caption = "Folder has been Locked"
        frmmsgbox.Show vbModal
    Else
        frmmsgbox.SkinLabel1.Caption = "Folder cannot be Locked. Sorry..."
        frmmsgbox.Show vbModal
    End If
End If

    'Recycle Bin
If IntSelect = 1 Then
 Ext = ".{645FF040-5081-101B-9F08-00AA002F954E}"
    Path = dirDir.Path
    Data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(Data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(Data) = UCase("desktop") Then
        
        filename = File & Data & Ext
        Name dirDir.Path As filename
        dirDir.Path = File
          frmmsgbox.SkinLabel1.Caption = "Folder has been Locked"
          frmmsgbox.Show vbModal
          Else
        frmmsgbox.SkinLabel1.Caption = "Folder cannot be Locked. Sorry..."
        frmmsgbox.Show vbModal
    End If
End If

    'Control Panel
If IntSelect = 2 Then
    Ext = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
    Path = dirDir.Path
    Data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(Data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(Data) = UCase("desktop") Then
        
        filename = File & Data & Ext
        Name dirDir.Path As filename
        dirDir.Path = File
        frmmsgbox.SkinLabel1.Caption = "Folder has been Locked"
        frmmsgbox.Show vbModal
    Else
        frmmsgbox.SkinLabel1.Caption = "Folder cannot be Locked. Sorry..."
        frmmsgbox.Show vbModal
    End If
End If

    'Dial Up Networking
If IntSelect = 3 Then
    Ext = ".{992CFFA0-F557-101A-88EC-00DD010CCC48}"
    Path = dirDir.Path
    Data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(Data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(Data) = UCase("desktop") Then
        
        filename = File & Data & Ext
        Name dirDir.Path As filename
        dirDir.Path = File
        frmmsgbox.SkinLabel1.Caption = "Folder has been Locked"
        frmmsgbox.Show vbModal
    Else
        frmmsgbox.SkinLabel1.Caption = "Folder cannot be Locked. Sorry..."
        frmmsgbox.Show vbModal
    End If
End If

    'Printers
If IntSelect = 4 Then
    Ext = ".{2227A280-3AEA-1069-A2DE-08002B30309D}"
    Path = dirDir.Path
    Data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(Data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(Data) = UCase("desktop") Then
        
        filename = File & Data & Ext
        Name dirDir.Path As filename
        dirDir.Path = File
        frmmsgbox.SkinLabel1.Caption = "Folder has been Locked"
        frmmsgbox.Show vbModal
    Else
        frmmsgbox.SkinLabel1.Caption = "Folder cannot be Locked. Sorry..."
        frmmsgbox.Show vbModal
    End If
    
    'Network Neighborhood
If IntSelect = 5 Then
    Ext = ".{208D2C60-3AEA-1069-A2D7-08002B30309D}"
    Path = dirDir.Path
    Data = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    File = Left$(Path, Len(Path) - Len(Data))
    If Not UCase$(Path) = UCase$(WindowsDirectory) _
    And Not UCase$(Data) = UCase("desktop") Then
        
        filename = File & Data & Ext
        Name dirDir.Path As filename
        dirDir.Path = File
        frmmsgbox.SkinLabel1.Caption = "Folder has been Locked"
        frmmsgbox.Show vbModal
    Else
        frmmsgbox.SkinLabel1.Caption = "Folder cannot be Locked. Sorry..."
        frmmsgbox.Show vbModal
    End If
 End If
End If

Err:
Me.Caption = Err.Description
Exit Sub
End Sub

Private Sub cmdunlock_Click()
On Error GoTo Err
    Dim Path As String
    Dim Temp As String
    Dim Data As String
    Dim File As String
    Dim Ext As String
    Dim filename As String
    Path = dirDir.Path
    Temp = Mid$(Path, InStrRev(Path, "\") + 1, Len(Path))
    Data = Left$(Temp, InStr(Temp, ".{") - 1)
    File = Left$(Path, Len(Path) - Len(Temp))
    filename = File & Data
    Name dirDir.Path As filename
    dirDir.Path = File
    frmmsgbox.SkinLabel1.Caption = "Folder has been UnLocked"
    frmmsgbox.Show vbModal

Err:
Me.Caption = Err.Description
    Exit Sub
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
frmAbout.Show vbModal
End Sub

Private Sub dirDir_Change()
Text1.Text = dirDir.Path
dirDir.Path = drvDrive.Drive
Me.Caption = dirDir.Path
End Sub

Private Sub drvDrive_Change()
On Error GoTo NotReady
    dirDir.Path = drvDrive.Drive
    Exit Sub
NotReady:
    MsgBox "Drive is not ready.", vbExclamation + vbApplicationModal, "Not Ready..."
End Sub

Private Sub Form_Initialize()
Combo1.ListIndex = 0

End Sub

Private Sub Form_Load()
On Error Resume Next
If Blskin = True Then
   Skin1.ApplySkin Me.hwnd
End If

Dim ret As Long
Dim buff As String
    buff = Space(255)
    ret = GetWindowsDirectory(buff, 255)
    WindowsDirectory = Left$(buff, InStr(buff, vbNullChar) - 1)
    Text1.Text = dirDir.Path
End Sub

Private Sub Picture9_Click(Index As Integer)
IntSelect = Index
End Sub

Private Sub Picture9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
IntIndex = Index
End Sub

Private Sub Timer1_Timer()
Dim i
GetCursorPos lpPoint
mHwnd = WindowFromPoint(lpPoint.X, lpPoint.Y)

If mHwnd = Picture9(IntIndex).hwnd Then
   Picture9(IntIndex).Cls
   Picture9(IntIndex).PaintPicture Image4(IntIndex).Picture, 250, 250, 600, 600
   Picture9(IntIndex).CurrentX = 50
   Picture9(IntIndex).CurrentY = 50
   Picture9(IntIndex).Font.Name = "Tahoma"
   Picture9(IntIndex).FontBold = True
   If IntIndex = 0 Then
          Picture9(IntIndex).Print "My Computer"
   ElseIf IntIndex = 1 Then
          Picture9(IntIndex).Print "Recycle Bin"
   ElseIf IntIndex = 2 Then
          Picture9(IntIndex).Print "Control Panel"
   ElseIf IntIndex = 3 Then
          Picture9(IntIndex).Print "Dial-Up"
   ElseIf IntIndex = 4 Then
          Picture9(IntIndex).Print "Printers"
   ElseIf IntIndex = 5 Then
          Picture9(IntIndex).Print "Network"
   End If
Else
   For i = 0 To 5
     If mHwnd <> Picture9(i).hwnd Then
     Picture9(i).Cls
     Picture9(i).PaintPicture Image4(i).Picture, 220, 50, 500, 500
     Picture9(i).CurrentX = 50
     Picture9(i).CurrentY = 600
     Picture9(i).FontBold = False
     Picture9(i).Font.Name = "Tahoma"
     End If
   Next i
   
    Picture9(0).Print "My Computer"
    Picture9(1).Print "Recycle Bin"
    Picture9(2).Print "Control Panel"
    Picture9(3).Print "Dial- Up"
    Picture9(4).Print "Printers"
    Picture9(5).Print "Network"
End If
End Sub

Private Sub Timer2_Timer()
Me.Caption = "Folder Lock"
End Sub
