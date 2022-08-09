VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3525
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4230
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2433.018
   ScaleMode       =   0  'User
   ScaleWidth      =   3972.189
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   4215
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "Skin On"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2775
         TabIndex        =   10
         Top             =   600
         Width           =   1245
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1260
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   735
         Index           =   5
         Left            =   120
         OleObjectBlob   =   "frmAbout.frx":0000
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   0
         Left            =   840
         OleObjectBlob   =   "frmAbout.frx":0123
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   540
         Left            =   120
         Picture         =   "frmAbout.frx":0198
         ScaleHeight     =   337.12
         ScaleMode       =   0  'User
         ScaleWidth      =   337.12
         TabIndex        =   1
         Top             =   240
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   1
         Left            =   840
         OleObjectBlob   =   "frmAbout.frx":0E62
         TabIndex        =   3
         Top             =   480
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   2
         Left            =   840
         OleObjectBlob   =   "frmAbout.frx":0EEB
         TabIndex        =   4
         Top             =   1080
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   3
         Left            =   840
         OleObjectBlob   =   "frmAbout.frx":0F52
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   615
         Index           =   4
         Left            =   840
         OleObjectBlob   =   "frmAbout.frx":0FD3
         TabIndex        =   6
         Top             =   1680
         Width           =   3255
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   6
         Left            =   840
         OleObjectBlob   =   "frmAbout.frx":1100
         TabIndex        =   11
         Top             =   720
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdSysInfo_Click()
If cmdSysInfo.Caption = "Skin On" Then
   cmdSysInfo.Caption = "Skin Off"
   Blskin = True
   Unload Form1
   Unload Me
   Form1.Show
Else
   cmdSysInfo.Caption = "Skin On"
   Blskin = False
End If
End Sub

Private Sub Form_Load()
If Blskin = True Then
   Form1.Skin1.ApplySkin Me.hwnd
End If
Me.Caption = "About " & App.Title
End Sub


