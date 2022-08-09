VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmmsgbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   Icon            =   "frmmsgbox.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "frmmsgbox.frx":000C
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmmsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Blskin = True Then
'   Form1.Skin1.ApplySkin Me.hwnd
End If
End Sub

