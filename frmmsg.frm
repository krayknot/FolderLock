VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form frmmsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "How to use"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   Icon            =   "frmmsg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   4800
      TabIndex        =   9
      Top             =   0
      Width           =   1215
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
         Top             =   2160
         Width           =   1000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":000C
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":00A1
         TabIndex        =   2
         Top             =   480
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":0176
         TabIndex        =   3
         Top             =   840
         Width           =   4455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   3
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":0221
         TabIndex        =   4
         Top             =   1200
         Width           =   4335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   4
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":02DA
         TabIndex        =   5
         Top             =   1560
         Width           =   4215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   5
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":0355
         TabIndex        =   6
         Top             =   1800
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   375
         Index           =   7
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":042C
         TabIndex        =   7
         Top             =   2280
         Width           =   3975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Index           =   8
         Left            =   120
         OleObjectBlob   =   "frmmsg.frx":04FB
         TabIndex        =   8
         Top             =   2760
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmmsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Blskin = True Then
   Form1.Skin1.ApplySkin Me.hwnd
End If
End Sub

