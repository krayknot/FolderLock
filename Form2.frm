VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   885
      Top             =   1890
   End
   Begin VB.Image Image1 
      Height          =   3150
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   -15
      Width           =   5985
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Timer1_Timer()
i = i + 1
If i = 5 Then
    Form1.Show
    Unload Me
End If
End Sub
