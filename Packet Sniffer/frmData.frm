VERSION 5.00
Begin VB.Form frmData 
   Caption         =   "Raw Data"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   Icon            =   "frmData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5940
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtData 
      Height          =   4320
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   105
      Width           =   5685
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Resize()
    txtData.Top = 0
    txtData.Left = 0
    txtData.Width = Me.ScaleWidth
    txtData.Height = Me.ScaleHeight
End Sub
