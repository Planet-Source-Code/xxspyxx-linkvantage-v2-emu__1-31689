VERSION 5.00
Begin VB.Form frmSplash1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1515
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   2025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2025
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   -10
      Picture         =   "frmSplash1.frx":000C
      ScaleHeight     =   1575
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   -10
      Width           =   2055
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
frmSplash.Visible = True
Unload Me

End Sub
