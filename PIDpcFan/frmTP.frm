VERSION 5.00
Begin VB.Form frmTP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMD Thermal Profile"
   ClientHeight    =   6825
   ClientLeft      =   6660
   ClientTop       =   2445
   ClientWidth     =   9165
   Icon            =   "frmTP.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9165
   Begin VB.Image Image 
      Appearance      =   0  'Flat
      Height          =   6795
      Left            =   0
      Picture         =   "frmTP.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image_Click()
Unload frmTP

End Sub
