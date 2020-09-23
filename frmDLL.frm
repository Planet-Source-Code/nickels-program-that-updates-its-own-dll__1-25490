VERSION 5.00
Begin VB.Form frmDLL 
   Caption         =   "Form that uses the DLL with CreateObject()"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmDLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MyDLL

Private Sub Form_Load()
    Set MyDLL = CreateObject("UpdateDLL.SimpleClass")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set MyDLL = Nothing
End Sub
