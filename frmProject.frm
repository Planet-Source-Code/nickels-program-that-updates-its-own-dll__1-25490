VERSION 5.00
Begin VB.Form frmProject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update DLL Example"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3030
      TabIndex        =   8
      Top             =   1005
      Width           =   5460
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3030
      TabIndex        =   7
      Top             =   405
      Width           =   5460
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Rollback to v1.0.0"
      Height          =   360
      Index           =   4
      Left            =   5760
      TabIndex        =   4
      Top             =   1575
      Width           =   2250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update to v1.0.1"
      Height          =   360
      Index           =   3
      Left            =   3345
      TabIndex        =   3
      Top             =   1575
      Width           =   2250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Unload DLL"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   165
      TabIndex        =   2
      Top             =   990
      Width           =   2250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use DLL"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   165
      TabIndex        =   1
      Top             =   570
      Width           =   2250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load DLL"
      Height          =   360
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   2250
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmProject.frx":0000
      Height          =   930
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   7800
   End
   Begin VB.Label Label4 
      Caption         =   "Current DLL File"
      Height          =   240
      Left            =   3075
      TabIndex        =   6
      Top             =   765
      Width           =   3060
   End
   Begin VB.Label Label3 
      Caption         =   "New DLL File (v1.0.1)"
      Height          =   240
      Left            =   3075
      TabIndex        =   5
      Top             =   165
      Width           =   3060
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click(Index As Integer)
    Dim up As New clsUpdate
    up.SourceLib = Text1(1).Text
    up.DestLib = Text1(2).Text
    
    Select Case Index
        Case 0 'load dll
            Load frmDLL
            Command1(0).Enabled = False
            Command1(1).Enabled = True
            Command1(2).Enabled = False
            Command1(3).Enabled = False
            Command1(4).Enabled = False
            Text1(1).Enabled = False
            Text1(2).Enabled = False
        
        Case 1 'use dll
            MsgBox frmDLL.MyDLL.VersionInfo
            Command1(0).Enabled = False
            Command1(1).Enabled = False
            Command1(2).Enabled = True
            Command1(3).Enabled = False
            Command1(4).Enabled = False
        
        Case 2 'unload dll
            Unload frmDLL
            Command1(0).Enabled = True
            Command1(1).Enabled = False
            Command1(2).Enabled = False
            Text1(1).Enabled = True
            Text1(2).Enabled = True
            
            SetUpdateButtons
            
        Case 3 'update dll
            
            up.Update
                    
            SetUpdateButtons
                    
            MsgBox "Updated!"
        
        Case 4 'rollback dll
            
            up.Rollback
                    
            SetUpdateButtons
                    
            MsgBox "Rollback!"
            
    End Select
    
    Set up = Nothing

End Sub

Private Sub Form_Load()
    Command1(0).Enabled = True
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    
    SetUpdateButtons
    
    Text1(1).Text = App.Path & "\DLL Code\UpdateDLLv101_dll"
    Text1(2).Text = App.Path & "\DLL Code\UpdateDLL.dll"
End Sub

Private Sub SetUpdateButtons()
    Load frmDLL
    If frmDLL.MyDLL.VersionInfo = "1.0.0" Then
        Command1(3).Enabled = True
        Command1(4).Enabled = False
    Else
        Command1(3).Enabled = False
        Command1(4).Enabled = True
    End If
    Unload frmDLL
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
