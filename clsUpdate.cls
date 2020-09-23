VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CoFreeUnusedLibraries Lib "ole32" () As Long
Private Const RegSvr32 = "regsvr32.exe"

Private pSourceLib As String
Private pDestLib As String
Private pBackupLib As String


Public Property Get SourceLib() As String
    SourceLib = pSourceLib
End Property
Public Property Let SourceLib(ByVal NewVal As String)
    If FileExists(NewVal) Then
        pSourceLib = NewVal
    Else
        Err.Raise vbObjectError, "DestLib", "File not found."
    End If
End Property

Public Property Get DestLib() As String
    DestLib = pDestLib
End Property
Public Property Let DestLib(ByVal NewVal As String)
    If FileExists(NewVal) Then
        pBackupLib = GetTempFileName(NewVal)
        pDestLib = NewVal
    Else
        Err.Raise vbObjectError, "DestLib", "File not found."
    End If
End Property

Public Property Get BackupLib() As String
    BackupLib = pBackupLib
End Property



Public Function Update() As Boolean
    
    If SourceLib = "" And DestLib = "" Then
        Err.Raise vbObjectError, "Update", "You must set the SourceLib and DestLib properties before using this function."
    End If
    
    On Error Resume Next
    
    'With out this function it could take
    'up to a minute for libraries to freeup
    CoFreeUnusedLibraries
    
    'Unregister DLL
    Shell RegSvr32 & " /u /s """ & DestLib & """"
    
    'Make Backup DLL
    FileCopy DestLib, BackupLib
    
    'Copy New DLL
    FileCopy SourceLib, DestLib
    
    'Register New DLL
    Shell RegSvr32 & " /s """ & DestLib & """"

    If Err = 0 Then
        Update = True
    Else
        Err.Clear
        Update = False
    End If

End Function

Public Function Rollback() As Boolean
    
    If SourceLib = "" And DestLib = "" Then
        Err.Raise vbObjectError, "Rollback", "You must set the SourceLib and DestLib properties before using this function."
    End If
    
    On Error Resume Next
    
    'With out this function it could take
    'up to a minute for libraries to freeup
    CoFreeUnusedLibraries
    
    'Unregister DLL
    Shell RegSvr32 & " /u /s """ & DestLib & """"
    
    'Copy Backup DLL
    FileCopy BackupLib, DestLib
    
    'Registery Backup DLL
    Shell RegSvr32 & " /s """ & DestLib & """"

    If Err = 0 Then
        Rollback = True
    Else
        Err.Clear
        Rollback = False
    End If

End Function



Private Function FileExists(ByVal FileName As String) As Boolean
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(FileName)
    Set fso = Nothing
End Function

Private Function GetTempFileName(ByVal fName As String) As String
    If InStr(fName, ".") > 0 Then
        fName = Left(fName, InStr(fName, ".") - 1) & "_backup"
    Else
        fName = fName & "_backup"
    End If
    GetTempFileName = fName
End Function
