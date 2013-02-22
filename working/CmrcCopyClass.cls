VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CmrcCopyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("CmrcCopyClass", MethodName, Description, ErrorNumber)
End Sub
'
'
'
Public Function GetCopy(main As Object, PageName As String, defaultCopy As String)
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Dim RecordID As Long
    '
    GetCopy = main.GetContentCopy2(PageName, "copy content", defaultCopy)
'    If GetCopy = "" Then
'        If GetCopy = "" Then
'            GetCopy = defaultCopy
'        End If
'        cs = main.OpenCSContent_Internal("Copy Content", "name=" & KmaEncodeSQLText(PageName))
'        If main.IsCSOK(cs) Then
'            RecordID = main.GetCSInteger(cs, "ID")
'            Call main.SetCS(cs, "Copy", GetCopy)
'        End If
'        Call main.CloseCS(cs)
'        Call main.PublishEdit("Copy Content", RecordID)
'    ElseIf GetCopy = PageName Then
'        GetCopy = defaultCopy
'        cs = main.OpenCSContent("Copy Content", "name=" & main.EncodeSQLText(PageName))
'        If main.IsCSOK(cs) Then
'            Call main.SetCS(cs, "Copy", GetCopy)
'        End If
'        Call main.CloseCS(cs)
'    End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCopy"): Resume Next
End Function



