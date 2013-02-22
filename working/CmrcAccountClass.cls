VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CmrcAccountClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'
' ----- Global Objects
'
Private Main As Object
Private ClassInitialized As Boolean
'
Public CommerceAccountMemberID As Long
Public AllowGroupPolicies As Boolean
Public IsMemberCommerceAdmin As Boolean
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("CmrcAccountClass", MethodName, Description, ErrorNumber)
End Sub
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(MainPointer As Object)
    On Error GoTo ErrorTrap
    '
    Dim Pricing As New CmrcPricingClass
    '
    If Not ClassInitialized Then
        Set Main = MainPointer
        '
        ' --- common storage
        '
        CommerceAccountMemberID = Main.MemberID
        AllowGroupPolicies = Main.IsAuthenticated
        Call Pricing.GetRightsPolicies(Main, Main.MemberID, Main.IsAuthenticated, IsMemberCommerceAdmin)
        If IsMemberCommerceAdmin Then
            CommerceAccountMemberID = Main.GetVisitProperty("CommerceAccountMemberID", 0)
        End If
        ClassInitialized = True
    End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("Init"): Resume Next
    End Sub

