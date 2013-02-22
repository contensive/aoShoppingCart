VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderProcessClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private main As Object
Public CommerceAccountMemberID As Long
Public AllowGroupPolicies As Boolean
Public IsMemberCommerceAdmin As Boolean
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("OrderProcessClass", MethodName, Description, ErrorNumber)
End Sub
'
'===========================================================================================================
'
'===========================================================================================================
'
Public Function Execute(CsvObject As Object, mainObject As Object, optionString As String, FilterInput As String)
    '
    Call Init(mainObject)
    Execute = GetContent(optionString)
    '
    Exit Function
ErrorTrap:
    Call HandleError("OrderProcessClass", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(mainObject As Object)
    '
    Dim Pricing As New CmrcPricingClass
    '
    Set main = mainObject
    '
    ' --- common storage
    '
    CommerceAccountMemberID = main.memberId
    AllowGroupPolicies = main.IsAuthenticated
    Call Pricing.GetRightsPolicies(main, main.memberId, main.IsAuthenticated, IsMemberCommerceAdmin)
    If IsMemberCommerceAdmin Then
        CommerceAccountMemberID = main.GetVisitProperty("CommerceAccountMemberID", 0)
    End If
    
    
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("Init"): Resume Next
End Sub
'
'=================================================================================
'
'=================================================================================
'
Public Function Upgrade() As String
    On Error GoTo ErrorTrap
    '
    Dim MethodName As String
    '
    MethodName = "Upgrade"
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("Upgrade"): Resume Next
End Function
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(optionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim content As String
    Dim Cmc As CommerceClass2
    '
    Set Cmc = New CommerceClass2
    Call Cmc.Init(main)
    Cmc.IsPartOfShoppingCart = False
    content = Cmc.GetOrderProcess
    Set Cmc = Nothing
    '
    ' Put page together
    '
    GetContent = "" _
        & vbCrLf & vbTab & "<div class=""shoppingCartCon"">" _
        & kmaIndent(content) _
        & vbCrLf & vbTab & "</div>" _
        & ""
    '
    ' change ccButtonCon to an editable style
    '
    GetContent = Replace(GetContent, "ccButtonCon", "scButtonCon", , , vbTextCompare)
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetContent"): Resume Next
End Function




