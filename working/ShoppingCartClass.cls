VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShoppingCartClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private main As Object
Public CommerceAccountMemberID As Long
Public AllowGroupPolicies As Boolean
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
    Call HandleError("ShoppingCartClass", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("ShoppingCartClass", MethodName, Description, ErrorNumber)
End Sub
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(mainObject As Object)
    Set main = mainObject
    CommerceAccountMemberID = main.memberId
    AllowGroupPolicies = main.IsAuthenticated
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
    HandleClassErrorAndResumeNext ("Upgrade"): Resume Next
End Function
'
'=================================================================================
'   Aggregate Object Interface
'=================================================================================
'
Public Function GetContent(optionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim MethodName As String
    'Dim SourceFormID As Long
    'Dim CatalogFormID As Long
    Dim srcShopFormId As Long
    Dim dstShopFormId As Long
    'Dim AccountFormID As Long
    Dim s As String
    Dim Cat As CatalogClass2
    Dim Cmc As CommerceClass2
    Dim Acct As CommerceClass2
    Dim AddItemID As Long
    Dim qs As String
    Dim PageBody As String
    Dim PageTabs As String
    Dim Pricing As New CmrcPricingClass
    Dim IsMemberCommerceAdmin As Boolean
    Dim CommerceAccountMemberName As String
    Dim IsCatalogSearch As Boolean
    Dim SetAccountMemberID As Long
    Dim adminFooter As String
    Dim isCheckoutForm As Boolean
    Dim isAccountForm As Boolean
    Dim isCatalogForm As Boolean
    
    '
    MethodName = "GetContent"
    '
    Call Pricing.GetRightsPolicies(main, main.memberId, AllowGroupPolicies, IsMemberCommerceAdmin)
    AddItemID = main.GetStreamInteger(RequestNameCatalogOrder)
    If AddItemID <> 0 Then
        srcShopFormId = 0
        dstShopFormId = shopFormCheckoutReview
    Else
        srcShopFormId = main.GetStreamInteger(rnSrcShopFormId)
        dstShopFormId = main.GetStreamInteger(rnDstShopFormId)
        If dstShopFormId = 0 Then
            '
            ' old hardcoded formid
            '
            dstShopFormId = main.GetStreamInteger("catalogFormId")
        End If
    End If
    '
    If IsMemberCommerceAdmin Then
        '
        ' CommerceAdmin - get or set the AccountMemberID
        '
        SetAccountMemberID = main.GetStreamInteger("SetAccountMemberid")
        If (SetAccountMemberID <> 0) Then
            CommerceAccountMemberID = SetAccountMemberID
        Else
            CommerceAccountMemberID = main.GetVisitProperty("CommerceAccountMemberID", 0)
            If CommerceAccountMemberID = 0 Then
                CommerceAccountMemberID = main.memberId
            End If
        End If
        Call main.SetVisitProperty("CommerceAccountMemberID", CommerceAccountMemberID)
    End If
    '
    isCatalogForm = False
    isCheckoutForm = False
    isAccountForm = False
    If (AddItemID <> 0) Then
        isCheckoutForm = True
    ElseIf (srcShopFormId > 0) And (srcShopFormId < 20) Then
        isCatalogForm = True
    ElseIf (srcShopFormId >= 20) And (srcShopFormId < 30) Then
        isCheckoutForm = True
    ElseIf (srcShopFormId >= 30) Then
        isAccountForm = True
    ElseIf (dstShopFormId > 0) And (dstShopFormId < 20) Then
        isCatalogForm = True
    ElseIf (dstShopFormId >= 20) And (dstShopFormId < 30) Then
        isCheckoutForm = True
    ElseIf (dstShopFormId >= 30) Then
        isAccountForm = True
    End If
    '
    If isCheckoutForm Then
        '
        ' checkout
        '
        Set Cmc = New CommerceClass2
        Call Cmc.Init(main)
        Cmc.CommerceAccountMemberID = CommerceAccountMemberID
        Cmc.AllowGroupPolicies = AllowGroupPolicies
        Cmc.IsMemberCommerceAdmin = IsMemberCommerceAdmin
        Cmc.IsPartOfShoppingCart = True
        PageBody = Cmc.getCheckout(srcShopFormId, dstShopFormId)
        qs = ModifyQueryString(main.RefreshQueryString, rnSrcShopFormId, shopFormCheckoutReview)
        PageTabs = GetShoppingTabs(qs, main.IsOrderOK(), main.IsAuthenticated(), True, False, False, False, IsMemberCommerceAdmin)
        Set Cmc = Nothing
    ElseIf isAccountForm Then
        '
        ' account
        '
        Set Acct = New CommerceClass2
        Call Acct.Init(main)
        Acct.CommerceAccountMemberID = CommerceAccountMemberID
        Acct.AllowGroupPolicies = AllowGroupPolicies
        Acct.IsMemberCommerceAdmin = IsMemberCommerceAdmin
        PageBody = Acct.GetAccountForm
        PageTabs = GetShoppingTabs(main.RefreshQueryString, main.IsOrderOK(), main.IsAuthenticated(), False, False, True, False, IsMemberCommerceAdmin)
        Set Cmc = Nothing
    Else
        '
        ' catalog
        '
        Set Cat = New CatalogClass2
        Call Cat.Init(main)
        Cat.CommerceAccountMemberID = CommerceAccountMemberID
        Cat.AllowGroupPolicies = AllowGroupPolicies
        Cat.IsMemberCommerceAdmin = IsMemberCommerceAdmin
        PageBody = Cat.GetCatalog2(srcShopFormId, dstShopFormId)
        IsCatalogSearch = (dstShopFormId = shopFormCatalogSearch)
        PageTabs = GetShoppingTabs(main.RefreshQueryString, main.IsOrderOK(), main.IsAuthenticated(), False, (Not IsCatalogSearch), False, IsCatalogSearch, IsMemberCommerceAdmin)
        Set Cat = Nothing
    End If
    '
    '
    If IsMemberCommerceAdmin Then
        '
        ' Create adminFooter so admin will know what mode they are in
        '
        If CommerceAccountMemberID = 0 Then
            adminFooter = adminFooter & "<div>Account:&nbsp;(not selected)</div>"
        Else
            CommerceAccountMemberName = main.GetRecordName("people", CommerceAccountMemberID)
            adminFooter = adminFooter & "<div>Account:&nbsp;" & CommerceAccountMemberName & "</div>"
        End If
        adminFooter = "" _
            & vbCrLf & vbTab & "<h2>Commerce Administrator</h2>" _
            & adminFooter _
            & ""
        adminFooter = "" _
            & vbCrLf & vbTab & "<div class=ccEditWrapperCaption>" _
            & kmaIndent(adminFooter) _
            & vbCrLf & vbTab & "</div>" _
            & ""
        'adminFooter = adminFooter & "</div><div style=""height:10px;""></div>"
    End If
    '
    ' Put page together
    '
    GetContent = "" _
        & vbCrLf & vbTab & "<div class=""shoppingCartCon"">" _
        & kmaIndent(PageTabs) _
        & kmaIndent(PageBody) _
        & kmaIndent(adminFooter) _
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
'
'
'
Public Function GetShoppingLinks(RefreshQueryString As String, OrderOK As Boolean, IsAuthenticated As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim qs As String
    Dim QSBase As String
    Dim s As String
    '
    s = "<div align=right>"
    QSBase = RefreshQueryString
    QSBase = ModifyQueryString(QSBase, rnSrcShopFormId, "", False)
    QSBase = ModifyQueryString(QSBase, "method", "", False)
    '
    ' Order Process
    '
    If OrderOK Then
        '
        ' link to Order Process
        '
        qs = QSBase
        qs = ModifyQueryString(qs, rnSrcShopFormId, shopFormCheckoutReview, True)
        s = s & "&nbsp;|&nbsp;<a href=""?" & qs & """ Class=ccCatMenuLink>Checkout</a>"
    End If
    '
    ' link to shopping
    '
    qs = QSBase
    s = s & "&nbsp;|&nbsp;<a href=""?" & qs & """ Class=ccCatMenuLink>Shopping</a>"
    '
    ' Authentication
    '
    If IsAuthenticated Then
        '
        ' link to Account forms
        '
        qs = QSBase
        qs = ModifyQueryString(qs, rnSrcShopFormId, "-1", True)
        s = s & "&nbsp;|&nbsp;<a href=""?" & qs & """ Class=ccCatMenuLink>Account</a>"
        '
        ' logout
        '
        qs = RefreshQueryString
        qs = ModifyQueryString(qs, "method", "logout", True)
        s = s & "&nbsp;|&nbsp;<a href=""?" & qs & """ Class=ccCatMenuLink>Logout</a>"
    Else
        '
        ' login
        '
        qs = RefreshQueryString
        qs = ModifyQueryString(qs, "method", "login", True)
        s = s & "&nbsp;|&nbsp;<a href=""?" & qs & """ Class=ccCatMenuLink>Login</a>"
    End If
    '
    ' link to Search forms
    '
    qs = QSBase
    qs = ModifyQueryString(qs, rnSrcShopFormId, shopFormCatalogSearch, True)
    s = s & "&nbsp;|&nbsp;<a href=""?" & qs & """ Class=ccCatMenuLink>Search</a>"
    '
    s = s & "</div>"
    '
    GetShoppingLinks = s
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetShoppingLinks"): Resume Next
End Function
'
'
'
Public Function GetShoppingTabs(RefreshQueryString As String, OrderOK As Boolean, IsAuthenticated As Boolean, IsOrderProcess As Boolean, IsShopping As Boolean, IsAccount As Boolean, isSearch As Boolean, IsMemberCommerceAdmin As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim qs As String
    Dim QSBase As String
    Dim s As String
    Dim Style As String
    '
    QSBase = RefreshQueryString
    QSBase = ModifyQueryString(QSBase, rnSrcShopFormId, "", False)
    QSBase = ModifyQueryString(QSBase, rnDstShopFormId, "", False)
    QSBase = ModifyQueryString(QSBase, RequestNameCatalogItemID, "", False)
    QSBase = ModifyQueryString(QSBase, "method", "", False)
    '
    ' link to shopping
    '
    qs = QSBase
    qs = ModifyQueryString(qs, "CatalogCatID", "", False)
    Style = "scNavItem"
    If IsShopping Then
        Style = Style & " scNavActive"
    End If
    s = s & vbCrLf & vbTab & "<li class=""" & Style & """><a href=""?" & qs & """>Shop</a></li>"
    's = s & "&nbsp;|&nbsp;<a href=""?" & qs & """ Class=ccCatMenuLink>Shopping</a>"
    'Call main.AddTabEntry("Shopping", "?" & qs, IsShopping)
    '
    ' link to Search forms
    '
    qs = QSBase
    qs = ModifyQueryString(qs, rnDstShopFormId, shopFormCatalogSearch, True)
    Style = "scNavItem"
    If isSearch Then
        Style = Style & " scNavActive"
    End If
    s = s & vbCrLf & vbTab & "<li class=""" & Style & """><a href=""?" & qs & """>Search</a></li>"
    'S = S & "&nbsp;|&nbsp;<a href=""?" & QS & """ Class=ccCatMenuLink>Search</a>"
    'Call main.AddTabEntry("Search", "?" & qs, isSearch)
    '
    '
    ' link to Account forms
    '
    qs = QSBase
    qs = ModifyQueryString(qs, rnDstShopFormId, shopFormAccountMenu, True)
    Style = "scNavItem"
    If IsAccount Then
        Style = Style & " scNavActive"
    End If
    s = s & vbCrLf & vbTab & "<li class=""" & Style & """><a href=""?" & qs & """>Account</a></li>"
    'S = S & "&nbsp;|&nbsp;<a href=""?" & QS & """ Class=ccCatMenuLink>Account</a>"
    'Call main.AddTabEntry("Account", "?" & qs, IsAccount)
    'If IsAuthenticated Then
    '    '
    '    ' logout
    '    '
    ''    QS = RefreshQueryString
    '    QS = ModifyQueryString(QS, "method", "logout", True)
    '    'S = S & "&nbsp;|&nbsp;<a href=""?" & QS & """ Class=ccCatMenuLink>Logout</a>"
    'Else
    '    '
    '    ' login
    '    '
    '    QS = RefreshQueryString
    '    QS = ModifyQueryString(QS, "method", "login", True)
    '    'S = S & "&nbsp;|&nbsp;<a href=""?" & QS & """ Class=ccCatMenuLink>Login</a>"
    'End If
    '
    ' Order Process
    '
    If OrderOK Then
    'If OrderOK Or IsMemberCommerceAdmin Then
        '
        ' link to Order Process
        '
        qs = QSBase
        qs = ModifyQueryString(qs, rnDstShopFormId, shopFormCheckoutReview, True)
        Style = "scNavItem"
        If IsOrderProcess Then
            Style = Style & " scNavActive"
        End If
        s = s & vbCrLf & vbTab & "<li class=""" & Style & """><a href=""?" & qs & """>Checkout</a></li>"
        'S = S & "&nbsp;|&nbsp;<a href=""?" & QS & """ Class=ccCatMenuLink>Checkout</a>"
        'Call main.AddTabEntry("Checkout", "?" & qs, IsOrderProcess)
    End If
'    '
'    'S = S & "</div>"
'    s = "" _
'        & vbCrLf _
'        & vbCrLf & "<!-- Tab Start -->" _
'        & vbCrLf & main.GetTabs() _
'        & vbCrLf & "<!-- Tab End -->" & vbCrLf & vbCrLf
    '
    GetShoppingTabs = "" _
        & vbCrLf & vbTab & "<ul class=""scNav"">" _
        & kmaIndent(s) _
        & vbCrLf & vbTab & "</ul>" _
        & ""
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetShoppingTabs"): Resume Next
End Function



