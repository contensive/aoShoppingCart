Attribute VB_Name = "CommerceModule"

Option Explicit
'
'
Public Const shopCopyNameReview = "Order Process: Review Page"
Public Const shopCopyNameReviewEmpty = "Order Process: Review Page Empty Cart"
Public Const shopCopyNameSecurity = "Order Process: Security Page"
Public Const shopCopyNameLogin = "Order Process: Login Page"
Public Const shopCopyNameContact = "Order Process: Contact Page"
Public Const shopCopyNameContactName = "Order Process: Contact Page: Contact"
Public Const shopCopyNameContactAddress = "Order Process: Contact Page: Contact Address"
Public Const shopCopyNameContactAddressNoEdit = "Order Process: Contact Page: Contact Address No Edit"
Public Const shopCopyNameBillAddress = "Order Process: Contact Page: Billing Address"
Public Const shopCopyNameShipAddress = "Order Process: Contact Page: Shipping Address"
Public Const shopCopyNameShipMethod = "Order Process: Contact Page: Shipping Method"
Public Const shopCopyNameContactLogin = "Order Process: Contact Page: Login Information"
Public Const shopCopyNamePaymentMethod = "Order Process: Contact Page: Payment Method"
Public Const shopCopyNameContactAccountLogin = "Order Process: Contact Page: Login Account"
'Public Const shopCopyNameContactAccountLogin_legacy = "Order Process: Contact Page: Account"
Public Const shopCopyNameContactSubscriptionAccountLogin = "Order Process: Contact Page: Subscription Account"
Public Const shopCopyNameContactAccountLoginAuth = "Order Process: Contact Page: Account Logged in"
Public Const shopCopyNamePayment = "Order Process: Payment Page"
Public Const shopCopyNamePurchase = "Order Process: Purchase Page"
Public Const shopCopyNamePaymentDetailPrefix = "Order Process: Payment Detail: "
Public Const shopCopyNameReceipt = "Order Process: Receipt Page"
Public Const shopCopyNameReceptDetailPrefix = "Order Process: Receipt Form: Payment Type: "
Public Const shopCopyNameAccountEdit = "Commerce: Edit Account Page"
Public Const shopCopyNameAccountOrderList = "Commerce: Account Order List"
Public Const shopCopyNameAccountOrderDetails = "Commerce: Account Order Details"
Public Const shopCopyNameAccountNewCustomer = "Commerce: Account New Customer"
Public Const shopCopyNameAccountLogin = "Commerce: Account Login Page"
Public Const shopCopyNameAccountSendPassword = "Commerce: Account Send Password"
Public Const shopCopyNameAccountJoin = "Commerce: Account Join Page"
Public Const shopCopyNameAccountMain = "Commerce: Account Main"
Public Const shopCopyNameAccountSelectAccount = "Commerce: Commerce Admin Select Account"
Public Const shopCopyNameAccountAddAccount = "Commerce: Commerce Admin Add Account"
Public Const shopCopyNameCatalogHome = "Commerce: Catalog Home"
Public Const shopCopyNameAccountHome = "Commerce: Account Home"
Public Const shopCopyNameSearchHome = "Commerce: Search Home"
'
' default copy
'
Public Const DivBR = ""
'Public Const DivBR = "<div>&nbsp</div>"
Public Const shopCopyDefaultReviewEmpty = "<p>Your shopping cart is empty.</p>"
Public Const shopCopyDefaultReview = "<p>Please review the items in your cart. To checkout, use the Continue Checkout button.</p>"
Public Const shopCopyDefaultSecurity = ""
Public Const shopCopyDefaultLogin = "<p>If you are a returning customer, please enter your username and password. If you are a new customer, you do not need to create an account to purchase. To create a new account, select New under the accounts tab.</p>"
Public Const shopCopyDefaultContact = "<h2>Billing and Shipping</h2><p>Please complete the shipping and billing information for this order, then click Continue.</p>"
Public Const shopCopyDefaultContactName = ""
Public Const shopCopyDefaultContactAddress = "<p>Please enter your contact address. This is how we can contact you with any questions about the order.</p>"
Public Const shopCopyDefaultContactAddressNoEdit = "<p>This is your contact address. This is how we can contact you with any questions about the order. To edit this information, please contact the site administrator.</p>"
Public Const shopCopyDefaultBillAddress = "<p>Please use the address that matches your billing method. For instance, for credit cards, use the address that appears on your credit card statement.</p>"
Public Const shopCopyDefaultShipAddress = "<p>This is the address where your order will be shipped.</p>"
Public Const shopCopyDefaultShipMethod = "<p>Please select a shipping method.</p>"
Public Const shopCopyDefaultPaymentMethod = "<p>Please select your payment method.</p>"
Public Const shopCopyDefaultContactAccountLogin = "<p>Provide login information to secure your account.</p>"
Public Const shopCopyDefaultContactSubscriptionAccountLogin = "<p>An account is required for this purchase because your cart contains an item with a subscription.</p>"
Public Const shopCopyDefaultContactAccountLoginAuth = "<p>You are using an account that has been protected with a valid login. You can not modify this account without logging in. To manage this account, use the Account tab.</p>"
Public Const shopCopyDefaultPayment = "<h2>Payment Form</h2>"
Public Const shopCopyDefaultPurchase = "<h2>Order Review</h2><p>When you have reviewed your order and believe it to be correct, click Order Now to complete your order.</p>"
Public Const shopCopyDefaultPaymentDetailPrefix = ""
Public Const shopCopyDefaultReceipt = "<h2>Receipt</h2><p>This form is your receipt. Please print a copy for your records. If you have any questions and our like to contact us, please have this receipt available.</p>"
Public Const shopCopyDefaultReceptDetailPrefix = ""
Public Const shopCopyDefaultAccountEdit = "<h2>Edit Your Account</h2><p>Use this form to edit your account.</p>"
Public Const shopCopyDefaultAccountOrderList = "<h2>Completed Orders</h2><p>The following is a list of your completed orders.</p>"
Public Const shopCopyDefaultAccountOrderDetails = "<p>Your Order Details</p>"
Public Const shopCopyDefaultAccountNewCustomer = "<h2>New Customer?</h2><p>Please enter your email address to create a new account. If you already have an account you can request your password and log in below.</p>"
Public Const shopCopyDefaultAccountLogin = "<h2>Already Have an Account?</h2><p>To access your account, enter your username and password here. You can use the form below to have your password sent.</p>"
Public Const shopCopyDefaultAccountJoin = "<h2>New Customer?</h2><p>You do not need to create an account to purchase from this site. If you create an account, you can check on previous orders, and speed through checkout with prepopulated forms.</p>"
Public Const shopCopyDefaultAccountSendPassword = "<h2>Forgot Your Password?</h2><p>Enter your email address and your password will be emailed to you.</p>"
Public Const shopCopyDefaultAccountMain = "<p>Use this section to review or modify your account details, or check on a previous order.</p>"
Public Const shopCopyDefaultAccountSelectAccount = "<h2>Select an Account</h2><p>(Administrator Only) Select an account to use with the current order.</p>"
Public Const shopCopyDefaultAccountAddAccount = "<h2>Commerce Administrator</h2><p>Use this form to create a new account for your customer.</p>"
Public Const shopCopyDefaultCatalogHome = "<h2>Welcome</h2><p>Use the navigation to browse or search for items. Manage your account and previous orders from the Account section.</p>"
'Public Const shopCopyDefaultAccountHome = "<h2>Welcome to the catalog</h2><p>Use the navigation to browse or search for items. Review previous transactions from the Accounts tab.</p>"
Public Const shopCopyDefaultSearchHome = "<h2>Catalog Search</h2><p>Search the catalog by entering keywords and clicking search. Results matching all your keywords will be displayed.</p>"
''
'' ------------------------------------------------------------------------
'' Catalog
'' ------------------------------------------------------------------------
''
'Public Const CatalogIndexFormatGeneral = 0
'Public Const CatalogIndexFormatSpecials = 1
'
' ------------------------------------------------------------------------
' CatalogForm Data
' ------------------------------------------------------------------------
'
Public Const shopFormCatalogIndex = 10  ' Drill down index page
Public Const shopFormCatalogSearch = 11 ' search for items
Public Const shopFormCatalogListing = 12 ' List of products within a crieria
Public Const shopFormCatalogDetails = 13 ' details of one product
'
Public Const shopFormCheckoutReview = 21        '
Public Const shopFormCheckoutLogin = 22         '
Public Const shopFormCheckoutShipping = 23      '
Public Const shopFormCheckoutBill = 24          '
Public Const shopFormCheckoutPurchase = 25      '
Public Const shopFormCheckoutSecurity = 26      '
Public Const shopFormCheckoutReceipt = 27       '
Public Const shopFormCheckoutShippingBilling = 28
Public Const shopFormCheckoutAccount = 29       ' login or create an account (by entering a valid email address)
'
Public Const shopFormAccountLogin = 31               ' Login
Public Const shopFormAccountMenu = 32                ' Menu to account forms
Public Const shopFormAccountEdit = 33                ' Edit Information in the account
Public Const shopFormAccountOrderList = 34           ' List out all orders on this account
Public Const shopFormAccountOrderDetails = 35        ' Look at the details of one order
Public Const shopFormAccountJoin = 36                '
Public Const shopFormAccountSelect = 37  ' CommerceAdmin select a commerceMemberAcount
Public Const shopFormAccountAdd = 38     ' CommerceAdmin select a commerceMemberAcount
'
' order buttons
'
Public Const shopButtonSkip = "     Skip     "
Public Const shopButtonBackToShopping = " Return to Shopping "
Public Const shopButtonBack = "     Back     "
Public Const shopButtonContinue = "  Continue Checkout   "
Public Const shopButtonRecalculate = "  Recalculate  "
Public Const shopButtonRemoveAll = "  Remove All  "
Public Const shopButtonSecure = "    Secure    "
Public Const shopButtonUnsecure = "   Normal   "
Public Const shopButtonLogin = "    Login     "
Public Const shopButtonCreateAccount = "Create Account"
Public Const shopButtonSendPassword = "Send Password"
Public Const shopButtonOrderNow = "  Order Now  "
'
'------------------------------------------------------------------------------
' Form buttons
'------------------------------------------------------------------------------
'
Public Const scButtonCreateAccount = " Create Account "
Public Const AccountButtonLogin = "    Login    "
Public Const AccountButtonEmail = "Send Password"
Public Const AccountButtonSave = "    Save    "
Public Const AccountButtonCancel = "   Cancel   "
'
Public Const AccountActionSomething = 1 '
'
'
'
' Request Names
'
Public Const rnAccountUsername = "scAccountEmail"
Public Const rnContactUpdateGroupId = "scBlockContactUpdateGroup"
Private Const RequestNameFormID = "formid"
Private Const cmcFormCatalog = 0
'
Public Const rnSrcShopFormId = "srcShopFormId"
Public Const rnDstShopFormId = "dstShopFormId"
Public Const rnButton = "button"
'
'Public Const rnSrcShopFormId = "AccountFormID"
'
' spans
'
Public Const scSpanNormalStyle = "<span class=""scNormalFont"">"

'
'
'
Public Function GetShoppingCartLink(CurrentPage As String, RefreshQueryString As String) As String
    GetShoppingCartLink = CurrentPage
    If RefreshQueryString <> "" Then
        GetShoppingCartLink = GetShoppingCartLink & "?" & RefreshQueryString
    End If
End Function
'
'
'
Public Function GetFormColumnHeader(Title As String, Optional Width As String) As String
    '
    Dim copy As String
    '
'    If Title <> "" Then
'        Copy = "<P class=""ccAdminSmall"" align=""center"">" & Title & "</p>"
'    End If
    If Width = "" Then
        copy = "<td align=""center"" valign=""bottom"" class=""ccPanel"">" & copy & "</td>"
    Else
        copy = "<td align=""center"" valign=""bottom"" class=""ccPanel"" width=""" & Width & """>" & copy & "</td>"
    End If
    '
    GetFormColumnHeader = copy
    '
    End Function

'
'
'
Public Function GetFormFooterRow(FormInput As String, Caption As String) As String
    '
    GetFormFooterRow = "" _
        & "<tr><td width=""200"" align=""right"">" & FormInput & "</td>" _
        & "<td width=""100%"" align=""left"">&nbsp;" & Caption & "</td></tr>"
    '
    End Function


Public Function GetFormColumnTotal(copy As String)
    '
    GetFormColumnTotal = "" _
        & "<td align=""right"" class=""ccPanelRowOdd"">" & SpanClassAdminSmall _
        & "<img src=""/ccLib/images/black.gif"" width=""100%"" height=""1""><br>" _
        & copy & "</SPAN></td>"
    '
    End Function

'
'
'
Public Function GetFormRowStart() As String
    GetFormRowStart = "<tr>"
    End Function

'
'
'
Public Function GetFormRowEnd() As String
    GetFormRowEnd = "</tr>"
    End Function
