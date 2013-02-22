VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommerceClass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'========================================================================
' This file and its contents are copyright by Kidwell McGowan Associates.
'
' Notes on the architecture
'
'   CommerceAccountMemberID - the MemberID of the person on the Order
'   AllowGroupPolicies - if true, it is OK to calculate pricing discounts based on CommerceAccountMemberID
'   IsMemberCommerceAdmin - when true, the current member is a commerce admin
'
'   LoadOrderGlobals
'       - takes values out of OrderDetails into OrderDetails() array
'       - verifies that the order's memberid matches the commerceAccountMemberId
'           - if not, it loads defaults from the member into the order and saves
'
'   CalculateAndSaveOrder - saves the OrderDetails based on only
'       OrderDetails().Quantity
'       OrderDetails().ItemID
'
'   non-complete orders are tracked with visitor. VisitorOrderID is not
'   0 if there is a current, non-complete, non-empty order for this
'   visitor.
'   complete orders are tracked to the member record.
'
'   basically, to complete an order, you have to have an account.
'   you shop in your visitor space, but completed orders are only
'   accessible to members.
'
' PrintOrderForms
'
'   run on CommerceForm / CommerceSourceForm variable
'       Read CommerceForm
'           If Given, go there
'           If Not Given, Read CommerceButton and CommerceSourceForm
'               If Given, process them to create CommerceForm
'               If Not Given, go to Shopping cart form with visitors last open order
'========================================================================
'
' ----- Global Child Objects
'
' none
'
' ----- Global Objects
'
Private main As Object
'
Public CommerceAccountMemberID As Long
Public AllowGroupPolicies As Boolean
Public IsMemberCommerceAdmin As Boolean
Public IsPartOfShoppingCart As Boolean
'
' ----- global scope variables
'
Private ClassInitialized        ' if true, the module has been
Private CommerceDataSourceName As String
'
' ----- Initilized during Init to be used throughout
'
Public isSupportedAddress2 As Boolean
Public orderId As Long
Private OrderMemberID As Long
'
' Order Contact
'
Private orderContactName As String
Private orderContactEmail As String
Private orderContactPhone As String
Private orderContactFax As String
Private orderContactCompany As String
Private orderContactAddress As String
Private orderContactAddress2 As String
Private orderContactCity As String
Private orderContactState As String
Private orderContactZip As String
Private orderContactCountry As String
Private orderContactConfirmed As Boolean
'Private IsOrderEmailBad As Boolean
'
' Order Billing
'
Private orderBillingRequired As Boolean
Private orderBillName As String
Private orderBillEmail As String
Private orderBillPhone As String
Private orderBillFax As String
Private orderBillCompany As String
Private orderBillAddress As String
Private orderBillAddress2 As String
Private orderBillCity As String
Private orderBillState As String
Private orderBillZip As String
Private orderBillCountry As String
Private orderBillConfirmed As Boolean
Private IsOrderEmailBad As Boolean
'
' Order Shipping
'
Private orderShippingRequired As Boolean
Private orderShipDate As Date
Private orderShipName As String
Private orderShipCompany As String
Private orderShipAddress As String
Private orderShipAddress2 As String
Private orderShipCity As String
Private orderShipState As String
Private orderShipZip As String
Private orderShipCountry As String
Private orderShipPhone As String
Private orderShipMethod As String
Private orderShipConfirmed As Boolean
Private orderShipWeight As Double
Private IsorderShipMethodBad As Boolean
'
Private OrderItemCharge As Double
Private orderShipCharge As Double
Private OrderTaxCharge As Double
Private OrdertotalCharge As Double
Private OrderPayMethod As String
Private OrderPayDate As Date
Private OrderCCType As String
Private OrderCCNumber As String
Private OrderCCName As String
Private OrderCCExpires As String
Private OrderCCZip As String
Private OrderCompleted As Boolean
Private OrderDateCompleted As Date
Private OrderActive As Boolean
Private OrderCreditCardApproved As Boolean
Private OrderSpecialPricing As Double
'
Private OrderPriceLevelFreeShipping As Boolean
Private OrderPriceLevelDiscount As Double
'
Private AccountUsername As String
Private accountPassword As String
Private AccountUsernameError As Boolean
Private AccountPasswordError As Boolean
'
' common to order process
'
Private OrderAction As Long         ' The action to be performed based on SourceForm / Button
'Private OrderForm As Long     ' the page that submitted the form
Private shopButton As String         ' The button pressed
'Private OrderForm As Long           ' the next form that should be displayed next
Private OrderForceForm As Long      ' make the process back go to a form
'Private ignore_UpdateMemberRecord As Boolean      ' signals if viewer would like thier account updated with the forms information
Private OrderMemberUsername As String     ' used during login only
Private OrderMemberPassword As String     ' used during login only
Private OrderMemberEmail As String        ' used during login only
'
'
'
Private OrderAllowCreditCards
Private OrderAllowVisa As Boolean
Private OrderAllowMC As Boolean
Private OrderAllowDiscover As Boolean
Private OrderAllowAmex As Boolean
Private OrderAllowNetTerms As Boolean
Private OrderAllowCODCompanyCheck As Boolean
Private OrderAllowCODCertifiedFunds As Boolean
Private OrderAllowCreditCardOnline As Boolean
Private OrderAllowCreditCardByPhone As Boolean
Private OrderAllowCreditCardByFax As Boolean
Private OrderAllowPersonalChecks As Boolean
Private OrderAllowCompanyChecks As Boolean
Private OrderAllowCredit As Boolean
Private OrderAllowPayPal As Boolean
'
' ----- Order items internal storage
'
Private Type OrderDetailsType
    Id As Long
    itemId As Long
    ItemCID As Long
    itemName As String
    quantity As Long
    Price As Double
    Weight As Double
    Active As Boolean
    ItemNumber As String
    shippingRequired As Boolean
    Options As String
    receiptCopy As String
End Type
'
' Order item array, holds order details during page processing
'
Private OrderDetails(50) As OrderDetailsType
'
'========================================================================
' ----- Account Variables
'========================================================================
'
Private AccountSourceForm As Long           ' Form that submitted the button/action
Private AccountForm As Long             ' Next form to be displayed
'
Private AccountButton As String           ' The button hit to submit the previous form
'
Private AccountAction As Long           ' The account action required
Private AccountLoginUsername As String
Private AccountLoginPassword As String
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("CommerceClass", MethodName, Description, ErrorNumber)
End Sub
'
'========================================================================
'   Init()
'========================================================================
'
Public Sub Init(MainPointer As Object)
    On Error GoTo ErrorTrap
    '
    Dim MethodName As String
    Dim OrderFormForce As Long
    Dim cs As Long
    Dim Pricing As New CmrcPricingClass
    '
    MethodName = "Commerce.Init()"
    '
    If Not ClassInitialized Then
        Set main = MainPointer
        '
        CommerceAccountMemberID = main.memberId
        AllowGroupPolicies = main.IsAuthenticated
        '
        ' orderBillAddress2 and orderShipAddress2 were created together in 4.1.513
        '
        isSupportedAddress2 = main.IsContentFieldSupported("people", "BillAddress2")
        '
        Call Pricing.GetRightsPolicies(main, main.memberId, main.IsAuthenticated, IsMemberCommerceAdmin)
        If IsMemberCommerceAdmin Then
            '
            ' A commerce Admin is placing order, pick up the person they are ordering for
            '
            CommerceAccountMemberID = main.GetVisitProperty("CommerceAccountMemberID", main.memberId)
            If CommerceAccountMemberID = 0 Then
                CommerceAccountMemberID = main.memberId
            End If
        End If
        '
        ' ----- Initialize
        '
        Call LoadProperties
        orderId = 0
        main.OrderItemCount = 0
        CommerceDataSourceName = main.GetContentDataSource("Orders")
        ClassInitialized = True
        End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext(MethodName): Resume Next
    End Sub
'
'=============================================================================
' Process Forms for Order System
'   Returns the Form to be displayed next
'=============================================================================
'
Private Function processCheckoutForm(srcShopFormId As Long) As Long
    On Error GoTo ErrorTrap
    '
    Dim dstShopFormId As Long
    Dim emailDup As Boolean
    Dim firstName As String
    Dim lastName As String
    Dim common As New commonClass
    Dim OrderCCExpiresDate As Date
    Dim OrderCCExpiresMonth As Long
    Dim OrderCCExpiresYear As Long
    Dim CompletedOrdersCID As Long
    Dim OrderProcessedOK As Boolean
    Dim copy As String
    Dim CSShipMethods As Long
    Dim ItemCount As Long
    Dim TempMessage As String
    Dim RedirectPage As String
    Dim ErrorMessage As String
    Dim ErrorCode As Long
    Dim CatalogAllowSpecialPrice As Boolean
    Dim quantity As Long
    Dim ItemNumber As Long
    Dim cs As Long
    'Dim ProcPay As New ProcPay.ProcessClass3
    Dim Comment1 As String
    Dim Comment2 As String
    Dim transactionString As String
    Dim Link As String
    Dim OrderFormForce As Long
    Dim ignore_UpdateMemberRecord As Boolean

    Dim optionString As String
    Dim resultString As String
    Dim Approved As Boolean
    Dim Response As String
    Dim Reference As String
    Dim doc As New DOMDocument60
    Dim csCust As Long
    Dim subscriptionFlag As Boolean
    Dim CartRows As Long
    Dim RowID As Long
    Dim RowQuantity As Long
    Dim RowDetailID As Long
    Dim RowPtr As Long
    Dim ItemPtr As Long
    Dim SQL As String
    Dim accountEmail As String
    '
    '----------------------------------------------------------------------
    ' ----- read in process variables
    '----------------------------------------------------------------------
    '
    subscriptionFlag = isSubscriptionOrder()
    shopButton = main.GetStreamText(rnButton)
    'OrderForm = main.GetStreamNumber(rnSrcShopFormId)
    OrderMemberUsername = main.GetStreamText("oau")
    OrderMemberPassword = main.GetStreamText("oap")
    OrderMemberEmail = main.GetStreamText("oae")
    OrderFormForce = main.GetStreamInteger("off")
    '
    If OrderFormForce <> 0 Then
        srcShopFormId = OrderFormForce
    End If
    dstShopFormId = srcShopFormId
    '
    '----------------------------------------------------------------------
    ' check required conditions
    '----------------------------------------------------------------------
    '
    'Should already be loaded from init - Call LoadOrderGlobals
    If Not main.IsOrderOK() Then
        '
        ' ----- There is not current, non-complete order
        '
        If (srcShopFormId = shopFormCheckoutReview) And ((shopButton = shopButtonBack) Or (shopButton = shopButtonBackToShopping)) Then
            '
            ' ----- BackButton on Review form state
            '       the only action without IsOrderOK()
            '
            Call main.SetVisitProperty("OrderForm", srcShopFormId)
            Call main.Redirect(GetCurrentLink(main.ServerPageSecure))
        End If
        'Exit Function
    End If
    '
    '----------------------------------------------------------------------
    ' ----- Process Source Forms / Buttons
    '----------------------------------------------------------------------
    '
    If (srcShopFormId > 0) Then
    ' no not require shopButton because the login form uses contensive button anyway
    'If (shopButton <> "") And (srcShopFormId > 0) Then
        Select Case srcShopFormId
            Case shopFormCheckoutAccount
                '
                ' ----- process the account page (login or create account)
                '       if not in these checks, go back to account login form
                '
                dstShopFormId = processCommonForm_Login(shopFormCheckoutShippingBilling, shopFormCheckoutAccount)
            Case shopFormCheckoutReview
                '
                ' ----- Process review form
                '
                CatalogAllowSpecialPrice = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowSpecialPrice", False))
                '
                CartRows = main.GetStreamInteger("CartRows")
                If CartRows > 0 And orderId = 0 Then
                    orderId = AddOrderGetID()
                End If
                ItemPtr = 0
                For RowPtr = 0 To CartRows - 1
                    RowQuantity = main.GetStreamInteger("Q" & CStr(RowPtr))
                    RowID = main.GetStreamInteger("ID" & CStr(RowPtr))
                    RowDetailID = main.GetStreamInteger("D" & CStr(RowPtr))
                    If RowDetailID <> 0 And RowQuantity = 0 Then
                        '
                        ' delete detail
                        '
                        Call main.DeleteContentRecord("Order Details", RowDetailID)
                    ElseIf RowDetailID <> 0 Then
                        '
                        ' update quantity
                        '
                        Call main.ExecuteSQL("default", "update OrderDetails set Quantity=" & RowQuantity & " where ID=" & RowDetailID)
                    ElseIf (RowID > 0) And (RowQuantity > 0) Then
                        '
                        ' insert new detail
                        '
                        Call AddItem(RowID)
                        'CS = Main.InsertCSRecord("Order Details")
                        'If Main.IsCSOK(CS) Then
                        '    Call Main.SetCS(CS, "orderID", OrderID)
                        '    Call Main.SetCS(CS, "ItemID", RowID)
                        '    Call Main.SetCS(CS, "Quantity", RowQuantity)
                        'End If
                        'Call Main.CloseCS(CS)
                    End If
                Next
                Call LoadOrderGlobals
                '
                Select Case shopButton
                    Case shopButtonContinue
                        '
                        ' ----- Continue Button
                        '
                        Call CalculateAndSaveOrder
                        'If Not main.IsAuthenticated Then
                        '    dstShopFormId = shopFormCheckoutAccount
                        'Else
                            dstShopFormId = shopFormCheckoutShippingBilling
                        'End If
                    Case shopButtonBack, shopButtonBackToShopping
                        '
                        ' ----- Back Button, return to site
                        '
                        Call CalculateAndSaveOrder
                        Call main.Redirect(GetCurrentLink(False))
                        Exit Function
                    Case shopButtonRecalculate
                        '
                        ' ----- Recalculate Button
                        '
                        Call CalculateAndSaveOrder
                        dstShopFormId = shopFormCheckoutReview
                    Case shopButtonRemoveAll
                        '
                        ' ----- Remove All Button
                        '
                        If main.OrderItemCount > 0 Then
                            For ItemCount = 0 To main.OrderItemCount - 1
                                OrderDetails(ItemCount).quantity = 0
                                Next
                            End If
                        Call CalculateAndSaveOrder
                        main.VisitorOrderID = 0
                        Call main.SaveVisitor
                        Call main.Redirect(GetCurrentLink(False))
                    Case Else
                        '
                        ' ----- unknown Button
                        '
                        Call HandleClassErrorAndResumeNext("OrderInit", "Unexpected Order Button = [" & shopButton & "]")
                    End Select
            Case shopFormCheckoutShippingBilling
                '
                ' ----- process the shipping/billing form
                '
                dstShopFormId = processCheckout_ShippingBilling(shopButton, ignore_UpdateMemberRecord)
            Case shopFormCheckoutPurchase
                '
                ' ----- Process purchase form
                '
                dstShopFormId = shopFormCheckoutPurchase
                If subscriptionFlag And (Not main.IsAuthenticated()) Then
                    '
                    ' this could happen if they waited too long to order
                    '
                    Call main.AddUserError("Your order includes a subscription, which requires you be logged in. There may have been an error, or your visit to the site may have timed out. Please log in and try again")
                    dstShopFormId = shopFormCheckoutReview
                Else
                    '
                    '
                    '
                    Select Case shopButton
                        Case shopButtonBack
                            '
                            ' Back button, save the order and go back to the last form
                            '
                            Call CalculateAndSaveOrder
                            dstShopFormId = shopFormCheckoutShippingBilling
                        Case shopButtonOrderNow
                            '
                            ' Setup for completion, actual completion of order is done in the PrintOrderForms
                            '
                            If OrderPayMethod <> PayString(PayTypeCreditCardOnline) Then
                                '
                                ' non-Credit card transaction, authorize and complete the order
                                '
                                If main.GetUserError = "" Then
                                    '
                                    ' Complete the order
                                    '
                                    OrderCompleted = True
                                    OrderDateCompleted = Now()
                                    dstShopFormId = shopFormCheckoutReceipt
                                    '
                                    'Set common = New commonClass
                                    Call common.completeOrder(main, "", "Pending Authorization", "", OrderCCType, OrderCCNumber, OrderCCExpires, OrderCCName, OrderCCZip, CommerceAccountMemberID, orderId)
                                    dstShopFormId = shopFormCheckoutReceipt
                                    main.OrderAuthorize = False
                                    main.OrderAuthorized = True
                                End If
                            Else
                                '
                                ' ----- processCheckout needs fund authorization
                                '
                                OrderCCExpiresMonth = main.GetStreamInteger("OrderCCExpiresMonth")
                                OrderCCExpiresYear = main.GetStreamInteger("OrderCCExpiresYear")
                                OrderCCType = main.GetStreamText("OrderCCType")
                                OrderCCNumber = main.GetStreamText("OrderCCNumber")
                                OrderCCExpires = main.GetStreamText("OrderCCExpiresMonth") & "/" & main.GetStreamText("OrderCCExpiresYear")
                                OrderCCName = main.GetStreamText("OrderCCName")
                                OrderCCZip = main.GetStreamText("OrderCCZip")
                                If False Then
                                    ' this case should be covered in procPay
                                    'If (OrderCCNumber <> "") And (OrderCCNumber = main.GetSiteProperty("OrderTestCCNumber", "", True)) Then
                                    '
                                    ' Test Card number
                                    ' Complete the order
                                    '
                                    OrderCompleted = True
                                    OrderDateCompleted = Now()
                                    dstShopFormId = shopFormCheckoutReceipt
                                    '
                                    'Set common = New commonClass
                                    Call common.completeOrder(main, "RefCode - No Purchase, Test Credit Card", "ActionCode - No Purchase, Test Credit Card", "AuthCode - No Purchase, Test Credit Card", OrderCCType, OrderCCNumber, OrderCCExpires, OrderCCName, OrderCCZip, CommerceAccountMemberID, orderId)
                                    dstShopFormId = shopFormCheckoutReceipt
                                Else
                                    If ("" = OrderCCNumber) Then
                                        main.AddUserError "A credit card number is required."
                                    End If
                                    If ("" = OrderCCExpires) Then
                                        main.AddUserError "A credit card expiration date is required."
                                    Else
                                        OrderCCExpiresMonth = OrderCCExpiresMonth + 1
                                        If OrderCCExpiresMonth > 12 Then
                                            OrderCCExpiresMonth = 1
                                            OrderCCExpiresYear = OrderCCExpiresYear + 1
                                        End If
                                        OrderCCExpiresDate = CDate(OrderCCExpiresMonth & "/1/" & OrderCCExpiresYear)
                                        If OrderCCExpiresDate < Now() Then
                                            main.AddUserError "The credit card appears to be expired. Please check the expiration date."
                                        End If
                                    End If
                                    If ("" = OrderCCName) Then
                                        main.AddUserError "A credit card billing name is required."
                                    End If
                                    If ("" = OrderCCZip) Then
                                        main.AddUserError "A credit card billing postal code is required."
                                    End If
                                    If Not main.IsUserError Then
                                        '
                                        ' If no errors, process order payment
                                        '
                                        Comment1 = main.ApplicationName & " Shopping Cart Order " & orderId
                                        '
                                        optionString = ""
                                        optionString = optionString & "&CreditCardNumber=" & encodeNvaArgument(OrderCCNumber)
                                        optionString = optionString & "&CreditCardExpiration=" & encodeNvaArgument(OrderCCExpires)
                                        optionString = optionString & "&PaymentAmount=" & encodeNvaArgument(CStr(OrdertotalCharge))
                                        optionString = optionString & "&Comment1=" & encodeNvaArgument(CStr(Comment1))
                                        optionString = optionString & "&Comment2=" & encodeNvaArgument(CStr(Comment2))
                                        optionString = optionString & "&TransactionString=" & encodeNvaArgument(CStr(transactionString))
                                        optionString = optionString & "&transDesc=" & encodeNvaArgument(Comment1)
    
                                        '
                                        '   customer information
                                        '
    Dim addressLine As String
    Dim addressLine2 As String
    Dim csOrder As Long
    Dim billName As String
    Dim billFirstName As String
    Dim billLastName As String
    Dim pos As Long
                                        
                                        csOrder = main.OpenCSContentRecord("Orders", orderId)
                                        If main.IsCSOK(csOrder) Then
                                            '
                                            ' special credit card fields (merge will billing later)
                                            '
                                            billName = main.GetStreamText("OrdercCName")
                                            billFirstName = ""
                                            billLastName = ""
                                            If billName = "" Then
                                                billFirstName = main.GetCSText(csOrder, "FirstName")
                                                billLastName = main.GetCSText(csOrder, "LastName")
                                            Else
                                                billFirstName = billName
                                                pos = InStr(1, billFirstName, " ")
                                                If (pos > 0) And (pos < Len(billFirstName)) Then
                                                    billLastName = Mid(billFirstName, pos + 1)
                                                    billFirstName = Mid(billFirstName, 1, pos - 1)
                                                End If
                                            End If
                                            optionString = optionString & "&custFirstName=" & encodeNvaArgument(billFirstName)
                                            optionString = optionString & "&custLastName=" & encodeNvaArgument(billLastName)
                                            optionString = optionString & "&custZip=" & encodeNvaArgument(main.GetStreamText("OrdercCzip"))
                                            '
                                            ' billing info
                                            '
                                            addressLine = main.GetCSText(csOrder, "BillAddress")
                                            If isSupportedAddress2 Then
                                                addressLine2 = main.GetCSText(csOrder, "BillAddress2")
                                                If (addressLine2 <> "") And (addressLine <> "") Then
                                                    addressLine = addressLine & ", "
                                                End If
                                                addressLine = addressLine & addressLine2
                                            End If
                                            optionString = optionString & "&custAddress=" & encodeNvaArgument(addressLine)
                                            optionString = optionString & "&custCity=" & encodeNvaArgument(main.GetCSText(csOrder, "BillCity"))
                                            optionString = optionString & "&custState=" & encodeNvaArgument(main.GetCSText(csOrder, "BillState"))
                                            'optionString = optionString & "&custZip=" & encodeNvaArgument(main.GetStreamText("OrdercCzip"))
                                            optionString = optionString & "&custCountry=" & encodeNvaArgument(main.GetCSText(csOrder, "BillCountry"))
                                            optionString = optionString & "&custEmail=" & encodeNvaArgument(main.GetCSText(csOrder, "BillEmail"))
                                            '
                                            ' shipping info
                                            '
                                            pos = InStr(1, firstName, " ")
                                            firstName = Trim(main.GetCSText(csOrder, "shipname"))
                                            If (pos <> 0) Then
                                                lastName = Mid(firstName, pos + 1)
                                                firstName = Mid(firstName, 1, pos)
                                            End If
                                            addressLine = main.GetCSText(csOrder, "shipAddress")
                                            If isSupportedAddress2 Then
                                                addressLine2 = main.GetCSText(csOrder, "shipAddress2")
                                                If (addressLine2 <> "") And (addressLine <> "") Then
                                                    addressLine = addressLine & ", "
                                                End If
                                                addressLine = addressLine & addressLine2
                                            End If
                                            optionString = optionString & "&shipFirstName=" & encodeNvaArgument(firstName)
                                            optionString = optionString & "&shipLastName=" & encodeNvaArgument(lastName)
                                            optionString = optionString & "&shipCompany=" & encodeNvaArgument(main.GetCSText(csOrder, "shipCompany"))
                                            optionString = optionString & "&shipAddress=" & encodeNvaArgument(addressLine)
                                            optionString = optionString & "&shipCity=" & encodeNvaArgument(main.GetCSText(csOrder, "shipCity"))
                                            optionString = optionString & "&shipState=" & encodeNvaArgument(main.GetCSText(csOrder, "shipState"))
                                            optionString = optionString & "&shipZip=" & encodeNvaArgument(main.GetCSText(csOrder, "shipZip"))
                                            optionString = optionString & "&shipCountry=" & encodeNvaArgument(main.GetCSText(csOrder, "shipCountry"))
                                        End If
                                        Call main.CloseCS(csOrder)
                                        
                                        
'                                        csCust = main.OpenCSContent("People", "ID=" & CommerceAccountMemberID)
'                                        If main.CSOK(csCust) Then
'                                            addressLine = main.GetCSText(csCust, "BillAddress")
'                                            If isSupportedAddress2 Then
'                                                addressLine2 = main.GetCSText(csCust, "BillAddress2")
'                                                If (addressLine2 <> "") And (addressLine <> "") Then
'                                                    addressLine = addressLine & ", "
'                                                End If
'                                                addressLine = addressLine & addressLine2
'                                            End If
'                                            optionString = optionString & "&custFirstName=" & encodeNvaArgument(main.GetCSText(csCust, "FirstName "))
'                                            optionString = optionString & "&custLastName=" & encodeNvaArgument(main.GetCSText(csCust, "LastName"))
'                                            optionString = optionString & "&custAddress=" & encodeNvaArgument(addressLine)
'                                            optionString = optionString & "&custCity=" & encodeNvaArgument(main.GetCSText(csCust, "BillCity"))
'                                            optionString = optionString & "&custState=" & encodeNvaArgument(main.GetCSText(csCust, "BillState"))
'                                            optionString = optionString & "&custZip=" & encodeNvaArgument(main.GetCSText(csCust, "BillZip"))
'                                            optionString = optionString & "&custCountry=" & encodeNvaArgument(main.GetCSText(csCust, "BillCountry"))
'                                            optionString = optionString & "&custEmail=" & encodeNvaArgument(main.GetCSText(csCust, "BillEmail"))
'                                            '
'                                            firstName = Trim(main.GetCSText(csCust, "shipname"))
'                                            If (InStr(1, firstName, " ") <> 0) Then
'                                                lastName = Mid(firstName, InStr(1, firstName, " ") + 1)
'                                                firstName = Mid(firstName, 1, InStr(1, firstName, " "))
'                                            End If
'                                            addressLine = main.GetCSText(csCust, "shipAddress")
'                                            If isSupportedAddress2 Then
'                                                addressLine2 = main.GetCSText(csCust, "shipAddress2")
'                                                If (addressLine2 <> "") And (addressLine <> "") Then
'                                                    addressLine = addressLine & ", "
'                                                End If
'                                                addressLine = addressLine & addressLine2
'                                            End If
'                                            optionString = optionString & "&shipFirstName=" & encodeNvaArgument(firstName)
'                                            optionString = optionString & "&shipLastName=" & encodeNvaArgument(lastName)
'                                            optionString = optionString & "&shipCompany=" & encodeNvaArgument(main.GetCSText(csCust, "shipCompany"))
'                                            optionString = optionString & "&shipAddress=" & encodeNvaArgument(addressLine)
'                                            optionString = optionString & "&shipCity=" & encodeNvaArgument(main.GetCSText(csCust, "shipCity"))
'                                            optionString = optionString & "&shipState=" & encodeNvaArgument(main.GetCSText(csCust, "shipState"))
'                                            optionString = optionString & "&shipZip=" & encodeNvaArgument(main.GetCSText(csCust, "shipZip"))
'                                            optionString = optionString & "&shipCountry=" & encodeNvaArgument(main.GetCSText(csCust, "shipCountry"))
'                                        End If
'                                        Call main.CloseCS(csCust)
    
                                        resultString = main.ExecuteAddon(0, "{F71E8C9B-38A4-446E-8CAC-07548EE602BB}", optionString, ContextAdmin, "", 0, "", "", -1)
                                        If resultString = "" Then
                                            main.AddUserError ("There was a problem processing your payment - please try again.")
                                        Else
                                            Call doc.loadXML(resultString)
                                            If doc.hasChildNodes Then
                                                Approved = kmaEncodeBoolean(doc.getElementsByTagName("status").Item(0).Text)
                                                Response = doc.getElementsByTagName("responseMessage").Item(0).Text
                                                Reference = doc.getElementsByTagName("referenceNumber").Item(0).Text
                                            End If
                                            If Not Approved Then
                                                '
                                                ' ----- Funds not authroized
                                                '
                                                cs = main.OpenCSContentRecord("Orders", orderId)
                                                If main.IsCSOK(cs) Then
                                                    Call main.SetCS(cs, "ccRefCode", Reference)
                                                    Call main.SetCS(cs, "CCActionCode", Response)
                                                    Call main.SetCS(cs, "CCAuthCode", Response)
                                                End If
                                                Call main.CloseCS(cs)
                                                If Response = "" Then
                                                    main.AddUserError ("This order could not be completed because there was a problem authorizing the transaction.")
                                                Else
                                                    main.AddUserError ("This order could not be completed because there was a problem authorizing the transaction. The credit card processor returned the following message (" & Response & ").")
                                                End If
                                            Else
                                                '
                                                ' ----- Funds authorized
                                                ' Complete the order
                                                '
                                                OrderCompleted = True
                                                OrderDateCompleted = Now()
                                                dstShopFormId = shopFormCheckoutReceipt
                                                '
                                                'Set common = New commonClass
                                                Call common.completeOrder(main, Reference, Response, Reference, OrderCCType, OrderCCNumber, OrderCCExpires, OrderCCName, OrderCCZip, CommerceAccountMemberID, orderId)
                                                dstShopFormId = shopFormCheckoutReceipt
                                            End If
                                        End If
                                    
    '
    '   3/31/2010 - JF added call to ProcPay as addon
    '
    '                                    '
    '                                    ' If no errors, process order
    '                                    '
    '                                    Call ProcPay.Init(Main)
    '                                    OrderProcessedOK = ProcPay.PaymentProcessed(OrderCCNumber, OrderCCExpires, CStr(OrdertotalCharge), , Comment1, Comment2, transactionString)
    '                                    If Not OrderProcessedOK Then
    '                                        '
    '                                        ' ----- Funds not authroized
    '                                        '
    '                                        CS = Main.OpenCSContentRecord("Orders", OrderID)
    '                                        If Main.IsCSOK(CS) Then
    '                                            Call Main.SetCS(CS, "ccRefCode", ProcPay.ReferenceNumber)
    '                                            Call Main.SetCS(CS, "CCActionCode", ProcPay.ResponseMessage)
    '                                            Call Main.SetCS(CS, "CCAuthCode", ProcPay.ResponseMessage)
    '                                        End If
    '                                        Call Main.CloseCS(CS)
    '                                        '
    '                                        '
    '                                        '
    '                                        If ProcPay.ResponseMessage = "" Then
    '                                            Main.AddUserError ("This order could not be completed because there was a problem authorizing the transaction.")
    '                                        Else
    '                                            Main.AddUserError ("This order could not be completed because there was a problem authorizing the transaction. The credit card processor returned the following message (" & ProcPay.ResponseMessage & ").")
    '                                        End If
    '                                    Else
    '                                        '
    '                                        ' ----- Funds authorized
    '                                        '
    '                                        Call completeOrder(ProcPay.ReferenceNumber, ProcPay.ResponseMessage, ProcPay.ReferenceNumber, OrderCCType, OrderCCNumber, OrderCCExpires, OrderCCName, OrderCCZip)
    '                                        dstShopFormId = shopFormCheckoutReceipt
    '                                    End If
                                    End If
                                End If
                            End If
                        Case Else
                            Call HandleClassErrorAndResumeNext("OrderInit", "unexpected order button = [" & shopButton & "]")
                        End Select
                
                
                
                
                
                End If
            End Select
        End If
    '
    ' Save the Next Order Form in a visit property, so it can be read back in PrintOrderForms
    '
    Call main.SetVisitProperty("OrderForm", dstShopFormId)
    processCheckoutForm = dstShopFormId
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("processCheckout"): Resume Next
    End Function
'
'=============================================================================
' Process Forms for Order System
'   Returns the Form to be displayed next
'=============================================================================
'
Private Function processCheckout_ShippingBilling(shopButton As String, ignore_UpdateMemberRecord As Boolean) As Long
    On Error GoTo ErrorTrap
    '
    Dim userId As Long
    Dim userUsername As String
    Dim userPassword As String
    Dim csCheck As Long
    'Dim accountPassword As String
    Dim checkCreateKey As Long
    Dim SQL As String
    Dim dstShopFormId As Long
    Dim pos As Long
    Dim firstName As String
    Dim lastName As String
    Dim allowSaveCredentials As Boolean
    Dim AccountUsernameErrorMessaage As String
    Dim AccountUsernameErrorCode As Long
    Dim CompletedOrdersCID As Long
    Dim OrderProcessedOK As Boolean
    Dim copy As String
    Dim CSShipMethods As Long
    Dim ItemCount As Long
    Dim TempMessage As String
    Dim RedirectPage As String
    Dim ErrorMessage As String
    Dim ErrorCode As Long
    Dim CatalogAllowSpecialPrice As Boolean
    Dim quantity As Long
    Dim ItemNumber As Long
    Dim cs As Long
    'Dim ProcPay As New ProcPay.ProcessClass3
    Dim Comment1 As String
    Dim Comment2 As String
    Dim transactionString As String
    Dim Link As String
    Dim OrderFormForce As Long
    Dim csDup As Long
    Dim subscriptionFlag As Boolean
    Dim blockContactSave As Boolean
    Dim blockContactGroup As String
    Dim common As New commonClass
    '
    blockContactSave = False
    If main.IsAuthenticated Then
        blockContactGroup = Trim(main.GetSiteProperty("ShoppingCart Block Contact Update Group", ""))
        If blockContactGroup <> "" Then
            blockContactSave = common.isGroupMember(main, blockContactGroup, CommerceAccountMemberID)
        End If
    End If
    '
    dstShopFormId = shopFormCheckoutShippingBilling
    '
    Select Case shopButton
        Case shopButtonContinue
            '
            ' get contact information
            '
            subscriptionFlag = isSubscriptionOrder()
            orderShipMethod = main.GetStreamText("orderShipMethod")
            '
            If Not blockContactSave Then
                AccountUsername = main.GetStreamText("AccountUsername")
                accountPassword = main.GetStreamText("AccountPassword")
                '
                orderContactName = main.GetStreamText("orderContactName")
                orderContactCompany = main.GetStreamText("orderContactCompany")
                orderContactEmail = main.GetStreamText("orderContactEmail")
                orderContactAddress = main.GetStreamText("orderContactAddress")
                If isSupportedAddress2 Then
                    orderContactAddress2 = main.GetStreamText("orderContactAddress2")
                End If
                orderContactCity = main.GetStreamText("orderContactCity")
                orderContactCountry = main.GetStreamText("orderContactCountry")
                orderContactState = main.GetStreamText("orderContactState")
                orderContactZip = main.GetStreamText("orderContactZip")
                orderContactPhone = main.GetStreamText("orderContactPhone")
                orderContactFax = main.GetStreamText("orderContactFax")
                '
                ' (you must be authenticated to get here)
                ' - check for valid email and username
                '
                If (orderContactEmail = "") Then
                    main.AddUserError "Please enter a valid contact email address."
                Else
                    If (main.MemberEmail = orderContactEmail) Then
                        '
                        ' email was not changed, do not check for dups
                        '
                    Else
                        '
                        ' email changed - check for dup
                        '
                        SQL = "select top 1 id from ccmembers where (email=" & main.EncodeSQLText(orderContactEmail) & ")and(id<>" & main.memberId & ")"
                        cs = main.OpenCSSQL("", SQL)
                        If main.IsCSOK(cs) Then
                            main.AddUserError "The email address you entered is in use by another account. Please enter a different email."
                        End If
                        Call main.CloseCS(cs)
                    End If
                End If
                '
                ' Check for required billing values
                '
                If LCase(orderContactCountry) = "united states" Then
                    If False _
                        Or ("" = orderContactName) _
                        Or ("" = orderContactAddress) _
                        Or ("" = orderContactCity) _
                        Or ("" = orderContactState) _
                        Or ("" = orderContactZip) _
                        Or ("" = orderContactPhone) Then
                        main.AddUserError "One or more of the required contact fields were blank."
                    End If
                Else
                    If False _
                        Or ("" = orderContactName) _
                        Or ("" = orderContactAddress) _
                        Or ("" = orderContactCity) _
                        Or ("" = orderContactZip) _
                        Or ("" = orderContactPhone) Then
                        main.AddUserError "One or more of the required contact fields were blank."
                    End If
                End If
            End If
            '
            If main.GetStreamBoolean("billToContactAddress") Then
                orderBillName = orderContactName
                orderBillEmail = orderContactEmail
                orderBillCompany = orderContactCompany
                orderBillAddress = orderContactAddress
                If isSupportedAddress2 Then
                    orderBillAddress2 = orderContactAddress2
                End If
                orderBillCity = orderContactCity
                orderBillCountry = orderContactCountry
                orderBillState = orderContactState
                orderBillZip = orderContactZip
                orderBillPhone = orderContactPhone
                orderBillFax = orderContactFax
            Else
                orderBillName = main.GetStreamText("orderBillName")
                orderBillEmail = main.GetStreamText("orderBillEmail")
                orderBillCompany = main.GetStreamText("orderBillCompany")
                orderBillAddress = main.GetStreamText("orderBillAddress")
                If isSupportedAddress2 Then
                    orderBillAddress2 = main.GetStreamText("orderBillAddress2")
                End If
                orderBillCity = main.GetStreamText("orderBillCity")
                orderBillCountry = main.GetStreamText("orderBillCountry")
                orderBillState = main.GetStreamText("orderBillState")
                orderBillZip = main.GetStreamText("orderBillZip")
                orderBillPhone = main.GetStreamText("orderBillPhone")
                orderBillFax = main.GetStreamText("orderBillFax")
                '
                ' Check for required billing values
                '
                If LCase(orderBillCountry) = "united states" Then
                    If False _
                        Or ("" = orderBillName) _
                        Or ("" = orderBillAddress) _
                        Or ("" = orderBillCity) _
                        Or ("" = orderBillState) _
                        Or ("" = orderBillZip) _
                        Or ("" = orderBillPhone) Then
                        main.AddUserError "One or more of the required billing fields were blank."
                    End If
                Else
                    If False _
                        Or ("" = orderBillName) _
                        Or ("" = orderBillAddress) _
                        Or ("" = orderBillCity) _
                        Or ("" = orderBillZip) _
                        Or ("" = orderBillPhone) Then
                        main.AddUserError "One or more of the required billing fields were blank."
                    End If
                End If
            End If
            '
            If Not orderShippingRequired Then
                '
                ' shipping not required - leave default values in shipping settings
                '
            Else
                '
                ' shipping is required
                '
                If main.GetStreamBoolean("ShipToContactAddress") Then
                    orderShipName = orderContactName
                    orderShipCompany = orderContactCompany
                    orderShipAddress = orderContactAddress
                    If isSupportedAddress2 Then
                        orderShipAddress2 = orderContactAddress2
                    End If
                    orderShipCity = orderContactCity
                    orderShipCountry = orderContactCountry
                    orderShipState = orderContactState
                    orderShipZip = orderContactZip
                    orderShipPhone = orderContactPhone
                Else
                    orderShipName = main.GetStreamText("orderShipName")
                    orderShipCompany = main.GetStreamText("orderShipCompany")
                    orderShipAddress = main.GetStreamText("orderShipAddress")
                    If isSupportedAddress2 Then
                        orderShipAddress2 = main.GetStreamText("orderShipAddress2")
                    End If
                    orderShipCity = main.GetStreamText("orderShipCity")
                    orderShipCountry = main.GetStreamText("orderShipCountry")
                    orderShipState = main.GetStreamText("orderShipState")
                    orderShipZip = main.GetStreamText("orderShipZip")
                    orderShipPhone = main.GetStreamText("orderShipPhone")
                    '
                    ' check for required shiping values
                    '
                    If LCase(orderShipCountry) = "united states" Then
                        If ("" = orderShipName) _
                            Or ("" = orderShipAddress) _
                            Or ("" = orderShipState) _
                            Or ("" = orderShipCity) _
                            Or ("" = orderShipZip) _
                            Or ("" = orderShipPhone) Then
                            main.AddUserError "One or more of the required shipping fields were blank."
                        End If
                    Else
                        If ("" = orderShipName) _
                            Or ("" = orderShipAddress) _
                            Or ("" = orderShipCity) _
                            Or ("" = orderShipZip) _
                            Or ("" = orderShipPhone) Then
                            main.AddUserError "One or more of the required shipping fields were blank."
                        End If
                    End If
                End If
                '
                ' ----- check shipping method
                '
                If orderShipMethod = "" Then
                    '
                    ' shipping required but no shipping method selected
                    '
                    If Not isShippingConfigured() Then
                        '
                        ' not configured, set to FREE
                        '
                        orderShipMethod = "Free Shipping"
                    Else
                        '
                        ' they did not choose a shipping method
                        '
                        main.AddUserError "Please select a shipping method."
                    End If
                Else
                    '
                    ' shipping method selected -Check domestic/international shipping
                    '
                    CSShipMethods = main.OpenCSContent_Internal("Order Ship Methods", "(name<>'')and(name is not null)", , , , , "NAME,DOMESTIC,INTERNATIONAL")
                    If main.IsCSOK(CSShipMethods) Then
                        Do While main.IsCSOK(CSShipMethods)
                            If orderShipMethod = (main.GetCSText(CSShipMethods, "Name")) Then
                                Select Case LCase(orderShipCountry)
                                    Case "united states"
                                        If Not main.GetCSBoolean(CSShipMethods, "Domestic") Then
                                            main.AddUserError "You must select a domestic shipping method for delivery to " & orderShipCountry
                                        End If
                                    Case Else
                                        If Not main.GetCSBoolean(CSShipMethods, "International") Then
                                            main.AddUserError "You must select an International shipping method for delivery to " & orderShipCountry
                                        End If
                                End Select
                                Exit Do
                            End If
                            main.NextCSRecord (CSShipMethods)
                        Loop
                    End If
                    Call main.CloseCS(CSShipMethods)
                End If
            End If
            '
            '
            OrderPayMethod = main.GetStreamText("OrderPayMethod")
            '
            ' Check email address
            '
            IsOrderEmailBad = (Len(orderContactEmail) < 6) Or (InStr(1, orderContactEmail, ".", vbTextCompare) = 0) Or (InStr(1, orderContactEmail, "@", vbTextCompare) = 0)
            '
            ' Check for required billing values
            '
            If (OrderPayMethod = "") Then
                '
                ' required field is blank
                '
                main.AddUserError "A payment method has not been selected."
            End If
            '
            ' Check Credentials
            '   credentials are always required
            '   if allowEmailLogin then username is not there
            '
            allowSaveCredentials = False
            csCheck = main.OpenCSContent("people", "id=" & CommerceAccountMemberID)
            If main.IsCSOK(csCheck) Then
                userUsername = main.GetCSText(csCheck, "username")
                userPassword = main.GetCSText(csCheck, "password")
                checkCreateKey = main.GetCSInteger(csCheck, "createKey")
            End If
            Call main.CloseCS(csCheck)
            If main.visitId <> checkCreateKey Then
                '
                ' account not created here, do not offer login options
                '
            Else
                '
                ' login credentials required
                '
                allowSaveCredentials = True
                If accountPassword = "" Then
                    '
                    ' blank password, always required for first pass through
                    '
                    AccountPasswordError = True
                    main.AddUserError "Please enter a password."
                End If
                '
                If Not kmaEncodeBoolean(main.GetSiteProperty("ALLOWEMAILLOGIN", "")) Then
                    '
                    ' login by email is off, require username
                    '
                    If AccountUsername = "" Then
                        '
                        ' blank username
                        '
                        AccountUsernameError = True
                        main.AddUserError "Please enter a username."
                    End If
                    If Not main.IsUserError() Then
                        userId = main.memberId
                        If Not main.IsNewLoginOK(AccountUsername, accountPassword, AccountUsernameErrorMessaage) Then
                            '
                            ' bad username/password
                            '
                            AccountUsernameError = True
                            main.AddUserError AccountUsernameErrorMessaage
                        End If
                    End If
                End If
'                If main.IsAuthenticated Then
'                    '
'                    ' authenticated, skip username
'                    '
'                Else
'                    '
'                    ' not authenticated, but username is required
'                    '
'                    allowSaveCredentials = True
'                    If Not kmaEncodeBoolean(main.GetSiteProperty("ALLOWEMAILLOGIN", "")) Then
'                        '
'                        ' login by email is off, require username
'                        '
'                        If AccountUsername = "" Then
'                            '
'                            ' blank username
'                            '
'                            AccountUsernameError = True
'                            main.AddUserError "Please enter a username."
'                        ElseIf accountPassword = "" Then
'                            '
'                            ' blan password
'                            '
'                            AccountPasswordError = True
'                            main.AddUserError "Please enter a password."
'                        ElseIf Not main.IsNewLoginOK(AccountUsername, accountPassword, AccountUsernameErrorMessaage) Then
'                            '
'                            ' bad username/password
'                            '
'                            AccountUsernameError = True
'                            main.AddUserError AccountUsernameErrorMessaage
'                        End If
'                    End If
'                End If
            End If
'            If (subscriptionFlag) Then
'                '
'                ' username password required
'                '
'                If main.IsAuthenticated Then
'                    '
'                    ' authenticated, skip username
'                    '
'                Else
'                    '
'                    ' not authenticated, but username is required
'                    '
'                    allowSaveCredentials = True
'                    If Not kmaEncodeBoolean(main.GetSiteProperty("ALLOWEMAILLOGIN", "")) Then
'                        '
'                        ' login by email is off, require username
'                        '
'                        If AccountUsername = "" Then
'                            '
'                            ' blank username
'                            '
'                            AccountUsernameError = True
'                            main.AddUserError "Please enter a username."
'                        ElseIf AccountPassword = "" Then
'                            '
'                            ' blan password
'                            '
'                            AccountPasswordError = True
'                            main.AddUserError "Please enter a password."
'                        ElseIf Not main.IsNewLoginOK(AccountUsername, AccountPassword, AccountUsernameErrorMessaage) Then
'                            '
'                            ' bad username/password
'                            '
'                            AccountUsernameError = True
'                            main.AddUserError AccountUsernameErrorMessaage
'                        End If
'                    End If
'                End If
'            Else
'                '
'                ' username not required
'                '
'                If main.IsAuthenticated Then
'                    '
'                    ' authenticated, skip username save
'                    '
'                ElseIf (AccountUsername = "") Then
'                    '
'                    ' not authenticated, blank username, skip username save
'                    '
'                ElseIf (AccountPassword = "") Then
'                    '
'                    ' not authenticated, added username but no password
'                    '
'                    AccountUsernameError = True
'                    main.AddUserError "If you elect to add a username, your password can not be blank."
'                ElseIf Not main.IsNewLoginOK(AccountUsername, AccountPassword, AccountUsernameErrorMessaage) Then
'                    '
'                    ' not authenticated, bad username/password
'                    '
'                    AccountUsernameError = True
'                    main.AddUserError AccountUsernameErrorMessaage
'                Else
'                    '
'                    ' not authenticated, add username/password
'                    '
'                    allowSaveCredentials = True
'                End If
'            End If
            '
            If IsOrderEmailBad Then
                '
                ' Bad email address
                '
                main.AddUserError "Your email address does not appear to be valid."
            End If
            If Not main.IsUserError Then
                '
                ' update the people record
                '
                cs = main.OpenCSContentRecord("people", CommerceAccountMemberID)
                If main.IsCSOK(cs) Then
                    '
                    If (Not blockContactSave) Then
                        '
                        ' update contact
                        '
                        orderContactName = Trim(orderContactName)
                        If orderContactName <> "" Then
                            pos = InStr(1, orderContactName, " ")
                            If pos > 0 Then
                                firstName = Mid(orderContactName, 1, pos - 1)
                                lastName = Mid(orderContactName, pos + 1)
                                Call main.SetCS(cs, "firstName", firstName)
                                Call main.SetCS(cs, "lastName", lastName)
                            End If
                        End If
                        Call main.SetCS(cs, "Name", orderContactName)
                        Call main.SetCS(cs, "Email", orderContactEmail)
                        Call main.SetCS(cs, "Phone", orderContactPhone)
                        Call main.SetCS(cs, "Company", orderContactCompany)
                        Call main.SetCS(cs, "Address", orderContactAddress)
                        If isSupportedAddress2 Then
                            Call main.SetCS(cs, "Address2", orderContactAddress2)
                        End If
                        Call main.SetCS(cs, "City", orderContactCity)
                        Call main.SetCS(cs, "State", orderContactState)
                        Call main.SetCS(cs, "Zip", orderContactZip)
                        Call main.SetCS(cs, "Country", orderContactCountry)
                        If allowSaveCredentials Then
                            If Not main.IsAuthenticated() Then
                                Call main.LoginMemberByID(main.memberId)
                            End If
                            Call main.SetCS(cs, "Username", AccountUsername)
                            Call main.SetCS(cs, "Password", accountPassword)
                        End If
                    End If
                    '
                    Call main.SetCS(cs, "BillName", orderBillName)
                    Call main.SetCS(cs, "BillEmail", orderBillEmail)
                    Call main.SetCS(cs, "BillCompany", orderBillCompany)
                    Call main.SetCS(cs, "BillAddress", orderBillAddress)
                    If isSupportedAddress2 Then
                        Call main.SetCS(cs, "BillAddress2", orderBillAddress2)
                    End If
                    Call main.SetCS(cs, "BillCity", orderBillCity)
                    Call main.SetCS(cs, "BillState", orderBillState)
                    Call main.SetCS(cs, "BillZip", orderBillZip)
                    Call main.SetCS(cs, "BillCountry", orderBillCountry)
                    Call main.SetCS(cs, "BillPhone", orderBillPhone)
                    '
                    If orderShippingRequired Then
                        Call main.SetCS(cs, "ShipName", orderShipName)
                        Call main.SetCS(cs, "ShipCompany", orderShipCompany)
                        Call main.SetCS(cs, "ShipAddress", orderShipAddress)
                        If isSupportedAddress2 Then
                            Call main.SetCS(cs, "ShipAddress2", orderShipAddress2)
                        End If
                        Call main.SetCS(cs, "ShipCity", orderShipCity)
                        Call main.SetCS(cs, "ShipState", orderShipState)
                        Call main.SetCS(cs, "ShipZip", orderShipZip)
                        Call main.SetCS(cs, "ShipCountry", orderShipCountry)
                        Call main.SetCS(cs, "ShipPhone", orderShipPhone)
                    End If
                    '
                    ' if visitid = createkey, clear it now
                    '
                    If main.visitId = main.GetCSInteger(cs, "createkey") Then
                        Call main.SetCS(cs, "createkey", 0)
                    End If
                End If
                Call main.CloseCS(cs)
                '
                ' exit to next form
                '
                orderBillConfirmed = True
                orderShipConfirmed = True
                dstShopFormId = shopFormCheckoutPurchase
                Call CalculateAndSaveOrder
            End If
        Case shopButtonBack
            dstShopFormId = shopFormCheckoutReview
            Call CalculateAndSaveOrder
        Case Else
            'Call HandleClassErrorAndResumeNext("OrderInit", "unexpected order button = [" & shopButton & "]")
            ' end case
        End Select
    processCheckout_ShippingBilling = dstShopFormId
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("processCheckout_ShippingBilling"): Resume Next
    End Function
'
'=============================================================================
'   LoadOrderGlobals( OrderID )
'       Initializes the global order variables.
'       This does not need to be called from the page, as it is called
'       during initialization, and when an order is created
'   Results:
'       OrderID is 0 if order is not found
'   Mode:
'=============================================================================
'
Private Sub LoadOrderGlobals(Optional ForceLoad As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim OrderItemPointer As Long
    Dim isChangedMember As Boolean
    '
    ' Attempt to open order, if it fails, leave OrderID=0
    '
    If True Then
    'If (OrderID = 0) Or ForceLoad Then
        orderId = 0
        isChangedMember = False
        SQL = "select OrderID from ccvisitors where ID=" & main.VisitorID
        cs = main.OpenCSSQL("default", SQL)
        If main.IsCSOK(cs) Then
            orderId = main.GetCSInteger(cs, "OrderID")
        End If
        Call main.CloseCS(cs)
        '
        cs = main.OpenCSContentRecord("Orders", orderId)
        If main.IsCSOK(cs) Then
            OrderMemberID = (main.GetCSInteger(cs, "MemberID"))
            If (OrderMemberID <> CommerceAccountMemberID) Then
                OrderMemberID = CommerceAccountMemberID
                Call main.SetCSField(cs, "memberId", OrderMemberID)
                isChangedMember = True
            End If
            '
            orderShipMethod = (main.GetCSText(cs, "ShipMethod"))
            orderShipWeight = (main.GetCSNumber(cs, "ShipWeight"))
            OrderItemCharge = (main.GetCSNumber(cs, "ItemCharge"))
            orderShipCharge = (main.GetCSNumber(cs, "ShipCharge"))
            OrderTaxCharge = (main.GetCSNumber(cs, "TaxCharge"))
            OrdertotalCharge = (main.GetCSNumber(cs, "TotalCharge"))
            OrderPayMethod = (main.GetCSText(cs, "PayMethod"))
            OrderPayDate = (main.GetCSDate(cs, "PayDate"))
            OrderCCType = (main.GetCSText(cs, "CCType"))
            OrderCCName = (main.GetCSText(cs, "CCName"))
            OrderCCExpires = (main.GetCSText(cs, "CCExpires"))
            OrderCCZip = (main.GetCSText(cs, "CCZip"))
            OrderCompleted = (main.GetCSBoolean(cs, "OrderCompleted"))
            OrderDateCompleted = (main.GetCSDate(cs, "DateCompleted"))
            OrderSpecialPricing = (main.GetCSBoolean(cs, "SpecialPricing"))
            OrderActive = (main.GetCSBoolean(cs, "Active"))
            '
            orderBillName = (main.GetCSText(cs, "BillName"))
            orderBillEmail = (main.GetCSText(cs, "BillEmail"))
            orderBillPhone = (main.GetCSText(cs, "BillPhone"))
            orderBillFax = (main.GetCSText(cs, "BillFax"))
            orderBillCompany = (main.GetCSText(cs, "BillCompany"))
            orderBillAddress = (main.GetCSText(cs, "BillAddress"))
            If isSupportedAddress2 Then
                orderBillAddress2 = (main.GetCSText(cs, "BillAddress2"))
            End If
            orderBillCity = (main.GetCSText(cs, "BillCity"))
            orderBillState = (main.GetCSText(cs, "BillState"))
            orderBillZip = (main.GetCSText(cs, "BillZip"))
            orderBillCountry = (main.GetCSText(cs, "BillCountry"))
            orderBillConfirmed = (main.GetCSBoolean(cs, "BillConfirmed"))
            '
            orderShipDate = (main.GetCSDate(cs, "ShipDate"))
            orderShipName = (main.GetCSText(cs, "ShipName"))
            orderShipCompany = (main.GetCSText(cs, "ShipCompany"))
            orderShipAddress = (main.GetCSText(cs, "ShipAddress"))
            If isSupportedAddress2 Then
                orderShipAddress2 = (main.GetCSText(cs, "ShipAddress2"))
            End If
            orderShipCity = (main.GetCSText(cs, "ShipCity"))
            orderShipState = (main.GetCSText(cs, "ShipState"))
            orderShipZip = (main.GetCSText(cs, "ShipZip"))
            orderShipCountry = (main.GetCSText(cs, "ShipCountry"))
            orderShipPhone = (main.GetCSText(cs, "ShipPhone"))
            orderShipConfirmed = (main.GetCSBoolean(cs, "ShipConfirmed"))
            '
            If main.SiteProperty_BuildVersion >= "3.3.580" Then
                OrderPriceLevelDiscount = (main.GetCSNumber(cs, "PriceLevelDiscount"))
                OrderPriceLevelFreeShipping = (main.GetCSBoolean(cs, "PriceLevelFreeShipping"))
            End If
            '
            ' block credit card load
            '
            If Not kmaEncodeBoolean(main.GetSiteProperty("OrderSaveCreditCard", "0", True)) Then
                OrderCCNumber = ""
            Else
                OrderCCNumber = (main.GetCSText(cs, "CCNumber"))
            End If
        End If
        Call main.CloseCS(cs)
        '
        ' load contact address
        '
        cs = main.OpenCSContentRecord("people", CommerceAccountMemberID)
        If main.IsCSOK(cs) Then
            'OrderMemberID = CommerceAccountMemberID
            '
            orderContactName = (main.GetCSText(cs, "Name"))
            orderContactEmail = (main.GetCSText(cs, "Email"))
            orderContactPhone = (main.GetCSText(cs, "Phone"))
            orderContactFax = (main.GetCSText(cs, "Fax"))
            orderContactCompany = (main.GetCSText(cs, "Company"))
            orderContactAddress = (main.GetCSText(cs, "Address"))
            If isSupportedAddress2 Then
                orderContactAddress2 = (main.GetCSText(cs, "Address2"))
            End If
            orderContactCity = (main.GetCSText(cs, "City"))
            orderContactState = (main.GetCSText(cs, "State"))
            orderContactZip = (main.GetCSText(cs, "Zip"))
            orderContactCountry = (main.GetCSText(cs, "Country"))
            '
            If isChangedMember Then
                '
                ' Member identity has changed since the order was created, update Billing and Shipping
                '
                orderBillName = (main.GetCSText(cs, "BillName"))
                orderBillEmail = (main.GetCSText(cs, "BillEmail"))
                orderBillPhone = (main.GetCSText(cs, "BillPhone"))
                orderBillFax = (main.GetCSText(cs, "BillFax"))
                orderBillCompany = (main.GetCSText(cs, "BillCompany"))
                orderBillAddress = (main.GetCSText(cs, "BillAddress"))
                If isSupportedAddress2 Then
                    orderBillAddress2 = (main.GetCSText(cs, "BillAddress2"))
                End If
                orderBillCity = (main.GetCSText(cs, "BillCity"))
                orderBillState = (main.GetCSText(cs, "BillState"))
                orderBillZip = (main.GetCSText(cs, "BillZip"))
                orderBillCountry = (main.GetCSText(cs, "BillCountry"))
                orderBillConfirmed = False
                '
                orderShipName = (main.GetCSText(cs, "ShipName"))
                orderShipCompany = (main.GetCSText(cs, "ShipCompany"))
                orderShipAddress = (main.GetCSText(cs, "ShipAddress"))
                If isSupportedAddress2 Then
                    orderShipAddress2 = (main.GetCSText(cs, "ShipAddress2"))
                End If
                orderShipCity = (main.GetCSText(cs, "ShipCity"))
                orderShipState = (main.GetCSText(cs, "ShipState"))
                orderShipZip = (main.GetCSText(cs, "ShipZip"))
                orderShipCountry = (main.GetCSText(cs, "ShipCountry"))
                orderShipPhone = (main.GetCSText(cs, "ShipPhone"))
                orderShipConfirmed = False
                Call CalculateAndSaveOrder
            End If
        End If
        Call main.CloseCS(cs)
    End If
    '
    ' pick up order details
    '
    OrderItemPointer = 0
    If (orderId <> 0) Then
        SQL = "select D.ID,D.UnitPrice,I.ID as ItemID,I.ContentControlID as ItemCID,I.name,I.Weight,I.ItemNumber,I.ShippingRequired,D.Quantity,D.Options as Options,i.allowReceiptCopy,i.receiptCopy" _
            & " From OrderDetails D left join Items I on I.ID=D.ItemID" _
            & " where D.OrderID=" & orderId
        cs = main.OpenCSSQL("default", SQL)
        Do While main.IsCSOK(cs)
            With OrderDetails(OrderItemPointer)
                .Id = (main.GetCSInteger(cs, "ID"))
                .ItemCID = (main.GetCSInteger(cs, "ItemCID"))
                .itemId = (main.GetCSInteger(cs, "ItemID"))
                .itemName = (main.GetCSText(cs, "name"))
                .quantity = (main.GetCSInteger(cs, "quantity"))
                .Price = (main.GetCSNumber(cs, "unitprice"))
                .Weight = (main.GetCSNumber(cs, "weight"))
                .ItemNumber = (main.GetCSText(cs, "ItemNumber"))
                .shippingRequired = main.GetCSBoolean(cs, "ShippingRequired")
                .Options = (main.GetCSText(cs, "Options"))
                If main.GetCSBoolean(cs, "allowReceiptCopy") Then
                    .receiptCopy = main.GetCSText(cs, "receiptCopy")
                End If
                orderShippingRequired = orderShippingRequired Or .shippingRequired
                .Active = True
            End With
            OrderItemPointer = OrderItemPointer + 1
            Call main.NextCSRecord(cs)
        Loop
        Call main.CloseCS(cs)
    End If
    main.OrderItemCount = OrderItemPointer
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("LoadOrderGlobals"): Resume Next
    End Sub
'
'=============================================================================
' Add an item to the order
'   Must be public - used in MainClass
'
'   Note: you must call the init before calling this routine
'=============================================================================
'
Public Sub AddItem(itemId As Long)
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim csItems As Long
    Dim CSOrganizations As Long
    Dim SourceName As String
    Dim SourceItemNumber As Long
    Dim UnitPrice As Double
    Dim DateAdded As Date
    Dim UnitWeight As Double
    Dim ItemNumber As Long
    Dim OrganizationID As Long
    Dim ItemInCart As Boolean
    Dim ItemCID As Long
    Dim OrderDetailsID As Long
    Dim csOrder As Long
    Dim copy As String
    Dim CatalogAllowSpecialPrice As Boolean
    Dim csSource As Long
    Dim Pricing As New CmrcPricingClass
    Dim AllowGroupPolicies As Boolean
    Dim optionCnt As Long
    Dim OptionValue As String
    Dim OptionCaption As String
    Dim Options As String
    Dim Ptr As Long
    '
    ' ----- limit cart to the maximum number it can hold
    '
    Call LoadOrderGlobals
    If main.OrderItemCount >= 50 Then
        Call main.TestPoint("Commerce.AddItem, cart is full")
        main.AddUserError "Your shopping cart is full. Please complete your order and return to select more items."
        Exit Sub
    End If
    '
    ' ----- get table name
    '
    ItemCID = main.GetContentID("items")
    '
    If Not main.IsOrderOK() Then
        ''
        '' ----- First item, write the javascript that will open a popup help window
        ''
        'If Not IsMemberCommerceAdmin Then
        '    main.ClosePageHTML = main.ClosePageHTML & main.GetPopupPage(main.ServerPage & "?Method=NewOrderPage", 400, 372)
        'End If
        '
        ' ----- Create a new order in the database
        '
        orderId = AddOrderGetID()
        'Call Main.OpenMember(CommerceAccountMemberID)
    End If
    '
    ' Create the Options if required
    '
    optionCnt = main.GetStreamInteger("OptionCnt")
    Options = ""
    If optionCnt > 0 Then
        For Ptr = 0 To optionCnt - 1
            OptionValue = main.GetStreamText("Option_" & itemId & "_" & Ptr)
            OptionCaption = main.GetStreamText("OptionCaption" & Ptr)
            Options = Options & ", " & OptionCaption & ":" & OptionValue
        Next
        If Options <> "" Then
            Options = Mid(Options, 3)
        End If
    End If
    '
    ' ----- block addition if they already have one and add up ItemCharge
    '
    ItemInCart = False
    If main.OrderItemCount > 0 Then
        For ItemNumber = 0 To main.OrderItemCount - 1
            If (itemId = OrderDetails(ItemNumber).itemId) And (ItemCID = OrderDetails(ItemNumber).ItemCID) And (Options = OrderDetails(ItemNumber).Options) Then
                cs = main.OpenCSContentRecord("Order Details", OrderDetails(ItemNumber).Id)
                If main.IsCSOK(cs) Then
                    Call main.SetCS(cs, "Quantity", main.GetCSInteger(cs, "Quantity") + 1)
                End If
                Call main.CloseCS(cs)
                ItemInCart = True
                Exit For
            End If
        Next
    End If
    '
    ' ----- Add the item to the order
    '
    If Not ItemInCart Then
        csItems = main.OpenCSContentRecord("Items", itemId)
        If Not main.IsCSOK(csItems) Then
            '
            ' Item was not found in the Items Content Def, or any of its child def
            '
            Call main.AddUserError("There was a problem adding this item to your shopping cart. Please try again, or contact the site administrator for order the item directly.")
            Call HandleClassErrorAndResumeNext("AddItem", "The item# [" & itemId & "] was not found with the 'Items' Content Definition. Make sure the item record is valid, active, and included within the 'Items' Definition or its children.")
        Else
            '
            ' create new order detail record
            '
            cs = main.InsertCSContent("Order Details")
            If main.IsCSOK(cs) Then
                Call main.SetCS(cs, "OrderID", orderId)
                Call main.SetCS(cs, "ItemCID", ItemCID)
                Call main.SetCS(cs, "ItemID", itemId)
                Call main.SetCS(cs, "Quantity", 1)
                If main.SiteProperty_BuildVersion > "3.3.568" Then
                    Call main.SetCS(cs, "Options", Options)
                End If
            End If
            Call main.CloseCS(cs)
            '
            main.OrderItemCount = main.OrderItemCount + 1
        End If
        Call main.CloseCS(csItems)
    End If
    '
    ' load the values, calculate the new pricing and save it
    '
    Call LoadOrderGlobals
    Call CalculateAndSaveOrder
    Call LoadOrderGlobals
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("AddItem"): Resume Next
    End Sub
'
'==============================================================================
'   Calculate Order
'       Calculate the order details and order, set the orderdetail structure, and save them all
'       After Calucate Order, anything in the OrderDetails and Orders struct can be displayed
'==============================================================================
'
Private Sub CalculateAndSaveOrder()
    On Error GoTo ErrorTrap
    '
    Dim OrderPriceLevelDiscountPrice As Double
    Dim OrderPriceLevelDiscountDiscount As Double
    Dim PriceText As String
    Dim CSTax As Long
    Dim ItemPointer As Long
    Dim SQL As String
    Dim RS As Recordset
    Dim YourPrice As Double
    Dim CatalogAllowSpecialPrice As Boolean
    Dim MemberAllowSpecialPrice As Boolean
    Dim GroupDiscount As Double
    Dim MemberDiscount As Double
    Dim SQLNow As String
    Dim cs As Long
    Dim csItem As Long
    Dim ItemCount As Long
    Dim ItemCountNew As Long
    Dim CreditCardNumber As String
    Dim Pricing As New CmrcPricingClass
    Dim DiscountFactor As Double
    Dim AllowSpecialPrice As Boolean
    Dim SpecialPriceCaption As String
    Dim SalePriceCaption As String
    Dim RegularPriceCaption As String
    Dim PriceRegular As Double
    Dim PriceSpecial As Double
    Dim PriceSale As Double
    Dim GoodPriceSale As Boolean
    Dim GoodPriceRegular As Boolean
    Dim GoodPriceSpecial As Boolean
    '
    ' Calculate item charge, ship weight
    '
    OrderItemCharge = 0
    orderShipWeight = 0
    If main.OrderItemCount > 0 Then
        '
        ' Determine MemberDiscount and MemberAllowSpecialPrice
        '
        If main.SiteProperty_BuildVersion < "3.3.444" Then
            MemberDiscount = 0
            DiscountFactor = 1#
            MemberAllowSpecialPrice = False
        Else
            Call Pricing.GetPricingPolicies(main, CommerceAccountMemberID, AllowGroupPolicies, MemberDiscount, MemberAllowSpecialPrice)
            DiscountFactor = (100# - MemberDiscount) / 100#
        End If
        '
        ' Determine Site Properties
        '
        SpecialPriceCaption = main.GetSiteProperty("CatalogAliasSpecialPrice", "Member Price")
        SalePriceCaption = main.GetSiteProperty("CatalogAliasSalePrice", "Sale")
        RegularPriceCaption = main.GetSiteProperty("CatalogAliasRegularPrice", "Price")
        '
        AllowSpecialPrice = (SpecialPriceCaption <> "") And kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowSpecialPrice", False)) And MemberAllowSpecialPrice
        '
        For ItemPointer = 0 To main.OrderItemCount - 1
            If OrderDetails(ItemPointer).quantity < 1 Then
                OrderDetails(ItemPointer).Active = False
                Call main.DeleteContentRecord("Order Details", OrderDetails(ItemPointer).Id)
            Else
                csItem = main.OpenCSContentRecord("items", OrderDetails(ItemPointer).itemId)
                If Not main.IsCSOK(csItem) Then
                    Call HandleClassErrorAndResumeNext("CalculateAndSaveOrder", "can not find item in item table")
                    OrderDetails(ItemPointer).Active = False
                    Call main.DeleteContentRecord("Order Details", OrderDetails(ItemPointer).Id)
                Else
                    '
                    ' Calculate item Price
                    '
'
' ***** Blank Price means do not show it, 0 price means freee
'
                    PriceText = main.GetCSText(csItem, "UnitPrice")
                    If PriceText = "" Then
                        GoodPriceRegular = False
                    Else
                        PriceRegular = KmaEncodeNumber(PriceText)
                        GoodPriceRegular = (PriceRegular <> 0) And (RegularPriceCaption <> "")
                    End If
                    '
                    PriceText = main.GetCSText(csItem, "SpecialPrice")
                    If PriceText = "" Then
                        GoodPriceSpecial = False
                    Else
                        PriceSpecial = KmaEncodeNumber(PriceText)
                        GoodPriceSpecial = (PriceText <> "") And (SpecialPriceCaption <> "") And MemberAllowSpecialPrice And AllowSpecialPrice
                    End If
                    '
                    PriceText = main.GetCSText(csItem, "SalePrice")
                    If PriceText = "" Then
                        GoodPriceSale = False
                    Else
                        PriceSale = KmaEncodeNumber(PriceText)
                        GoodPriceSale = (main.GetCSBoolean(csItem, "OnSale")) And (PriceSale <> 0) And (SalePriceCaption <> "")
                    End If
                    '
                    YourPrice = Pricing.GetYourPrice(GoodPriceRegular, PriceRegular, GoodPriceSale, PriceSale, GoodPriceSpecial, PriceSpecial, DiscountFactor)
                    '
                    OrderDetails(ItemPointer).Price = YourPrice
                    OrderItemCharge = OrderItemCharge + (OrderDetails(ItemPointer).quantity * YourPrice)
                    orderShipWeight = orderShipWeight + (OrderDetails(ItemPointer).quantity * OrderDetails(ItemPointer).Weight)
                    '
                    ' Update the OrderDetail record with the results
                    '
                    cs = main.OpenCSContentRecord("Order Details", OrderDetails(ItemPointer).Id)
                    If main.IsCSOK(cs) Then
                        Call main.SetCS(cs, "ItemCID", OrderDetails(ItemPointer).ItemCID)
                        Call main.SetCS(cs, "ItemID", OrderDetails(ItemPointer).itemId)
                        Call main.SetCS(cs, "Quantity", OrderDetails(ItemPointer).quantity)
                        Call main.SetCS(cs, "UnitPrice", OrderDetails(ItemPointer).Price)
                        Call main.SetCS(cs, "UnitCost", main.GetCSNumber(csItem, "UnitCost"))
                        Call main.SetCS(cs, "SourceID", main.GetCSInteger(csItem, "OrganizationID"))
                        Call main.SetCS(cs, "SourceItemNumber", main.GetCSText(csItem, "SourceItemNumber"))
                        Call main.SetCS(cs, "Weight", main.GetCSNumber(csItem, "Weight"))
                        Call main.SetCS(cs, "Name", main.GetCSText(csItem, "Name"))
                    End If
                    Call main.CloseCS(cs)
                End If
                Call main.CloseCS(csItem)
                Set RS = Nothing
            End If
        Next
        '
        ' Price Level Discount
        '
        OrderPriceLevelDiscount = 0
        If kmaEncodeBoolean(main.GetSiteProperty("OrderAllowPriceLevelDiscount", False)) Then
            OrderPriceLevelDiscountPrice = KmaEncodeNumber(main.GetSiteProperty("OrderPriceLevelDiscountPrice", "0"))
            OrderPriceLevelDiscountDiscount = KmaEncodeNumber(main.GetSiteProperty("OrderPriceLevelDiscountDiscount", "0"))
            If OrderItemCharge > OrderPriceLevelDiscountPrice And (OrderPriceLevelDiscountDiscount > 0) Then
                OrderPriceLevelDiscount = (OrderItemCharge * OrderPriceLevelDiscountDiscount) / 100
            End If
        End If
        
        '
        ' Calculate ship charge
        '
Dim FreeShipping As Boolean
Dim AllowPriceLevelFreeShipping As Boolean
Dim FreeShippingPrice As Double
        If orderShipMethod <> "" Then
            FreeShipping = False
            AllowPriceLevelFreeShipping = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowFreeShippingPrice", False))
            FreeShippingPrice = KmaEncodeNumber(main.GetSiteProperty("OrderFreeShippingPrice", "0"))
            If AllowPriceLevelFreeShipping Then
                If OrderItemCharge >= FreeShippingPrice Then
                    FreeShipping = True
                End If
            End If
            
            Select Case orderShipMethod
                Case "UPS Next Day Air Early AM"
                    orderShipCharge = getShipPriceUPS("1DM", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Next Day Air"
                    orderShipCharge = getShipPriceUPS("1DA", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Next Day Air Saver"
                    orderShipCharge = getShipPriceUPS("1DP", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS 2nd Day Air AM"
                    orderShipCharge = getShipPriceUPS("2DM", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS 2nd Day Air"
                    orderShipCharge = getShipPriceUPS("2DA", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS 3 Day Select"
                    orderShipCharge = getShipPriceUPS("3DS", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Ground"
                    orderShipCharge = getShipPriceUPS("GND", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Canada Standard"
                    orderShipCharge = getShipPriceUPS("STD", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Worldwide Express"
                    orderShipCharge = getShipPriceUPS("XPR", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Worldwide Express Plus"
                    orderShipCharge = getShipPriceUPS("XDM", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Worldwide Expedited"
                    orderShipCharge = getShipPriceUPS("XPD", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case "UPS Worldwide Express Saver"
                    orderShipCharge = getShipPriceUPS("WXS", CStr(orderShipWeight), orderShipZip, orderShipCountry)
                Case Else
                    '
                    '   Original Order Calculation
                    '
                    cs = main.OpenCSContent("Order Ship Methods", "name=" & KmaEncodeSQLText(orderShipMethod))
                    If main.IsCSOK(cs) Then
                        orderShipCharge = CalcShippingCharge(main.GetCSNumber(cs, "BaseCharge"), main.GetCSNumber(cs, "ChargePerWeight"), main.GetCSNumber(cs, "ChargePerCharge"), orderShipWeight, OrderItemCharge, AllowPriceLevelFreeShipping, FreeShippingPrice)
                    End If
                    Call main.CloseCS(cs)
            End Select
            
        End If
        '
        ' Calculate tax charge
        '
Dim taxState As String
        taxState = orderShipState
        If taxState = "" Then
            taxState = orderBillState
        End If
        OrderTaxCharge = 0
        If OrderItemCharge <> 0 Then
            CSTax = main.OpenCSContent_Internal("States", "(name=" & KmaEncodeSQLText(taxState) & ")or(abbreviation=" & KmaEncodeSQLText(taxState) & ")", , , , , "SalesTax", 1)
            If main.IsCSOK(CSTax) Then
                OrderTaxCharge = OrderItemCharge * main.GetCSNumber(CSTax, "SalesTax")
            End If
            Call main.CloseCS(CSTax)
        End If
        '
        ' Calculate total charge
        '
        OrdertotalCharge = OrderItemCharge + orderShipCharge + OrderTaxCharge - OrderPriceLevelDiscount
        OrdertotalCharge = Int((OrdertotalCharge * 100) + 0.5) / CDbl(100)
    End If
    '
    If (orderId <> 0) Then
        '
        ' Save Order Details, resort array to remove quantity=0 items
        '
        ItemCountNew = 0
        For ItemCount = 0 To main.OrderItemCount - 1
            '
            ' if quantity is 0, make inactive
            '
            If OrderDetails(ItemCount).quantity <= 0 Then
                OrderDetails(ItemCount).quantity = 0
                OrderDetails(ItemCount).Active = False
            End If
            If Not OrderDetails(ItemCount).Active Then
                '
                ' Inactive detail, delete record
                '
                Call main.DeleteContentRecord("Order Details", OrderDetails(ItemCount).Id)
            Else
                '
                ' save the detail
                '
                SQL = "UPDATE OrderDetails SET" _
                    & " Quantity=" & KmaEncodeSQLNumber(OrderDetails(ItemCount).quantity) _
                    & ",UnitPrice=" & KmaEncodeSQLNumber(OrderDetails(ItemCount).Price) _
                    & ",Active=" & KmaEncodeSQLBoolean(OrderDetails(ItemCount).Active) _
                    & " WHERE ID=" & OrderDetails(ItemCount).Id & ";"
                Call main.ExecuteSQL("Default", SQL)
                If ItemCount <> ItemCountNew Then
                    '
                    ' move item to new index
                    '
                    OrderDetails(ItemCountNew).Id = OrderDetails(ItemCount).Id
                    OrderDetails(ItemCountNew).ItemCID = OrderDetails(ItemCount).ItemCID
                    OrderDetails(ItemCountNew).itemId = OrderDetails(ItemCount).itemId
                    OrderDetails(ItemCountNew).itemName = OrderDetails(ItemCount).itemName
                    OrderDetails(ItemCountNew).ItemNumber = OrderDetails(ItemCount).ItemNumber
                    OrderDetails(ItemCountNew).quantity = OrderDetails(ItemCount).quantity
                    OrderDetails(ItemCountNew).Price = OrderDetails(ItemCount).Price
                    OrderDetails(ItemCountNew).Weight = OrderDetails(ItemCount).Weight
                    OrderDetails(ItemCountNew).Active = OrderDetails(ItemCount).Active
                End If
                ItemCountNew = ItemCountNew + 1
            End If
        Next
        main.OrderItemCount = ItemCountNew
        '
        ' Save order
        '
        If kmaEncodeBoolean(main.GetSiteProperty("OrderSaveCreditCard", "0", True)) Then
            CreditCardNumber = OrderCCNumber
        Else
            CreditCardNumber = "[not saved]"
        End If
        SQL = "UPDATE Orders SET " _
            & " BillName = " & KmaEncodeSQLText(orderBillName) _
            & ",BillEmail = " & KmaEncodeSQLText(orderBillEmail) _
            & ",BillPhone = " & KmaEncodeSQLText(orderBillPhone) _
            & ",BillFax = " & KmaEncodeSQLText(orderBillFax) _
            & ",BillCompany = " & KmaEncodeSQLText(orderBillCompany) _
            & ",BillAddress = " & KmaEncodeSQLText(orderBillAddress)
        SQL = SQL _
            & ",BillCity = " & KmaEncodeSQLText(orderBillCity) _
            & ",BillState = " & KmaEncodeSQLText(orderBillState) _
            & ",BillZip = " & KmaEncodeSQLText(orderBillZip) _
            & ",BillCountry = " & KmaEncodeSQLText(orderBillCountry) _
            & ",BillConfirmed = " & KmaEncodeSQLBoolean(orderBillConfirmed)
        SQL = SQL _
            & ",ShipDate = " & KmaEncodeSQLDate(orderShipDate) _
            & ",ShipName = " & KmaEncodeSQLText(orderShipName) _
            & ",ShipCompany = " & KmaEncodeSQLText(orderShipCompany) _
            & ",ShipAddress = " & KmaEncodeSQLText(orderShipAddress) _
            & ",ShipCity = " & KmaEncodeSQLText(orderShipCity) _
            & ",ShipState = " & KmaEncodeSQLText(orderShipState) _
            & ",ShipZip = " & KmaEncodeSQLText(orderShipZip) _
            & ",ShipCountry = " & KmaEncodeSQLText(orderShipCountry) _
            & ",ShipConfirmed = " & KmaEncodeSQLBoolean(orderShipConfirmed) _
            & ",ShipWeight = " & KmaEncodeSQLNumber(orderShipWeight) _
            & ",ShipPhone = " & KmaEncodeSQLText(orderShipPhone) _
            & ",ShipMethod = " & KmaEncodeSQLText(orderShipMethod)
        SQL = SQL _
            & ",ItemCharge = " & KmaEncodeSQLNumber(OrderItemCharge) _
            & ",ShipCharge = " & KmaEncodeSQLNumber(orderShipCharge) _
            & ",TaxCharge = " & KmaEncodeSQLNumber(OrderTaxCharge) _
            & ",TotalCharge = " & KmaEncodeSQLNumber(OrdertotalCharge) _
            & ",PayMethod = " & KmaEncodeSQLText(OrderPayMethod) _
            & ",PayDate = " & KmaEncodeSQLDate(OrderPayDate) _
            & ",CCType = " & KmaEncodeSQLText(OrderCCType) _
            & ",CCNumber = " & KmaEncodeSQLText(CreditCardNumber) _
            & ",CCName = " & KmaEncodeSQLText(OrderCCName)
        If main.SiteProperty_BuildVersion >= "3.3.580" Then
            SQL = SQL _
                & ",PriceLevelDiscount = " & KmaEncodeSQLNumber(OrderPriceLevelDiscount) _
                & ",PriceLevelFreeShipping = " & KmaEncodeSQLBoolean(OrderPriceLevelFreeShipping)
        End If
        If isSupportedAddress2 Then
            SQL = SQL _
                & ",BillAddress2 = " & KmaEncodeSQLText(orderBillAddress2) _
                & ",shipAddress2 = " & KmaEncodeSQLText(orderShipAddress2) _
                & ""
        End If
        SQL = SQL _
            & ",CCExpires = " & KmaEncodeSQLText(OrderCCExpires) _
            & ",CCZip = " & KmaEncodeSQLText(OrderCCZip) _
            & ",MemberID = " & KmaEncodeSQLNumber(OrderMemberID) _
            & ",OrderCompleted = " & KmaEncodeSQLBoolean(OrderCompleted) _
            & ",DateCompleted = " & KmaEncodeSQLDate(OrderDateCompleted) _
            & ",SpecialPricing = " & KmaEncodeSQLBoolean(OrderSpecialPricing) _
            & ",Active = " & KmaEncodeSQLBoolean(OrderActive) _
            & ",visitID = " & KmaEncodeSQLNumber(main.visitId) _
            & ",visitorID = " & KmaEncodeSQLNumber(main.VisitorID) _
            & " WHERE ((Orders.ID)=" & KmaEncodeSQLNumber(orderId) & ");"
        Call main.ExecuteSQL(CommerceDataSourceName, SQL)
        End If
    
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set RS = Nothing
    Call HandleClassErrorAndResumeNext("CalculateAndSaveOrder"): Resume Next
    End Sub
'
'==============================================================================
'   Calculate Shipping Charge for a given shipping method
'       orderShipWeight and OrderCharge must be calculated already
'       returns the charge in dollars
'==============================================================================
'
Private Function CalcShippingCharge(BaseCharge As Double, ChargePerWeight As Double, ChargePerCharge As Double, orderShipWeight As Double, OrderCharge As Double, AllowPriceLevelFreeShipping As Boolean, FreeShippingPrice As Double) As Double
    On Error GoTo ErrorTrap
    '
    If AllowPriceLevelFreeShipping And (OrderCharge > FreeShippingPrice) Then
        CalcShippingCharge = 0
    Else
        CalcShippingCharge = BaseCharge + (orderShipWeight * ChargePerWeight) + (OrderCharge * ChargePerCharge)
    End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("CalcShippingCharge"): Resume Next
    End Function
'
'==============================================================================
'   getCheckoutForm_Buttons( Button, Label )
'       Prints an order process button, Button and label are strings
'==============================================================================
'
Private Function getCheckoutForm_Buttons(Button, Label) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As New FastStringClass
    '
    Call stream.Add("<table border=""0"" cellpadding=""1"" cellspacing=""0"" width=""100%"">")
    Call stream.Add("<tr>")
    Call stream.Add("<td width=""180"" align=""right""><input TYPE=""submit"" NAME=""" & rnButton & """ VALUE=""" & Button & """>&nbsp;<br><img alt="""" src=""/ccLib/Images/spacer.gif"" width=""178"" height=""1"" ></td>")
    Call stream.Add("<td width=""99%"">" & scSpanNormalStyle & Label & "</SPAN></td>")
    Call stream.Add("</tr>")
    Call stream.Add("</table >")
    '
    getCheckoutForm_Buttons = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Buttons"): Resume Next
    End Function
'
'==============================================================================
'   getCheckoutForm_Review()
'       Prints the order review form
'==============================================================================
'
Private Function getCheckoutForm_Review(allowOrdering As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim ItemCount As Long
    Dim SubTotal As Double
    Dim Total As Double
    Dim ItemPointer As Long
    Dim ItemNumber As Long
    Dim itemDescription As String
    Dim ItemNumber2 As String
    Dim stream As New FastStringClass
    Dim CatalogAllowSpecialPrice As Boolean
    Dim CatalogAliasSpecialPrice As String
    Dim IsAdminWithBadMemberID As Boolean
    Dim copy As String
    '
    ' this is a hassle
    IsAdminWithBadMemberID = False
    'IsAdminWithBadMemberID = IsPartOfShoppingCart And (IsMemberCommerceAdmin And (CommerceAccountMemberID = 0))
    If IsAdminWithBadMemberID Then
        Call main.AddUserError("Please use the Account tab to select or add an account. The account is needed to generate pricing.")
    End If
    '
    '   form
    '
    Call stream.Add(scSpanNormalStyle)
    '
    ' Header Copy
    '
    If main.OrderItemCount <= 0 Then
        stream.Add DivBR
        stream.Add GetCommerceCopy(shopCopyNameReviewEmpty, shopCopyDefaultReviewEmpty)
        stream.Add DivBR
    Else
        stream.Add DivBR
        copy = GetCommerceCopy(shopCopyNameReview, shopCopyDefaultReview)
        stream.Add copy
        stream.Add DivBR
    End If
    '
    Call stream.Add(main.GetUserError)
'    If main.OrderItemCount = 0 Then
'        Call stream.Add("<P>" & scSpanNormalStyle & "Your shopping cart is empty.</SPAN></p>")
'    End If
    Call stream.Add(main.GetFormStart())
    Call stream.Add(main.GetFormInputHidden(RequestNameRefreshBlock, main.GetFormSN))
    Call stream.Add("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutReview & """>")
    Call stream.Add(getCheckoutForm_ItemOrderForm(False, True, allowOrdering))
    '
    ' buttons
    '
    stream.Add ("<P>&nbsp;</p>")
    If main.OrderItemCount > 0 Then
        If IsPartOfShoppingCart Then
            If IsAdminWithBadMemberID Then
                Call stream.Add(main.GetPanelButtons(shopButtonBackToShopping & "," & shopButtonRemoveAll & "," & shopButtonRecalculate, rnButton))
            Else
                Call stream.Add(main.GetPanelButtons(shopButtonBackToShopping & "," & shopButtonRemoveAll & "," & shopButtonRecalculate & "," & shopButtonContinue, rnButton))
            End If
        Else
            If IsAdminWithBadMemberID Then
                Call stream.Add(main.GetPanelButtons(shopButtonRemoveAll & "," & shopButtonRecalculate, rnButton))
            Else
                Call stream.Add(main.GetPanelButtons(shopButtonRemoveAll & "," & shopButtonRecalculate & "," & shopButtonContinue, rnButton))
            End If
        End If
    ElseIf IsMemberCommerceAdmin Then
        If IsPartOfShoppingCart Then
            Call stream.Add(main.GetPanelButtons(shopButtonBackToShopping & "," & shopButtonRemoveAll & "," & shopButtonRecalculate, rnButton))
        Else
            Call stream.Add(main.GetPanelButtons(shopButtonRemoveAll & "," & shopButtonRecalculate, rnButton))
        End If
    Else
        If IsPartOfShoppingCart Then
            Call stream.Add(main.GetPanelButtons(shopButtonBackToShopping, rnButton))
        End If
    End If
    '
    Call stream.Add("</FORM>")
    '
    getCheckoutForm_Review = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Review"): Resume Next
    End Function
'
'==============================================================================
'   getCheckoutForm_Security()
'       Prints the order security form
'==============================================================================
'
Private Function getCheckoutForm_Security() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As New FastStringClass
    Dim cs As Long
    Dim AllowSecurePage As Boolean
    Dim PageCopy As String
    '
    '
    'Stream.Add ("<P align=""center""><img src=""/ccLib/images/ShopTimelineSecurity.gif"" width=""391"" height=""44"" ></p>")
    stream.Add (scSpanNormalStyle)
    '
    ' Header Copy
    '
    stream.Add DivBR
    stream.Add GetCommerceCopy(shopCopyNameSecurity, shopCopyDefaultSecurity)
    stream.Add DivBR
    '
    Call stream.Add(main.GetUserError)
    Call stream.Add(main.GetFormStart)
    Call stream.Add(main.GetFormInputHidden(RequestNameRefreshBlock, main.GetFormSN))
    Call stream.Add("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutSecurity & """>")
    '
    ' buttons
    '
    If kmaEncodeBoolean(main.GetSiteProperty("OrderAllowSecurePath")) Then
        stream.Add getCheckoutForm_Buttons(shopButtonSecure, "Continue using a secure connection")
        End If
    stream.Add getCheckoutForm_Buttons(shopButtonUnsecure, "Continue")
    stream.Add getCheckoutForm_Buttons(shopButtonBack, "Back to the Previous Step")
    stream.Add ("</FORM>")
    getCheckoutForm_Security = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Security"): Resume Next
    End Function
'
'==============================================================================
'   getCheckoutForm_Login()
'       Prints the account login form
'==============================================================================
'
Private Function getCheckoutForm_Login() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As New FastStringClass
    '
    ' tell em about it
    '
    'Stream.Add ("<P align=""center""><img src=""/ccLib/images/ShopTimelinePassword.gif"" width=""391"" height=""44"" ></p>")
    stream.Add (scSpanNormalStyle)
    stream.Add ("<br>")
    '
    ' Header Copy
    '
    stream.Add DivBR
    stream.Add GetCommerceCopy(shopCopyNameLogin, shopCopyDefaultLogin)
    stream.Add DivBR
    '
    stream.Add (main.GetUserError)
    stream.Add ("<br>")
    '
    ' If logged on, print name and continue button
    '
    If main.IsAuthenticated Then
    'If Main.IsMember Then
        stream.Add (main.GetFormStart)
        stream.Add ("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutLogin & """>")
        stream.Add ("<table width=""100%"" cellpadding=""0"" cellspacing=""5"" valign=""top"" border=""0"">")
        stream.Add ("<tr>")
        stream.Add ("<td colspan=""3"" align=""left"">" & scSpanNormalStyle)
        stream.Add ("<B>" & main.MemberName & "</B><br>")
        stream.Add ("<br>")
        stream.Add getCheckoutForm_Buttons(shopButtonContinue, "If this is your account.")
        stream.Add ("<br>")
        stream.Add ("</SPAN></td>")
        stream.Add ("</tr>")
        stream.Add ("</table >")
        stream.Add ("</FORM>")
    End If
    '
    ' returning customers
    '
    stream.Add (main.GetFormStart)
    stream.Add ("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutLogin & """>")
    stream.Add ("<table width=""100%"" cellpadding=""0"" cellspacing=""5"" valign=""top"" border=""0"">")
    stream.Add ("<tr>")
    stream.Add ("<td colspan=""3"" align=""left"">" & scSpanNormalStyle)
    stream.Add ("<B>Returning Customers</B><br>")
    stream.Add ("Enter your username and password. If you don't remember your password, use the SEND PASSWORD button below.")
    stream.Add ("<br>&nbsp;</SPAN></td></tr>")
    stream.Add ("<tr>")
    stream.Add ("<td align=""right"" width=""140"">" & scSpanNormalStyle & "Username</SPAN></td>")
    stream.Add ("<td align=""left"" width=""200""><input type=""text"" size=""22"" name=""oau"" VALUE=""""></td>")
    stream.Add ("<td align=""left"" width=""99%"">" & scSpanNormalStyle & "</SPAN></td>")
    stream.Add ("</tr>")
    stream.Add ("<tr>")
    stream.Add ("<td align=""right"" width=""140"">" & scSpanNormalStyle & "Password</SPAN></td>")
    stream.Add ("<td align=""left"" width=""200""><input type=""password"" size=""22"" name=""oap"" VALUE="""" maxlength=""20""></td>")
    stream.Add ("<td align=""left"" width=""99%"">" & scSpanNormalStyle & "</SPAN></td>")
    stream.Add ("</tr>")
    stream.Add ("<tr>")
    stream.Add ("<tr><td colspan=""3"" align=""left"">")
    stream.Add getCheckoutForm_Buttons(shopButtonLogin, "For previous customers.")
    stream.Add ("</SPAN>")
    stream.Add ("</td></tr>")
    stream.Add ("</table >")
    stream.Add ("</FORM>")
    '
    ' create a new account
    '
    stream.Add vbCrLf & vbTab & "<fieldset>"
    stream.Add vbCrLf & vbTab & "<legend>New Customers</legend>"
    stream.Add (main.GetFormStart)
    stream.Add ("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutLogin & """>")
    stream.Add ("<table width=""100%"" cellpadding=""0"" cellspacing=""5"" valign=""top"" border=""0"">")
    stream.Add ("<tr>")
    stream.Add ("<td colspan=""3"" align=""left"">" & scSpanNormalStyle)
    stream.Add ("<B>New Customers</B><br>")
    stream.Add ("To create an account, select a username and password. The username must be unique on this site.")
    stream.Add ("<br>&nbsp;</SPAN></td></tr>")
    stream.Add ("<tr>")
    stream.Add ("<td align=""right"" width=""140"">" & scSpanNormalStyle & "Username</SPAN></td>")
    stream.Add ("<td align=""left"" width=""200""><input type=""text"" size=""22"" name=""oau"" VALUE=""""></td>")
    stream.Add ("<td align=""left"" width=""99%"">" & scSpanNormalStyle & "</SPAN></td>")
    stream.Add ("</tr>")
    stream.Add ("<tr>")
    stream.Add ("<td align=""right"" width=""140"">" & scSpanNormalStyle & "Password</SPAN></td>")
    stream.Add ("<td align=""left"" width=""200""><input type=""password"" size=""22"" name=""oap"" VALUE="""" maxlength=""20""></td>")
    stream.Add ("<td align=""left"" width=""99%"">" & scSpanNormalStyle & "</SPAN></td>")
    stream.Add ("</tr>")
    stream.Add ("<tr><td colspan=""3"" align=""left"">")
    stream.Add getCheckoutForm_Buttons(shopButtonCreateAccount, " For New Customers.")
    stream.Add ("</td></tr>")
    stream.Add ("</table >")
    stream.Add ("</FORM>")
    stream.Add vbCrLf & vbTab & "</fieldset>"
    '
    ' forget your username and password
    '
    stream.Add (main.GetFormStart)
    stream.Add ("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutLogin & """>")
    stream.Add ("<table width=""100%"" cellpadding=""0"" cellspacing=""5"" valign=""top"" border=""0"">")
    stream.Add ("<tr>")
    stream.Add ("<td colspan=""3"" align=""left"">" & scSpanNormalStyle)
    stream.Add ("<B>Forget Your Username/Password</B><br>")
    stream.Add ("If you have created an account but can not remember your password, enter your email address and we will email the username and password associated with your email address to the address.")
    stream.Add ("<br>&nbsp;</SPAN></td></tr>")
    stream.Add ("<tr>")
    stream.Add ("<td align=""right"" width=""140"">" & scSpanNormalStyle & "Email</SPAN></td>")
    stream.Add ("<td align=""left"" width=""200""><input type=""text"" size=""22"" name=""oae"" VALUE=""""></td>")
    stream.Add ("<td align=""left"" width=""99%"">" & scSpanNormalStyle & "</SPAN></td>")
    stream.Add ("</tr>")
    stream.Add ("<tr><td colspan=""3"" align=""left"">")
    stream.Add getCheckoutForm_Buttons(shopButtonSendPassword, "Returning customers, Email is required.")
    stream.Add ("</td></tr>")
    stream.Add ("</table >")
    stream.Add ("</FORM>")
    '
    ' Back Buttons
    '
    stream.Add (main.GetFormStart)
    stream.Add ("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutLogin & """>")
    '
    stream.Add getCheckoutForm_Buttons(shopButtonBack, "Back to the Previous Step")
    stream.Add ("</FORM>")
    '
    getCheckoutForm_Login = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Login"): Resume Next
    End Function
'
'==============================================================================
'   getCheckoutForm_ShippingBilling()
'==============================================================================
'
Private Function getCheckoutForm_ShippingBilling() As String
    On Error GoTo ErrorTrap
    '
    Dim useContact As Boolean
    Dim testid As Long
    Dim iIsShippingConfigured As Boolean
    Dim CSShipMethods As Long
    Dim PaymentSelect As String
    Dim payTypeLast As String
    Dim SQL As String
    Dim RS As Recordset
    Dim RSShipMethods As Recordset
    Dim optionString As String
    Dim TempCharge As Double
    Dim ShippingDaysMin As Long
    Dim ShippingDaysMax As Long
    Dim MethodName As String
    Dim ShipMethodName As String
    Dim stream As New FastStringClass
    Dim cs As Long
    Dim PayMethodCount As Long
    Dim PayMethodFound As Boolean
    Dim PayDetailsString As String
    Dim ShippingSelect As String
    Dim OptionName As String
    Dim OptionFound As Boolean
    Dim IsFormRetry As Boolean
    Dim IsError As Boolean
    Dim IsRequired As Boolean
    Dim ShipOptionsJS As String
    Dim BillOptionsJS As String
    Dim ShipCountrySelectOptions As String
    Dim BillCountrySelectOptions As String
    Dim CountryStateList As String
    Dim CurrentCountryID As Long
    Dim countryId As Long
    Dim countryName As String
    Dim StateID As Long
    Dim StateName As String
    Dim StateSelects() As String
    Dim stateSelect As String
    Dim StateSelectPtr As Long
    Dim StateSelectCnt As Long
    Dim CountrySelect As String
    Dim IsCurrentShipCountry As Boolean
    Dim IsCurrentShipState As Boolean
    Dim IsCurrentBillCountry As Boolean
    Dim IsCurrentBillState As Boolean
    Dim UCaseorderShipCountry As String
    Dim UcaseorderShipState As String
    Dim UCaseorderBillCountry As String
    Dim UcaseorderBillState As String
    Dim InitialStatePtr As Long
    Dim DivName() As String
    Dim DefaultJS As String
    Dim billToContactAddress As Boolean
    Dim shipToContactAddress As Boolean
    Dim ShipJS As String
    Dim checkedAttr As String
    Dim checkedStyle As String
    Dim subscriptionFlag As Boolean
    Dim common As New commonClass
    Dim blockContactSave As Boolean
    Dim blockContactGroup As String
    Dim checkCreateKey As Long
    Dim csCheck As Long
    '
    blockContactSave = False
    If main.IsAuthenticated Then
        blockContactGroup = Trim(main.GetSiteProperty("ShoppingCart Block Contact Update Group", ""))
        If blockContactGroup <> "" Then
            blockContactSave = common.isGroupMember(main, blockContactGroup, CommerceAccountMemberID)
        End If
    End If
    '
    MethodName = "getCheckoutForm_ShippingBilling()"
    subscriptionFlag = isSubscriptionOrder()
    '
    'Main.TestPoint "getCheckoutForm_ShippingBilling, orderBillCountry=" & orderBillCountry
    'Main.TestPoint "getCheckoutForm_ShippingBilling, orderShipCountry=" & orderShipCountry
    '
    IsFormRetry = (main.GetStreamText(rnSrcShopFormId) = shopFormCheckoutShippingBilling)
    CurrentCountryID = 0
    UCaseorderShipCountry = UCase(orderShipCountry)
    UcaseorderShipState = UCase(orderShipState)
    UCaseorderBillCountry = UCase(orderBillCountry)
    UcaseorderBillState = UCase(orderBillState)
    '
    ' Set defaults
    '
    If UCaseorderShipCountry = "" Then
        UCaseorderShipCountry = "UNITED STATES"
    End If
    If UCaseorderBillCountry = "" Then
        UCaseorderBillCountry = "UNITED STATES"
    End If
    '
    ' Header Copy
    '
    stream.Add (scSpanNormalStyle)
    stream.Add DivBR
    stream.Add GetCommerceCopy(shopCopyNameContact, shopCopyDefaultContact)
    '
    stream.Add (main.GetUserError)
    stream.Add (main.GetFormStart)
    stream.Add ("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutShippingBilling & """>")
    stream.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"" >")
    '
    ' Contact Information
    '
    stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Contact Information</SPAN></td></tr>")
    If blockContactSave Then
        '
        ' must be autheniticated, and in group that can not edit contact info
        '
        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameContactAddressNoEdit, shopCopyDefaultContactAddressNoEdit) & "</td></tr>")
        stream.Add getCheckoutForm_ShippingBilling_Row("Name", orderContactName, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("Email", orderContactEmail, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("Company", orderContactCompany, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("Address", orderContactAddress, False, False, True)
        If isSupportedAddress2 Then
            stream.Add getCheckoutForm_ShippingBilling_Row("Address Line 2", orderContactAddress2, False, False, True)
        End If
        stream.Add getCheckoutForm_ShippingBilling_Row("City", orderContactCity, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("State", orderContactState, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("Country", orderContactCountry, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("Postal Code", orderContactZip, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("Phone", orderContactPhone, False, False, True)
        stream.Add getCheckoutForm_ShippingBilling_Row("Fax", orderContactFax, False, False, True)
        stream.Add "<input type=""hidden"" id=""orderContactZip"" name=""orderContactZip"" value=""" & orderContactZip & """>"
        stream.Add "<input type=""hidden"" id=""orderContactCountryValue"" name=""orderContactCountry"" value=""" & orderContactCountry & """>"
    Else
        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameContactAddress, shopCopyDefaultContactAddress) & "</td></tr>")
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Name", "orderContactName", orderContactName, True, IsFormRetry)
        stream.Add getCheckoutForm_ShippingBilling_Row("Email", "<input name=orderContactEmail VALUE=""" & orderContactEmail & """ SIZE=""30"" MAXLENGTH=""50"">", True, IsFormRetry And IsOrderEmailBad)
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Company", "orderContactCompany", orderContactCompany, False, IsFormRetry)
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Address", "orderContactAddress", orderContactAddress, True, IsFormRetry)
        If isSupportedAddress2 Then
            stream.Add getCheckoutForm_ShippingBilling_Row_Text("Address Line 2", "orderContactAddress2", orderContactAddress2, False, IsFormRetry)
        End If
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("City", "orderContactCity", orderContactCity, True, IsFormRetry)
        stream.Add vbCrLf & vbCrLf & "<!-- Billing Select -->" & vbCrLf
        CountrySelect = getCountrySelect(main, orderContactCountry, "orderContactCountry", "orderContactState", "cartContactStateSelect", "orderContactState")
        stream.Add getCheckoutForm_ShippingBilling_Row("Country", CountrySelect, True, IsFormRetry And (orderContactCountry = ""))
        optionString = "country=" & encodeNvaArgument(orderContactCountry) & "&state=" & encodeNvaArgument(orderContactState) & "&element=orderContactState"
        stateSelect = main.ExecuteAddon3("cartGetStateSelect", optionString)
        stream.Add getCheckoutForm_ShippingBilling_Row("State", "<div id=""cartContactStateSelect"">" & stateSelect & "</div>", True, (IsFormRetry And (orderContactState = "") And (LCase(orderContactCountry) = "united states")))
        stream.Add vbCrLf & vbCrLf & vbCrLf
        stream.Add getCheckoutForm_ShippingBilling_Row_Text_With_Function("Postal Code", "orderContactZip", orderContactZip, True, IsFormRetry)
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Phone", "orderContactPhone", orderContactPhone, True, IsFormRetry)
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Fax", "orderContactFax", orderContactFax, False, IsFormRetry)
    End If
    '
    ' Username and password for future login if not logged in
    '
    csCheck = main.OpenCSContent("people", "id=" & CommerceAccountMemberID)
    If main.IsCSOK(csCheck) Then
        'userPassword = main.GetCSText(csCheck, "password")
        checkCreateKey = main.GetCSInteger(csCheck, "createKey")
    End If
    Call main.CloseCS(csCheck)
    If main.visitId <> checkCreateKey Then
        '
        ' account not created here, do not offer login options
        '
    Else
        '
        '
        '
        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Create a Login</td></tr>")
        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameContactAccountLogin, shopCopyDefaultContactAccountLogin) & "</td></tr>")
        If Not kmaEncodeBoolean(main.GetSiteProperty("ALLOWEMAILLOGIN", "")) Then
            stream.Add getCheckoutForm_ShippingBilling_Row("Username", "<input name=AccountUsername type=text value=""" & AccountUsername & """ width=20>", True, AccountUsernameError)
        End If
        stream.Add getCheckoutForm_ShippingBilling_Row("Password", "<input name=Accountpassword type=password value=""" & accountPassword & """ width=20>", True, AccountPasswordError)
    End If
    '
    ' Billing Address
    '   - billToContact if billing is empty OR billing matches contact
    '
    If main.GetStreamInteger(rnSrcShopFormId) = shopFormCheckoutShippingBilling Then
        useContact = main.GetStreamBoolean("billToContactAddress")
    Else
        useContact = True _
            And (orderBillName = orderContactName) _
            And (orderBillCompany = orderContactCompany) _
            And (orderBillAddress = orderContactAddress) _
            And (orderBillAddress = orderContactAddress2) _
            And (orderBillCity = orderContactCity) _
            And (orderBillCountry = orderContactCountry) _
            And (orderBillState = orderContactState) _
            And (orderBillZip = orderContactZip) _
            And (orderBillPhone = orderContactPhone)
        If Not useContact Then
            useContact = True _
                And (orderBillName = "") _
                And (orderBillCompany = "") _
                And (orderBillAddress = "") _
                And (orderBillAddress2 = "") _
                And (orderBillCity = "") _
                And (orderBillCountry = "") _
                And (orderBillState = "") _
                And (orderBillZip = "") _
                And (orderBillPhone = "")
        End If
    End If
    '
    If useContact Then
        checkedAttr = " checked "
        checkedStyle = "display:none;"
    Else
        checkedAttr = ""
        checkedStyle = "display:block;"
    End If
    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
    stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Billing Information</td></tr>")
    stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameBillAddress, shopCopyDefaultBillAddress) & "</td></tr>")
    stream.Add ("<tr><td colspan=2>" & scSpanNormalStyle & "<input type=checkbox id=""billToContactAddress"" name=billToContactAddress value=1 " & checkedAttr & " onclick=""toggleHide(this.checked,'billInfo');cartResetShippingSelect();"">Same as Contact Address</SPAN></td></tr>")
    stream.Add ("<tr><td colspan=2>")
    stream.Add ("<div ID=""billInfo"" style=""" & checkedStyle & """>")
    stream.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"" >")
    stream.Add getCheckoutForm_ShippingBilling_Row_Text("Name", "orderBillName", orderBillName, True, IsFormRetry)
    stream.Add getCheckoutForm_ShippingBilling_Row("Email", "<input name=orderBillEmail VALUE=""" & orderBillEmail & """ SIZE=""30"" MAXLENGTH=""50"">", False, IsFormRetry And IsOrderEmailBad)
    stream.Add getCheckoutForm_ShippingBilling_Row_Text("Company", "orderBillCompany", orderBillCompany, False, IsFormRetry)
    stream.Add getCheckoutForm_ShippingBilling_Row_Text("Address", "orderBillAddress", orderBillAddress, True, IsFormRetry)
    If isSupportedAddress2 Then
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Address Line 2", "orderBillAddress2", orderBillAddress2, False, IsFormRetry)
    End If
    stream.Add getCheckoutForm_ShippingBilling_Row_Text("City", "orderBillCity", orderBillCity, True, IsFormRetry)
    stream.Add vbCrLf & vbCrLf & "<!-- Billing Select -->" & vbCrLf
    CountrySelect = getCountrySelect(main, orderBillCountry, "orderBillCountry", "orderBillState", "cartBillStateSelect", "orderBillState")
    stream.Add getCheckoutForm_ShippingBilling_Row("Country", CountrySelect, True, IsFormRetry And (orderBillCountry = ""))
    optionString = "country=" & encodeNvaArgument(orderBillCountry) & "&state=" & encodeNvaArgument(orderBillState) & "&element=orderBillState"
    stateSelect = main.ExecuteAddon3("cartGetStateSelect", optionString)
    stream.Add getCheckoutForm_ShippingBilling_Row("State", "<div id=""cartBillStateSelect"">" & stateSelect & "</div>", True, (IsFormRetry And (orderBillState = "") And (LCase(orderBillCountry) = "united states")))
    stream.Add getCheckoutForm_ShippingBilling_Row_Text_With_Function("Postal Code", "orderBillZip", orderBillZip, True, IsFormRetry)
    stream.Add getCheckoutForm_ShippingBilling_Row_Text("Phone", "orderBillPhone", orderBillPhone, True, IsFormRetry)
    stream.Add getCheckoutForm_ShippingBilling_Row_Text("Fax", "orderBillFax", orderBillFax, False, IsFormRetry)
    '
    stream.Add ("</table >")
    stream.Add ("</div>")
    stream.Add ("</td></tr>")
    '
    ' Shipping address
    '
    If orderShippingRequired Then
        If main.GetStreamInteger(rnSrcShopFormId) = shopFormCheckoutShippingBilling Then
            useContact = main.GetStreamBoolean("ShipToContactAddress")
        Else
            useContact = True _
                And (orderShipName = orderContactName) _
                And (orderShipCompany = orderContactCompany) _
                And (orderShipAddress = orderContactAddress) _
                And (orderShipAddress = orderContactAddress2) _
                And (orderShipCity = orderContactCity) _
                And (orderShipCountry = orderContactCountry) _
                And (orderShipState = orderContactState) _
                And (orderShipZip = orderContactZip) _
                And (orderShipPhone = orderContactPhone)
            If Not useContact Then
                useContact = True _
                    And (orderShipName = "") _
                    And (orderShipCompany = "") _
                    And (orderShipAddress = "") _
                    And (orderShipAddress2 = "") _
                    And (orderShipCity = "") _
                    And (orderShipCountry = "") _
                    And (orderShipState = "") _
                    And (orderShipZip = "") _
                    And (orderShipPhone = "")
            End If
        End If
        '
        If useContact Then
            checkedAttr = " checked "
            checkedStyle = "display:none;"
        Else
            checkedAttr = ""
            checkedStyle = "display:block;"
        End If
        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Shipping Information</td></tr>")
        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameShipAddress, shopCopyDefaultShipAddress) & "</td></tr>")
        stream.Add ("<tr><td colspan=2>" & scSpanNormalStyle & "<input type=checkbox id=""ShipToContactAddress"" name=ShipToContactAddress value=1 " & checkedAttr & " onclick=""toggleHide(this.checked,'shipInfo');cartResetShippingSelect();"">Same as Contact Address</SPAN></td></tr>")
        stream.Add ("<tr><td colspan=2>")
        stream.Add ("<div ID=""shipInfo"" style=""" & checkedStyle & """>")
        stream.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"" >")
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Name", "orderShipName", orderShipName, orderShippingRequired, IsFormRetry)
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Company", "orderShipCompany", orderShipCompany, False, IsFormRetry)
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Address", "orderShipAddress", orderShipAddress, orderShippingRequired, IsFormRetry)
        If isSupportedAddress2 Then
            stream.Add getCheckoutForm_ShippingBilling_Row_Text("Address Line 2", "orderShipAddress2", orderShipAddress2, False, IsFormRetry)
        End If
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("City", "orderShipCity", orderShipCity, orderShippingRequired, IsFormRetry)
        stream.Add vbCrLf & vbCrLf & "<!-- Shipping Select -->" & vbCrLf
        '
        CountrySelect = getCountrySelect(main, orderShipCountry, "orderShipCountry", "orderShipState", "cartShipStateSelect", "orderShipState")
        stream.Add getCheckoutForm_ShippingBilling_Row("Country", CountrySelect, True, IsFormRetry And (orderBillCountry = ""))
        optionString = "country=" & encodeNvaArgument(orderShipCountry) & "&state=" & encodeNvaArgument(orderShipState) & "&element=orderShipState"
        stateSelect = main.ExecuteAddon3("cartGetStateSelect", optionString)
        stream.Add getCheckoutForm_ShippingBilling_Row("State", "<div id=""cartShipStateSelect"">" & stateSelect & "</div>", True, (IsFormRetry And (orderBillState = "") And (LCase(orderBillCountry) = "united states")))
        '
        stream.Add vbCrLf & vbCrLf & vbCrLf
        stream.Add getCheckoutForm_ShippingBilling_Row_Text_With_Function("Postal Code", "orderShipZip", orderShipZip, orderShippingRequired, IsFormRetry)
        stream.Add getCheckoutForm_ShippingBilling_Row_Text("Phone", "orderShipPhone", orderShipPhone, orderShippingRequired, IsFormRetry)
        stream.Add ("<tr><td width=140><img src=/ccLib/Images/spacer.gif width=140 height=1></td><td width=99%></td></tr>")
        stream.Add ("</table >")
        stream.Add ("</Div>")
        '
        stream.Add ("</td></tr>")
        '
        ' Shipping Method select
        '
        iIsShippingConfigured = isShippingConfigured()
        If Not iIsShippingConfigured Then
            ShippingSelect = "<select name=""orderShipMethod""><option value=""Free Shipping"" selected>Free Shipping</option></select>"
        Else
            ShippingSelect = main.GetFormInputHidden("shipMethod", orderShipMethod, "shipMethod")
            ShippingSelect = ShippingSelect & main.GetFormInputHidden("shipWeight", orderShipWeight, "shipWeight")
            ShippingSelect = ShippingSelect & main.GetFormInputHidden("shipCharge", OrderItemCharge, "shipCharge")
            ShippingSelect = ShippingSelect & "<div id=""shipContainer""><select onClick=""cartDrawShippingSelect();""><option>Select One</option></select></div>"
        End If
        '
        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Shipping Method</td></tr>")
        If (Not iIsShippingConfigured) And (IsMemberCommerceAdmin) Then
            '
            ' Create PageBanner so admin will know what mode they are in
            '
            stream.Add "<tr><td colspan=2>" _
                & "<div class=ccEditWrapperCaption><h2>Commerce Administrator</h2>" _
                & "<p>Shipping is required but you have not configured any shipping methods. Configure shipping methods on the <a href=""" & main.SiteProperty_AdminURL & "?addonguid=%7B04290472-281A-49F4-B0A0-1A9FDD1A8E4D%7D"">Shopping Cart Settings</a> page under Shipping Properties.</p>" _
                & "</td></tr>"
        End If
        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameShipMethod, shopCopyDefaultShipMethod) & "</td></tr>")
        stream.Add getCheckoutForm_ShippingBilling_Row("&nbsp;", ShippingSelect, True, (IsRequired And IsFormRetry) Or (IsorderShipMethodBad))
    End If
    '
    '   Payment Method
    '
    If OrdertotalCharge = 0 Then
        PayMethodCount = PayMethodCount + 1
        PaymentSelect = "There is no payment required for your order.<input type=""hidden"" name=""OrderPayMethod"" value=""" & PayString(PayTypeNONE) & """>"
    Else
        PayMethodFound = False
        '
        ' ----- Get credit cards accepted
        '
        PayDetailsString = ""
        If OrderAllowVisa Then
            PayDetailsString = PayDetailsString & ",Visa"
        End If
        If OrderAllowMC Then
            PayDetailsString = PayDetailsString & ",MC"
        End If
        If OrderAllowDiscover Then
            PayDetailsString = PayDetailsString & ",Discover"
        End If
        If OrderAllowAmex Then
            PayDetailsString = PayDetailsString & ",American Express"
        End If
        If PayDetailsString <> "" Then
            PayDetailsString = " (" & Mid(PayDetailsString, 2) & ")"
        End If
        '
        ' ----- Special compatiblity case - Credit Cards are on and other options off
        '
        OrderAllowCreditCards = OrderAllowVisa Or OrderAllowMC Or OrderAllowDiscover Or OrderAllowAmex
        If OrderAllowCreditCards And Not (OrderAllowCreditCardOnline Or OrderAllowCreditCardByPhone Or OrderAllowCreditCardByFax) Then
            OrderAllowCreditCardOnline = True
            OrderAllowCreditCardByPhone = True
            OrderAllowCreditCardByFax = True
        End If
        '
        ' ----- Net terms
        '
        If OrderAllowNetTerms Then
            payTypeLast = PayTypeNetTerms
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeNetTerms, OrderPayMethod)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeNetTerms))
        End If
        '
        ' ----- COD Company Check
        '
        If OrderAllowCODCompanyCheck Then
            payTypeLast = PayTypeCODCompanyCheck
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCODCompanyCheck, OrderPayMethod)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCODCompanyCheck))
        End If
        '
        ' ----- COD Certified Funds
        '
        If OrderAllowCODCertifiedFunds Then
            payTypeLast = PayTypeCODCertifiedFunds
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCODCertifiedFunds, OrderPayMethod)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCODCertifiedFunds))
        End If
        '
        ' ----- Credit card online
        '
        If OrderAllowCreditCardOnline Then
            payTypeLast = PayTypeCreditCardOnline
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCreditCardOnline, OrderPayMethod, PayDetailsString)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCreditCardOnline))
        End If
        '
        ' ----- Credit card by phone
        '
        If OrderAllowCreditCardByPhone Then
            payTypeLast = PayTypeCreditCardByPhone
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCreditCardByPhone, OrderPayMethod, PayDetailsString)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCreditCardByPhone))
        End If
        '
        ' ----- Credit card by fax
        '
        If OrderAllowCreditCardByFax Then
            payTypeLast = PayTypeCreditCardByFax
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCreditCardByFax, OrderPayMethod, PayDetailsString)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCreditCardByFax))
        End If
        '
        ' ----- Personal Checks
        '
        If OrderAllowPersonalChecks Then
            payTypeLast = PayTypeCHECK
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCHECK, OrderPayMethod)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCHECK))
        End If
        '
        ' ----- Company Checks
        '
        If OrderAllowCompanyChecks Then
            payTypeLast = PayTypeCHECKCOMPANY
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCHECKCOMPANY, OrderPayMethod)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCHECKCOMPANY))
        End If
        '
        ' ----- PayPal
        '
        If OrderAllowPayPal Then
            payTypeLast = PayTypePAYPAL
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypePAYPAL, OrderPayMethod)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypePAYPAL))
        End If
        '
        ' ----- On Credit
        '
        If OrderAllowCredit Then
            payTypeLast = PayTypeCREDIT
            PaymentSelect = PaymentSelect & getCheckoutForm_HTMLBillingMethod(PayTypeCREDIT, OrderPayMethod)
            PayMethodCount = PayMethodCount + 1
            PayMethodFound = PayMethodFound Or (OrderPayMethod = PayString(PayTypeCREDIT))
        End If
    End If
    If PayMethodCount = 0 Then
        '
        ' no pay method, bill then
        '
        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Payment Method</SPAN></td></tr>")
        If (Not iIsShippingConfigured) And (IsMemberCommerceAdmin) Then
            '
            ' Create PageBanner so admin will know what mode they are in
            '
            stream.Add "<tr><td colspan=2>" _
                & "<div class=ccEditWrapperCaption><h2>Commerce Administrator</h2>" _
                & "<p>Payment is required but you have not configured any payment methods. Configure payment methods on the <a href=""" & main.SiteProperty_AdminURL & "?addonguid=%7B04290472-281A-49F4-B0A0-1A9FDD1A8E4D%7D"">Shopping Cart Settings</a> page by checking Accept PayPal Payments, or by checking one or more of the options under Accept Online Payments. If you choose online credit card processing you will also have to configure one of the payment processor sections in the Payment Processing add-on.</p>" _
                & "</td></tr>"
        End If
        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>&nbsp;</td></tr>")
        'stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNamePaymentMethod, shopCopyDefaultPaymentMethod) & "</td></tr>")
        PaymentSelect = "You will be billed for this order.<input type=""hidden"" name=""OrderPayMethod"" value=""" & PayString(PayTypeCREDIT) & """>"
        stream.Add getCheckoutForm_ShippingBilling_Row("&nbsp;", PaymentSelect, False, IsFormRetry And (OrderPayMethod = ""))
    ElseIf PayMethodCount = 1 Then
        '
        ' one pay method, no need to ask
        '
        stream.Add ("<tr><td colspan=2><input type=""hidden"" name=""orderPayMethod"" value=""" & PayString(payTypeLast) & """></td></tr>")
'        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
'        stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Payment Method</SPAN></td></tr>")
'        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNamePaymentMethod, shopCopyDefaultPaymentMethod) & "</td></tr>")
'        stream.Add ("<input type=""hidden"" name=""orderPayMethod"" value=""" & payTypeLast & """>")
    Else
        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Payment Method</SPAN></td></tr>")
        stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNamePaymentMethod, shopCopyDefaultPaymentMethod) & "</td></tr>")
        PaymentSelect = Mid(PaymentSelect, 5)
        stream.Add getCheckoutForm_ShippingBilling_Row("Method&nbsp;&nbsp;", PaymentSelect, True, IsFormRetry And (OrderPayMethod = ""))
    End If
'    '
'    ' Username and password for future login if not logged in
'    '
'    If main.IsAuthenticated() Then
'        '
'        ' they are logged in, assume they have a valid username and password
'        '
'    Else
'        If subscriptionFlag Then
'            '
'            ' username required
'            ' not authenticated - not recognized either because in that case there was a logout
'            '   must add a username and password
'            '
'            stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
'            stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Create a Login</td></tr>")
'            stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameContactSubscriptionAccountLogin, shopCopyDefaultContactSubscriptionAccountLogin) & "</td></tr>")
'            stream.Add getCheckoutForm_ShippingBilling_Row("Username", "<input name=AccountUsername type=text value=""" & AccountUsername & """ width=20>", subscriptionFlag, AccountUsernameError)
'            stream.Add getCheckoutForm_ShippingBilling_Row("Password", "<input name=Accountpassword type=password value="""" width=20>", subscriptionFlag, AccountPasswordError)
'        Else
'            '
'            ' username optional
'            '
'            stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
'            stream.Add ("<tr><td colspan=""2"" align=""left"" class=ccCatRowHeader>Create a Login</td></tr>")
'            stream.Add ("<tr><td style=""padding-top:10px;padding-bottom:10px;"" colspan=2>" & GetCommerceCopy(shopCopyNameContactAccountLogin, shopCopyDefaultContactAccountLogin) & "</td></tr>")
'            stream.Add getCheckoutForm_ShippingBilling_Row("Username", "<input name=AccountUsername type=text value=""" & AccountUsername & """ width=20>", subscriptionFlag, AccountUsernameError)
'            stream.Add getCheckoutForm_ShippingBilling_Row("Password", "<input name=Accountpassword type=password value="""" width=20>", subscriptionFlag, AccountPasswordError)
'        End If
'    End If
'    '
'    '   Update Account Checkbox
'    '
'    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
'    stream.Add ("<tr><td style=""border-top:1px solid black;"" colspan=2>&nbsp;</td></tr>")
'    stream.Add ("<tr><td colspan=""2"">" & scSpanNormalStyle & "<input type=""checkbox"" name=""ignore_UpdateMemberRecord"" value=""ON"" checked> Update your account with this information.</SPAN></td></tr>")
    '
    ' Required fields
    '
    stream.Add ("<tr><td colspan=""2"">" & scSpanNormalStyle & "* required fields</SPAN></td></tr>")
    '
    '   Spacers
    '
    stream.Add ("<tr>")
    stream.Add ("<td width=""140""><img alt="""" src=""/ccLib/Images/spacer.gif"" width=""140"" height=""1"" ></td>")
    stream.Add ("<td width=""99%""></td>")
    stream.Add ("</tr>")
    '
    stream.Add ("</table >")
    '
    ' buttons
    '
    Call stream.Add(main.GetPanelButtons(shopButtonBack & "," & shopButtonContinue, rnButton))
    '
    ' end of section
    '
    stream.Add ("</FORM>")
    '
    ' add paypal form - moved to next form
    '
'    If OrderAllowPayPal Then
'        Stream.Add (getCheckoutForm_Billing_PayPalForm)
'    End If
    '
    ' Add dynamic select code
    '
    'stream.Add ("<SCRIPT>initDynamicOptionLists();</SCRIPT>")
    getCheckoutForm_ShippingBilling = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Set RS = Nothing
    Call HandleClassErrorAndResumeNext("getCheckoutForm_ShippingBilling"): Resume Next
    End Function
'
'
'
Private Function getCheckoutForm_ShippingBilling_Row_Text(Label As String, InputName As String, InputValue As String, IsRequired As Boolean, IsFormRetry As Boolean, Optional addReadOnlyWrapper As Boolean) As String
    getCheckoutForm_ShippingBilling_Row_Text = getCheckoutForm_ShippingBilling_Row(Label, "<input name=""" & InputName & """ VALUE=""" & InputValue & """ SIZE=""30"" MAXLENGTH=""50"">", IsRequired, IsRequired And IsFormRetry And (InputValue = ""), addReadOnlyWrapper)
End Function
'
'
'
Private Function getCheckoutForm_ShippingBilling_Row_Text_With_Function(Label As String, InputName As String, InputValue As String, IsRequired As Boolean, IsFormRetry As Boolean, Optional addReadOnlyWrapper As Boolean) As String
    getCheckoutForm_ShippingBilling_Row_Text_With_Function = getCheckoutForm_ShippingBilling_Row(Label, "<input onkeyup=""cartResetShippingSelect();"" name=""" & InputName & """ id=""" & InputName & """ VALUE=""" & InputValue & """ SIZE=""30"" MAXLENGTH=""50"">", IsRequired, IsRequired And IsFormRetry And (InputValue = ""), addReadOnlyWrapper)
End Function
'
'
'
Private Function getCheckoutForm_ShippingBilling_Row_Password(Label As String, InputName As String, InputValue As String, IsRequired As Boolean, IsFormRetry As Boolean, Optional addReadOnlyWrapper As Boolean) As String
    getCheckoutForm_ShippingBilling_Row_Password = getCheckoutForm_ShippingBilling_Row(Label, "<input type=""password"" name=""" & InputName & """ VALUE=""" & InputValue & """ SIZE=""30"" MAXLENGTH=""50"">", IsRequired, IsRequired And IsFormRetry And (InputValue = ""), addReadOnlyWrapper)
End Function
'
'
'
Private Function getCheckoutForm_ShippingBilling_Row(Label As String, ByVal element As String, IsRequired As Boolean, IsError As Boolean, Optional addReadOnlyWrapper As Boolean) As String
    Dim StyleField As String
    Dim StyleCaption As String
    '
    If IsError Then
        StyleField = " class=""coInputFieldError"""
        StyleCaption = " class=""coInputCaptionError"""
        'Style = "class=""coInputCaptionError""  style=""color:white;background-color:#F08080;font-weight:bold;"""
    Else
        StyleField = " class=""coInputField"""
        StyleCaption = " class=""coInputCaption"""
    End If
    If addReadOnlyWrapper Then
        element = "<div class=""coInputFieldReadOnly"" unselectable=""on"">" & element & "&nbsp;</div>"
    End If
    If IsRequired Then
        Label = "*&nbsp;" & Label
    End If
    getCheckoutForm_ShippingBilling_Row = "" _
        & "<tr><td colspan=2 style=""height:1px;""></td></tr>" _
        & "<tr>" _
        & "<td " & StyleCaption & ">" & Label & "</td>" _
        & "<td " & StyleField & ">" & element & "</SPAN></td>" _
        & "</tr>"
End Function
'
' Returns true if this paytype is being used
'
Private Function getCheckoutForm_BillingMethod(PayType, CurrentPayMethod, Optional PayDetails As String) As String
    Exit Function
    On Error GoTo ErrorTrap
    '
    Dim stream As New FastStringClass
    Dim s As String
    '
    stream.Add ("<tr><td width=""140"" valign=""top"" align=""right"">&nbsp;</td><td width=""99%"" valign=""top"">")
    stream.Add (getCheckoutForm_HTMLBillingMethod(PayType, CurrentPayMethod, PayDetails))
    ''
    'Stream.Add ("<input type=""radio"" name=""OrderPayMethod"" value=""" & PayString(PayType) & """ ")
    'If CurrentPayMethod = PayString(PayType) Then
    '    Stream.Add ("checked")
    'End If
    'Stream.Add (scSpanNormalStyle & PayString(PayType) & PayDetails & "</SPAN>")
    '
    stream.Add ("</td></tr>")
    getCheckoutForm_BillingMethod = stream.Text
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_BillingMethod"): Resume Next
    End Function
'
' Returns true if this paytype is being used
'
Private Function getCheckoutForm_HTMLBillingMethod(PayType, CurrentPayMethod, Optional PayDetails As String) As String
    On Error GoTo ErrorTrap
    '
    Dim PayStringTemp As String
    Dim s As String
    '
    PayStringTemp = PayString(PayType)
    If CurrentPayMethod = PayStringTemp Then
        s = s & "<br><input class=""scInputRadio"" type=radio name=OrderPayMethod value=""" & PayStringTemp & """ checked>"
    Else
        s = s & "<br><input class=""scInputRadio"" type=radio name=OrderPayMethod value=""" & PayStringTemp & """>"
    End If
    '
    getCheckoutForm_HTMLBillingMethod = s & scSpanNormalStyle & PayString(PayType) & PayDetails & "</SPAN>"
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_HTMLBillingMethod"): Resume Next
    End Function
'
'------------------------------------------------------------------------------------------------------------
'   Returns the paypal form
'------------------------------------------------------------------------------------------------------------
'
Private Function getCheckoutForm_Billing_PayPalButton() As String
    On Error GoTo ErrorTrap
    '
    Dim o As String
    '
    o = o & kmaStartTable(10, 0, 0) & kmaStartTableRow & kmaStartTableCell("100%", , , "center")
    If main.ServerPageSecure Then
        o = o & "<a href=""javascript:document.getElementById('PayPal').submit();""><img border=0 src=""https://www.paypal.com/en_US/i/btn/x-click-but01.gif"" alt=""Make payments with PayPal - it's fast, free and secure!"">"
    Else
        o = o & "<a href=""javascript:document.getElementById('PayPal').submit();""><img border=0 src=""http://www.paypal.com/en_US/i/btn/x-click-but01.gif"" alt=""Make payments with PayPal - it's fast, free and secure!"">"
        'O = O & "<a onClick=""alert(document.getElementById('PayPal').action);""><img border=0 src=""http://www.paypal.com/en_US/i/btn/x-click-but01.gif"" alt=""Make payments with PayPal - it's fast, free and secure!"">"
        End If
    o = o & kmaEndTableRow
    o = o & kmaStartTableRow
    o = o & kmaEndTableCell & "<td width=""100%"" align=""center"" valign=""middle"">"
    o = o & "To use PayPal, simply click on this icon. You will be forwarded to the PayPal system to complete the transaction."
    o = o & kmaEndTableCell & kmaEndTableRow & kmaEndTable
    getCheckoutForm_Billing_PayPalButton = o
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Billing_PayPalButton"): Resume Next
    End Function
'
'==============================================================================
'   getCheckoutForm_Purchase()
'==============================================================================
'
Private Function getCheckoutForm_Purchase(allowOrdering As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim copy As String
    Dim addressLine As String
    Dim echoString As String
    Dim ItemPointer As Long
    Dim UnitPrice As Double
    Dim ItemQuantity As Long
    Dim ItemCharge As Double
    Dim itemName As String
    Dim SubTotal As Double
    Dim YearPointer As Long
    Dim stream As New FastStringClass
    Dim WorkingLink As String
    Dim WorkingQS As String
    Dim RQS As String
    Dim Style As String
    Dim ccZipCode As String
    Dim ccBillName As String
    '
    stream.Add (scSpanNormalStyle)
    RQS = main.RefreshQueryString
    RQS = ModifyQueryString(RQS, rnSrcShopFormId, shopFormCheckoutPurchase)
    WorkingLink = main.ServerPage & "?"
    Style = " style=""border-bottom:1px solid #a0a0a0;padding:5px; "
    ccZipCode = main.GetStreamText("OrdercCZip")
    If ccZipCode = "" Then
        ccZipCode = orderBillZip
    End If
    ccBillName = main.GetStreamText("OrdercCName")
    If ccBillName = "" Then
        ccBillName = orderBillName
    End If
    '
    ' Header Copy
    '
    stream.Add DivBR
    stream.Add GetCommerceCopy(shopCopyNamePurchase, shopCopyDefaultPurchase)
    stream.Add DivBR
    '
    If main.OrderAuthorizeResponse <> "" Then
        Call stream.Add(main.OrderAuthorizeResponse)
        End If
    Call stream.Add(main.GetUserError)
    stream.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">")
    '
    '   Order Details
    '
    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
    '
    WorkingQS = ModifyQueryString(RQS, "off", shopFormCheckoutReview)
    WorkingLink = main.ServerPage & "?" & WorkingQS
    stream.Add ("<tr align=""center""><td width=""100%"" colspan=""2"">")
    
    stream.Add (getCheckoutForm_ItemOrderForm(True, False, allowOrdering))
    stream.Add ("</td></tr>")
    '
    '   Billing Contact
    '
    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
    stream.Add ("<tr><td colspan=""2"" class=ccCatRowHeader>Billing Contact</td></tr>")
    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
    '
    WorkingQS = ModifyQueryString(RQS, "off", shopFormCheckoutBill)
    WorkingLink = main.ServerPage & "?" & WorkingQS
    stream.Add ("<tr><td width=""100%"" colspan=""2"">")
    stream.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">")
    'Stream.Add getCheckoutForm_Purchase_Row("Name", orderBillName, Style)
    stream.Add getCheckoutForm_Purchase_Row("Name", orderBillName, Style)
    stream.Add getCheckoutForm_Purchase_Row("Email", orderBillEmail, Style)
    stream.Add getCheckoutForm_Purchase_Row("Phone", orderBillPhone, Style)
    stream.Add getCheckoutForm_Purchase_Row("Fax", orderBillFax, Style)
    stream.Add getCheckoutForm_Purchase_Row("Company", orderBillCompany, Style)
    stream.Add getCheckoutForm_Purchase_Row("Address", orderBillAddress, Style)
    If isSupportedAddress2 Then
        stream.Add getCheckoutForm_Purchase_Row("Address line 2", orderBillAddress2, Style)
    End If
    stream.Add getCheckoutForm_Purchase_Row("City", orderBillCity, Style)
    stream.Add getCheckoutForm_Purchase_Row("State", orderBillState, Style)
    stream.Add getCheckoutForm_Purchase_Row("Postal Code", orderBillZip, Style)
    stream.Add getCheckoutForm_Purchase_Row("Country", orderBillCountry, Style)
    stream.Add ("</table >")
    stream.Add ("</td></tr>")
    '
    '   Shipping Contact
    '
    If Not orderShippingRequired Then
        'If IsMemberCommerceAdmin Then
        '    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        '    stream.Add ("<tr><td colspan=""2"" class=ccCatRowHeader>Shipping</td></tr>")
        '    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        '    stream.Add ("<tr><td width=""20%"" align=left>Commerce Administrator: This order requires no shipping.</td></tr>")
        'End If
    ElseIf orderShipMethod = "" Then
        'If IsMemberCommerceAdmin Then
        '    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        '    stream.Add ("<tr><td colspan=""2"" class=ccCatRowHeader>Shipping</td></tr>")
        '    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        '    stream.Add ("<tr><td width=""20%"" align=left>Commerce Administrator: No Order Shipping method is definied, but shipping is required.</td></tr>")
        'End If
    Else
        WorkingQS = ModifyQueryString(RQS, "off", shopFormCheckoutShipping)
        WorkingLink = main.ServerPage & "?" & WorkingQS
        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        stream.Add ("<tr><td colspan=""2"" class=ccCatRowHeader>Shipping</td></tr>")
        stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
        '
        addressLine = orderShipAddress
        If isSupportedAddress2 Then
            If addressLine <> "" And orderShipAddress2 <> "" Then
                addressLine = addressLine & "<br>"
            End If
            addressLine = addressLine & orderShipAddress2
        End If
        addressLine = addressLine & "<br>" & orderShipCity & ", " & orderShipState & " " & orderShipZip & "<br>" & orderShipCountry & "<br>Phone: " & orderShipPhone
        stream.Add ("<tr><td width=""100%"" colspan=""2"">")
        stream.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">")
        stream.Add getCheckoutForm_Purchase_Row("Name", orderShipName, Style)
        stream.Add getCheckoutForm_Purchase_Row("Company", orderShipCompany, Style)
        stream.Add getCheckoutForm_Purchase_Row("Address", addressLine, Style)
        stream.Add getCheckoutForm_Purchase_Row("Method", orderShipMethod, Style)
        stream.Add ("</table >")
        stream.Add ("</td></tr>")
    End If
    '
    '   Payment Methods
    '
    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
    stream.Add ("<tr><td colspan=""2"" class=ccCatRowHeader>Payment&nbsp;</td></tr>")
    stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
    '
    WorkingQS = ModifyQueryString(RQS, "off", shopFormCheckoutBill)
    WorkingLink = main.ServerPage & "?" & WorkingQS
    stream.Add ("<tr><td colspan=""2"" width=""100%"">")
    stream.Add (main.GetFormStart & "<div align=""left"">")
    stream.Add (main.GetFormInputHidden(RequestNameRefreshBlock, main.GetFormSN))
    stream.Add ("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormCheckoutPurchase & """>")
    '
    stream.Add ("<table border=""0"" cellspacing=""10"" width=""100%"" cellpadding=""2"">")
    'stream.Add ("<tr><td align=""left"" colspan=""2"">" & scSpanNormalStyle)
    'stream.Add GetCommerceCopy(shopCopyNamePaymentDetailPrefix & OrderPayMethod, OrderPayMethod)
    'stream.Add ("</SPAN></td></tr>")
    ''
    '' Total Charge
    ''
    'stream.Add ("<tr><td width=""30%"" align=""right"" valign=""baseline"">" & scSpanNormalStyle & "Total Charge</SPAN></td><td width=""70%"">" & scSpanNormalStyle & FormatCurrency(OrdertotalCharge) & "</SPAN></td></tr>")
    'stream.Add ("<tr><td colspan=2>&nbsp;</td></tr>")
    '
    If OrderPayMethod <> PayString(PayTypeCreditCardOnline) Then
        '
        ' No credit card processing - check for paypal
        '
        If OrderPayMethod = PayString(PayTypePAYPAL) Then
            stream.Add ("<tr><td colspan=""5"" width=""100%"">")
            stream.Add (getCheckoutForm_Billing_PayPalButton)
            stream.Add ("</td></tr>")
        Else
            copy = shopCopyNamePaymentDetailPrefix & OrderPayMethod
            copy = main.GetContentCopy2(copy, "copy content", OrderPayMethod)
            stream.Add copy
        End If
    Else
        '
        ' Credit Card
        '
        stream.Add ("<tr><td width=""30%"" align=""right"" valign=""baseline"">" & scSpanNormalStyle & "Credit Card</SPAN></td><td width=""70%"">" & scSpanNormalStyle)
        stream.Add ("<select NAME=""OrdercCType"" size=""1"">")
        If OrderAllowVisa Then
            stream.Add ("<option value=""" & CCString(CCTYPEVISA) & """>" & CCString(CCTYPEVISA) & "</option>")
        End If
        If OrderAllowMC Then
            stream.Add ("<option value=""" & CCString(CCTYPEMC) & """>" & CCString(CCTYPEMC) & "</option>")
        End If
        If OrderAllowAmex Then
            stream.Add ("<option value=""" & CCString(CCTYPEAMEX) & """>" & CCString(CCTYPEAMEX) & "</option>")
        End If
        If OrderAllowDiscover Then
            stream.Add ("<option value=""" & CCString(CCTYPEDISCOVER) & """>" & CCString(CCTYPEDISCOVER) & "</option>")
        End If
        stream.Add ("</select></SPAN></td></tr>")
        stream.Add ("<tr><td width=""30%"" align=""right"" valign=""baseline"">" & scSpanNormalStyle & "Card #</SPAN></td><td width=""70%"">" & scSpanNormalStyle & "<input NAME=""OrdercCNumber"" VALUE Size=""20""></SPAN></td></tr>")
        stream.Add ("<tr><td width=""30%"" align=""right"" valign=""baseline"">" & scSpanNormalStyle & "Expires</SPAN></td><td width=""70%"">" & scSpanNormalStyle)
        stream.Add ("<select NAME=""OrdercCExpiresMonth"" size=""1"">")
        stream.Add ("<option value=""01"">01</option>" _
            & "<option value=""02"">02</option>" _
            & "<option value=""03"">03</option>" _
            & "<option value=""04"">04</option>" _
            & "<option value=""05"">05</option>" _
            & "<option value=""06"">06</option>" _
            & "<option value=""07"">07</option>" _
            & "<option value=""08"">08</option>" _
            & "<option value=""09"">09</option>" _
            & "<option value=""10"">10</option>" _
            & "<option value=""11"">11</option>" _
            & "<option value=""12"">12</option>")
        stream.Add ("</select>/<select NAME=""OrdercCExpiresYear"" size=""1"">")
        For YearPointer = Year(Now) To Year(Now) + 9
            stream.Add ("<option value=""" & Mid(CStr(YearPointer), 3) & """>" & YearPointer & "</option>")
            Next
        stream.Add ("</select>")
        stream.Add ("</SPAN></td></tr>")
        stream.Add ("<tr><td width=""30%"" align=""right"" valign=""baseline"">" & scSpanNormalStyle & "Billing Name on Card</SPAN></td><td width=""70%"">" & scSpanNormalStyle & "<input NAME=""OrdercCName"" VALUE=""" & ccBillName & """ Size=""20""></SPAN></td></tr>")
        stream.Add ("<tr><td width=""30%"" align=""right"" valign=""baseline"">" & scSpanNormalStyle & "Card Billing postal code</SPAN></td><td width=""70%"">" & scSpanNormalStyle & "<input NAME=""OrdercCZip"" VALUE=""" & ccZipCode & """ Size=""20""></SPAN></td></tr>")
        End If
    stream.Add ("</table >")
    '
    ' buttons
    '
    stream.Add ("<br>")
    If OrderPayMethod = PayString(PayTypePAYPAL) Then
        stream.Add main.GetPanelButtons(shopButtonBack, rnButton)
    Else
        stream.Add main.GetPanelButtons(shopButtonBack & "," & shopButtonOrderNow, rnButton)
    End If
    '
    stream.Add ("</FORM>")
    '
    '   no nested form in HTML - add PayPal after main form
    '
    echoString = CStr(main.visitId) & "," & CStr(OrdertotalCharge)
    stream.Add (getCheckoutForm_Billing_PayPalForm(main, echoString, OrdertotalCharge, "Order " & orderId))
    stream.Add ("</td></tr>")
    stream.Add ("</table >")
    getCheckoutForm_Purchase = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Purchase"): Resume Next
    End Function
'
'
'
Private Function getCheckoutForm_Purchase_Row(RowCaption As String, RowValue As String, Style As String)
    On Error GoTo ErrorTrap
    '
    Dim RowValueBuffered As String
    '
    RowValueBuffered = RowValue
    If RowValue = "" Then
        RowValueBuffered = "&nbsp;"
    End If
    '
    getCheckoutForm_Purchase_Row = "<tr><td width=""150"" align=""right""><strong>" & scSpanNormalStyle & RowCaption & "</SPAN></strong></td><td " & Style & " width=""99%"">" & scSpanNormalStyle & RowValueBuffered & "</SPAN></td></tr>"
    'getCheckoutForm_Purchase_Row = "<tr><td width=""20%"" align=""right""><strong>" & scSpanNormalStyle & RowCaption & "</SPAN></strong></td><td " & Style & " width=""80%"">" & scSpanNormalStyle & RowValueBuffered & "</SPAN></td></tr>"
    Exit Function
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Purchase_Row"): Resume Next
End Function
'
'==============================================================================
'   getCheckoutForm_Receipt()
'       If there is an account with a completed order, print the last
'       If no account with completed order, print error form
'==============================================================================
'
Private Function getCheckoutForm_Receipt() As String
    On Error GoTo ErrorTrap
    '
    Dim itemId As Long
    Dim c As String
    Dim location As String
    Dim Caption As String
    Dim OrderPriceLevelDiscountPrice As Double
    Dim OrderPriceLevelDiscountDiscount As Double
    Dim SQL As String
    Dim CSOrders As Long
    Dim CSOrderDetails As Long
    Dim stream As New FastStringClass
    Dim Link As String
    Dim qs As String
    Dim PayMethod As String
    Dim csItem As Long
    Dim customCopy As String
    '
    ' Header Copy
    '
    stream.Add DivBR
    stream.Add GetCommerceCopy(shopCopyNameReceipt, shopCopyDefaultReceipt)
    stream.Add DivBR
    '
    ' lookup the last complete order on the this account
    '
    SQL = "select * from orders where ((active<>0)and(DateCompleted is not null)AND(MemberID=" & CommerceAccountMemberID & ")) order by DateCompleted DESC;"
    CSOrders = main.OpenCSSQL(CommerceDataSourceName, SQL)
    If Not main.IsCSOK(CSOrders) Then
        Call HandleClassErrorAndResumeNext("getCheckoutForm_Receipt", "Error opening order by MemberID, sql = " & SQL)
        stream.Add getCheckoutForm_Error
    Else
        If True Then
            CSOrderDetails = main.OpenCSContent_Internal("Order Details", "(OrderID=" & KmaEncodeSQLNumber(main.GetCSText(CSOrders, "ID")) & ")")
            If main.IsCSOK(CSOrderDetails) Then
                '
                ' Print the receipt form
                '
                'If main.OrderAuthorizeResponse <> "" Then
                '    stream.Add (scSpanNormalStyle & main.OrderAuthorizeResponse & "<P>")
                'Else
                '    stream.Add (scSpanNormalStyle & "<p>Thank you for your order.</p><p>We appreciate your business. Please print this page as a record of your purchase.</p>")
                '    End If
                '
                stream.Add ("<table class=""scReceipt"" border=""1"" cellpadding=""8"" cellspacing=""0"" width=""100%"">")
                stream.Add ("<tr>")
                stream.Add ("<td width=""100%"" colspan=""2"" >" & scSpanNormalStyle & "<B>Order Number </B>" & main.GetCSText(CSOrders, "ID"))
                stream.Add ("</tr>")
                '
                ' Print Billing/Shipping Information
                '
                stream.Add ("<tr>")
                stream.Add ("<td width=""50%"" align=""left"" valign=""top"">" & scSpanNormalStyle)
                stream.Add ("<B>Ship To</B><br>")
                c = main.GetCSText(CSOrders, "ShipName")
                If c <> "" Then
                    stream.Add (c & "<br>")
                End If
                c = main.GetCSText(CSOrders, "ShipCompany")
                If c <> "" Then
                    stream.Add (c & "<br>")
                End If
                c = main.GetCSText(CSOrders, "ShipAddress")
                If c <> "" Then
                    stream.Add (c & "<br>")
                End If
                If isSupportedAddress2 Then
                    c = main.GetCSText(CSOrders, "ShipAddress2")
                    If c <> "" Then
                        stream.Add (c & "<br>")
                    End If
                End If
                location = ""
                c = main.GetCSText(CSOrders, "ShipCity")
                If c <> "" Then
                    location = c
                End If
                c = main.GetCSText(CSOrders, "ShipState")
                If c <> "" Then
                    If location <> "" Then
                        location = location & ", "
                    End If
                    location = location & c
                End If
                c = main.GetCSText(CSOrders, "ShipZip")
                If c <> "" Then
                    If location <> "" Then
                        location = location & " "
                    End If
                    location = location & c
                End If
                If location <> "" Then
                    stream.Add (location & "<br>")
                End If
                c = main.GetCSText(CSOrders, "ShipCountry")
                If c <> "" Then
                    stream.Add (c & "<br>")
                End If
                c = main.GetCSText(CSOrders, "ShipPhone")
                If c <> "" Then
                    stream.Add ("Phone: " & c & "<br>")
                End If
                stream.Add ("</SPAN></td>")
                stream.Add ("<td width=""50%"" align=""left"" valign=""top"">" & scSpanNormalStyle)
                stream.Add ("<B>Billing</B><br>")
                stream.Add (main.GetCSText(CSOrders, "BillName") & "<br>")
                stream.Add ("Email: " & main.GetCSText(CSOrders, "BillEmail") & "<br>")
                stream.Add ("Phone: " & main.GetCSText(CSOrders, "BillPhone") & "<br>")
                stream.Add ("</SPAN></td>")
                stream.Add ("</tr>")
                '
                ' Shipping Method
                '
                stream.Add ("<tr><td colspan=""2"">" & scSpanNormalStyle & "<B>Ship By </B>" & main.GetCSText(CSOrders, "ShipMethod") & "</td></tr>")
                '
                ' Itemize the purchase
                '
                stream.Add ("<tr><td colspan=""2"">")
                stream.Add ("<table border=""0"" cellpadding=""4"" cellspacing=""2"" width=""100%"" >")
                stream.Add ("<tr>")
                stream.Add ("<td align=""center"" width=""50"">" & scSpanNormalStyle & "<B>Qty</B></SPAN><br><img alt="""" src=""/ccLib/Images/spacer.gif"" width=""50"" height=""1""></td>")
                stream.Add ("<td align=""left"" width=""99%"">" & scSpanNormalStyle & "<B>Description</B></SPAN><br><img alt="""" src=""/ccLib/Images/spacer.gif"" width=""100%"" height=""1""></td>")
                stream.Add ("<td align=""right"" width=""70"">" & scSpanNormalStyle & "<B>Total</B></SPAN><br><img alt="""" src=""/ccLib/Images/spacer.gif"" width=""70"" height=""1""></td>")
                stream.Add ("</tr>")
                Do While main.IsCSOK(CSOrderDetails)
                    stream.Add ("<tr>")
                    stream.Add ("<td align=""center"" class=""ccPanelRowEven"">" & scSpanNormalStyle & main.GetCSText(CSOrderDetails, "quantity") & "</SPAN></td>")
                    Caption = main.GetCSText(CSOrderDetails, "name")
                    If main.GetCSText(CSOrderDetails, "options") <> "" Then
                        Caption = Caption & ", " & main.GetCSText(CSOrderDetails, "options")
                    End If

                    stream.Add ("<td align=""left"" class=""ccPanelRowEven"">" & scSpanNormalStyle & Caption & "</SPAN></td>")
                    stream.Add ("<td align=""right"" class=""ccPanelRowEven"">" & scSpanNormalStyle & FormatNumber(main.GetCSNumber(CSOrderDetails, "unitprice") * main.GetCSNumber(CSOrderDetails, "quantity"), 2) & "</SPAN></td>")
                    stream.Add ("</tr>")
                    itemId = main.GetCSInteger(CSOrderDetails, "itemId")
                    csItem = main.OpenCSContentRecord("items", itemId)
                    If main.IsCSOK(csItem) Then
                        If main.GetCSBoolean(csItem, "allowReceiptCopy") Then
                            customCopy = customCopy & main.GetCSText(csItem, "receiptCopy")
                        End If
                    End If
                    Call main.CloseCS(csItem)
                    main.NextCSRecord (CSOrderDetails)
                Loop
                stream.Add ("<tr>")
                stream.Add ("<td align=""right"" colspan=""2"">" & scSpanNormalStyle & "<B>Subtotal</B></SPAN></td>")
                stream.Add ("<td align=""right"" class=""ccPanelRowEven"">" & scSpanNormalStyle & FormatNumber(main.GetCSNumber(CSOrders, "itemCharge"), 2) & "</SPAN></td>")
                stream.Add ("</tr>")
                If main.GetCSNumber(CSOrders, "PriceLevelDiscount") <> 0 Then
                    OrderPriceLevelDiscountPrice = KmaEncodeNumber(main.GetSiteProperty("OrderPriceLevelDiscountPrice", "0"))
                    OrderPriceLevelDiscountDiscount = KmaEncodeNumber(main.GetSiteProperty("OrderPriceLevelDiscountDiscount", "0"))
                    stream.Add ("<tr>")
                    stream.Add ("<td align=""right"" colspan=""2"">" & scSpanNormalStyle & "<B>" & OrderPriceLevelDiscountDiscount & "% Discount for orders over $" & OrderPriceLevelDiscountPrice & "</B></SPAN></td>")
                    stream.Add ("<td align=""right"" class=""ccPanelRowEven"">" & scSpanNormalStyle & FormatNumber(main.GetCSNumber(CSOrders, "PriceLevelDiscount"), 2) & "</SPAN></td>")
                    stream.Add ("</tr>")
                End If
                stream.Add ("<tr>")
                stream.Add ("<td align=""right"" colspan=""2"">" & scSpanNormalStyle & "<B>Sales Tax</B></SPAN></td>")
                stream.Add ("<td align=""right"" class=""ccPanelRowEven"">" & scSpanNormalStyle & FormatNumber(main.GetCSNumber(CSOrders, "TaxCharge"), 2) & "</SPAN></td>")
                stream.Add ("</tr>")
                stream.Add ("<tr>")
                stream.Add ("<td align=""right"" colspan=""2"">" & scSpanNormalStyle & "<B>Shipping</B></SPAN></td>")
                stream.Add ("<td align=""right"" class=""ccPanelRowEven"">" & scSpanNormalStyle & FormatNumber(main.GetCSNumber(CSOrders, "ShipCharge"), 2) & "</SPAN></td>")
                stream.Add ("</tr>")
                stream.Add ("<tr>")
                stream.Add ("<td align=""right"" colspan=""2"">" & scSpanNormalStyle & "<B>Total Charge</B></SPAN></td>")
                stream.Add ("<td align=""right"" class=""ccPanelRowEven"">" & scSpanNormalStyle & FormatNumber(main.GetCSNumber(CSOrders, "TotalCharge"), 2) & "</SPAN></td>")
                stream.Add ("</tr></table >")
                stream.Add ("</td></tr>")
                '
                PayMethod = main.GetCSText(CSOrders, "PayMethod")
                stream.Add ("<tr><td width=""100%"" colspan=""2"" >" & scSpanNormalStyle & "<B>Payment Method</B> " & GetCommerceCopy(shopCopyNameReceptDetailPrefix & PayMethod, shopCopyDefaultReceptDetailPrefix & PayMethod) & "</td></tr>")
                '
                ' custom instructions
                '
                If customCopy <> "" Then
                    stream.Add ("<tr><td width=""100%"" colspan=""2"" ><B>Special Instructions</B>" & customCopy & "</td></tr>")
                End If
                stream.Add ("</table >")
                'Stream.Add ("<P>Return to <A href=""" & GetShoppingCartLink(Main.ServerPage, Main.RefreshQueryString) & """>Catalog</A></p>")
                stream.Add ("</SPAN>")
            End If
            Call main.CloseCS(CSOrderDetails)
        End If
    End If
    Call main.CloseCS(CSOrders)
    '
    getCheckoutForm_Receipt = stream.Text
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Receipt"): Resume Next
    End Function
'
'==============================================================================
' getCheckoutForm_Error()
'   print a blocking form if the there was a problem
'==============================================================================
'
Private Function getCheckoutForm_Error() As String
    On Error GoTo ErrorTrap
    '
    Dim stream As New FastStringClass
    '
    stream.Add (scSpanNormalStyle & "<strong>")
    stream.Add ("There has been an error accessing your order.")
    stream.Add ("</strong></SPAN>")
    stream.Add ("<P>" & scSpanNormalStyle & "The order may have been completed, or too much time may have elapsed since you last accessed the order.</SPAN></p>")
    stream.Add ("<P>" & scSpanNormalStyle & "")
    stream.Add ("If the order was completed and you need to make a change, please call Kidwell McGowan Associates at 703 406-3655, Monday through Friday, 9:00am to 5:00pm EST, or email <A href=""mailto:orders@Kma.net"">orders@Kma.net</A>.")
    stream.Add ("Include a description of your change and a phone number where you can be reached.</SPAN></p>")
    stream.Add ("<P>" & scSpanNormalStyle & "If too much time has elapsed since you last accessed the order, return to the site using the link below, and your items will still be in your shopping cart.</SPAN></p>")
    'Stream.Add ("<P>" & scSpanNormalStyle & "Return to <A href=""" & GetShoppingCartLink(Main.ServerPage, Main.RefreshQueryString) & """>Catalog</A></SPAN>")
    getCheckoutForm_Error = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Error"): Resume Next
    End Function
'
'
'
Private Function GetEditFormRow(SubheadRow As Boolean, Caption As String, ByVal element As String, Optional addReadOnlyWrapper As Boolean) As String
    On Error GoTo ErrorTrap
    '
    If SubheadRow Then
        GetEditFormRow = "<tr><td colspan=""2"" align=""Left"" class=ccCatRowHeader>" & Caption & "</td></tr>"
    Else
        If addReadOnlyWrapper Then
            element = "<div class=""coInputFieldReadOnly"" unselectable=""on"">" & element & "&nbsp;</div>"
        End If
        GetEditFormRow = "" _
            & "<tr>" _
            & "<td class=""coInputCaption""><nobr>" & Caption & "</nobr></td>" _
            & "<td class=""coInputField"">" & element & "</td>" _
            & "</tr>"
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetEditFormRow")
    End Function

'
'
'
Private Function GetEditFormRowCS(CSPointer As Long, FieldName As String, Optional fieldType As Integer) As String
    On Error GoTo ErrorTrap
    '
    Dim copy As String
    Dim number As Double
    '
    Select Case fieldType
        Case FieldTypeCurrency
            copy = "$&nbsp;" & FormatNumber(main.GetCSNumber(CSPointer, FieldName), 2)
        Case Else
            copy = main.GetCSText(CSPointer, FieldName)
    End Select
    copy = "<div class=""coInputFieldReadOnly "">" & copy & "&nbsp;</div>"
    GetEditFormRowCS = GetEditFormRow(False, main.GetCSFieldCaption(CSPointer, FieldName), copy)
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetEditFormRowCS")
    End Function
'
'
'
Private Function getCheckoutForm(dstShopFormId, allowOrdering As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim stream As New FastStringClass
    Dim common As New commonClass
    Dim testid As Long
    Dim forceLogin As Boolean
    '
    Call main.AddPageTitle("Checkout")
    '
    ' verify United States is in the list - for UPS shipping defaults
    '
    testid = main.GetRecordID("countries", "united states")
    If testid = 0 Then
        'Set common = New commonClass
        Call common.VerifyShippingStateCountry(main)
    End If
    '
    ' If order is empty, allow receipt and review only
    '
    If (dstShopFormId <> shopFormCheckoutReview) And (dstShopFormId <> shopFormCheckoutReceipt) And (Not main.IsOrderOK()) Then
        dstShopFormId = shopFormCheckoutReview
    End If
    '
    ' select the correct form
    '
    Call main.AddRefreshQueryString(rnDstShopFormId, dstShopFormId)
    '
    ' intercept shipping billing and force login if not -- but only after the correct for is set in refresh
    '
    forceLogin = False
    If (dstShopFormId <> 0) And (dstShopFormId <> 21) Then
        If Not main.IsAuthenticated Then
            forceLogin = True
            'dstShopFormId = shopFormCheckoutAccount
        End If
    End If
    If forceLogin Then
        stream.Add getCommonForm_Login(shopFormCheckoutAccount, CInt(dstShopFormId))
    Else
    Select Case dstShopFormId
        Case shopFormCheckoutAccount
            '
            Call main.AddRefreshQueryString(rnDstShopFormId, shopFormCheckoutShippingBilling)
            stream.Add getCommonForm_Login(shopFormCheckoutAccount, shopFormCheckoutAccount)
            '
        Case shopFormCheckoutShippingBilling
            stream.Add getCheckoutForm_ShippingBilling
            '
        Case shopFormCheckoutReceipt
            stream.Add getCheckoutForm_Receipt
            '
        Case shopFormCheckoutLogin
            stream.Add getCheckoutForm_Login
            '
        Case shopFormCheckoutPurchase
            stream.Add getCheckoutForm_Purchase(allowOrdering)
            '
        Case Else
            stream.Add getCheckoutForm_Review(allowOrdering)
    End Select
    End If
    '
    'Call Main.AddRefreshQueryString("off", dstShopFormId)
    'Call Main.AddRefreshQueryString(rnSrcShopFormId, dstShopFormId)
    getCheckoutForm = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForms"): Resume Next
    End Function
'
'
'
Private Sub Class_Initialize()
    ClassInitialized = False
    End Sub
'
'
'
Private Sub Class_Terminate()
    Set main = Nothing
    End Sub

'
'
'
Private Sub LoadProperties()
    OrderAllowVisa = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowVisa", False))
    OrderAllowMC = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowMC", False))
    OrderAllowDiscover = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowDiscover", False))
    OrderAllowAmex = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowAmex", False))
    OrderAllowCreditCardOnline = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCreditCardOnline", False))
    OrderAllowNetTerms = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowNetTerms", False))
    OrderAllowCODCompanyCheck = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCODCompanyCheck", False))
    OrderAllowCODCertifiedFunds = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCODCertifiedFunds", False))
    OrderAllowCreditCardOnline = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCreditCardOnline", False))
    OrderAllowCreditCardByPhone = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCreditCardByPhone", False))
    OrderAllowCreditCardByFax = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCreditCardByFax", False))
    OrderAllowPersonalChecks = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowPersonalChecks", False))
    OrderAllowCompanyChecks = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCompanyChecks", False))
    OrderAllowCredit = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowCredit", False))
    OrderAllowPayPal = kmaEncodeBoolean(main.GetSiteProperty("OrderAllowPayPal", False))
    End Sub
'
'
'
Private Function GetCommerceCopy(PageName As String, defaultCopy As String)
    On Error GoTo ErrorTrap
    '
    'Dim CS As Long
    'Dim RecordID As Long
    Dim copy As New CmrcCopyClass
    '
    GetCommerceCopy = copy.GetCopy(main, PageName, defaultCopy)
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCommerceCopy"): Resume Next
End Function
'
'
'
Public Function GetOrderProcess()
    Dim srcShopFormId As Long
    Dim dstShopFormId As Long
    '
    srcShopFormId = main.GetStreamInteger(rnSrcShopFormId)
    dstShopFormId = main.GetStreamInteger(rnDstShopFormId)
    '
    GetOrderProcess = getCheckout(srcShopFormId, dstShopFormId)
End Function
'
Public Function getCheckout(srcShopFormId As Long, dstShopFormId As Long)
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim allowOrdering As Boolean
    Dim common As New commonClass
    '
    ' Verify that the Order's Member matches the CommerceAccountMemberID, else updated the order
    ' Can not be in Init, because calling routine sets CommerceAccountMemberID after Init returns
    '
    Call LoadOrderGlobals(False)
    '
    allowOrdering = common.getAllowOrdering(main, CommerceAccountMemberID)
    '
    ' Verify Credentials
    '   You must be either logged in, or on a guest account without username or email
    '
    If (Not main.IsAuthenticated()) And ((main.MemberEmail <> "") Or (main.MemberUsername <> "")) Then
        Call main.LogoutMember
    End If
    '
    If srcShopFormId <> 0 Then
        dstShopFormId = processCheckoutForm(srcShopFormId)
    End If
    '
    ' Assemble page
    '
    getCheckout = "" _
            & vbCrLf & vbTab & "<div class=""ccCatBody"">" _
            & kmaIndent(getCheckoutForm(dstShopFormId, allowOrdering)) _
            & vbCrLf & vbTab & "</div>" _
'    If IsPartOfShoppingCart Then
'        getCheckout = "" _
'            & "<table border=0 cellpadding=0 cellspacing=0 width=100%><tr><td class=ccCatBody>" & getCheckoutForm(dstShopFormId, allowOrdering) & "</td></tr></table >"
'    Else
'        getCheckout = "" _
'            & "<table border=0 cellpadding=0 cellspacing=0 width=100%><tr><td class=ccCatBody>" & getCheckoutForm(dstShopFormId, allowOrdering) & "</td></tr></table >"
'    End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckout"): Resume Next
End Function
'
'
'
Public Function VerifyOrderMember()
    Call HandleClassErrorAndResumeNext("VerifyOrderMember", "VerifyOrderMember is a deprecated method. Integrety checked during load."): Resume Next
End Function
'
'
'
Private Function GetFormAction() As String
    Dim qs As String
    Dim FormAction As String
    '
    qs = main.RefreshQueryString
    GetFormAction = main.ServerFormActionURL
    If qs <> "" Then
        GetFormAction = GetFormAction & "?" & qs
    End If
End Function
'
'
'
Private Function GetCurrentLink(SecureProtocol As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim qs As String
    '
    qs = main.RefreshQueryString
    If qs <> "" Then
        qs = "?" & qs
    End If
    If main.ServerPageSecure And SecureProtocol Then
        '
        ' Do not change protocol
        '
        GetCurrentLink = main.ServerPage & qs
    ElseIf SecureProtocol Then
        '
        ' Change to Secure
        '
        GetCurrentLink = "https://" & main.ServerHost & main.ServerAppRootPath & main.ServerAppPath & main.ServerPage & qs
    Else
        '
        ' Change to non-secure
        '
        GetCurrentLink = "http://" & main.ServerHost & main.ServerAppRootPath & main.ServerAppPath & main.ServerPage & qs
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCurrentLink"): Resume Next
End Function
'
'
'
'Private Function GetShippingSelect(CurrentName As String, orderShipWeight As Double, OrderItemCharge As Double)
'    On Error GoTo ErrorTrap
'    '
'    Dim CS As String
'    Dim Name As String
'    Dim ChargeValue As Double
'    Dim Time As String
'    Dim Label As String
'    Dim Domestic As Boolean
'    Dim International As Boolean
'    Dim ShippingDaysMin As Long
'    Dim ShippingDaysMax As Long
'    Dim AllowPriceLevelFreeShipping As Boolean
'    Dim SiteAllowPriceLevelFreeShipping As Boolean
'    Dim SiteFreeShippingPrice As Double
'    Dim upsVal As String
'    Dim appURL As String
'    Dim objXML As New XMLHTTP60
'    '
'    CS = Main.OpenCSContent("Order Ship Methods", "(name<>'') and (name is not null)")
'    If Main.IsCSOK(CS) Then
'        SiteAllowPriceLevelFreeShipping = (Main.SiteProperty_BuildVersion > "3.3.580") And kmaEncodeBoolean(Main.GetSiteProperty("OrderAllowFreeShippingPrice", "0"))
'        SiteFreeShippingPrice = KmaEncodeNumber(Main.GetSiteProperty("OrderFreeShippingPrice", "0"))
'        Do While Main.IsCSOK(CS)
'            Label = ""
'            If SiteAllowPriceLevelFreeShipping Then
'                AllowPriceLevelFreeShipping = Main.GetCSInteger(CS, "AllowPriceLevelFreeShipping")
'            End If
'            Name = Main.GetCSText(CS, "Name")
'            ChargeValue = CalcShippingCharge(Main.GetCSNumber(CS, "BaseCharge"), Main.GetCSNumber(CS, "ChargePerWeight"), Main.GetCSNumber(CS, "ChargePerCharge"), orderShipWeight, OrderItemCharge, AllowPriceLevelFreeShipping, SiteFreeShippingPrice)
'            '
'            If ChargeValue < 0.01 Then
'                Label = "FREE SHIPPING"
'            Else
'                Label = "est " & FormatCurrency(CStr(ChargeValue))
'            End If
'            '
'            ShippingDaysMin = Main.GetCSInteger(CS, "DaysMin")
'            ShippingDaysMax = Main.GetCSInteger(CS, "DaysMax")
'            If (ShippingDaysMin > 0) Or (ShippingDaysMax > 0) Then
'                Label = Label & ", " & ShippingDaysMin
'                If ShippingDaysMin = ShippingDaysMax Then
'                    If ShippingDaysMin = 1 Then
'                        Label = Label & " business day"
'                    Else
'                        Label = Label & " business days"
'                    End If
'                Else
'                    Label = Label & " to " & ShippingDaysMax & " business days"
'                End If
'            End If
'            '
'            Domestic = Main.GetCSBoolean(CS, "Domestic")
'            International = Main.GetCSBoolean(CS, "International")
'            If Domestic And International Then
'                Label = Label & ", Domestic and International"
'            ElseIf Domestic Then
'                Label = Label & ", Domestic only"
'            ElseIf Domestic Then
'                Label = Label & ", International only"
'            End If
'            '
'            Label = Name & " (" & Label & ")"
'            '
'            If Name = CurrentName Then
'                GetShippingSelect = GetShippingSelect & "<option value=""" & Name & """ selected>" & Label & "</option>"
'            Else
'                GetShippingSelect = GetShippingSelect & "<option value=""" & Name & """>" & Label & "</option>"
'            End If
'            '
'            Main.NextCSRecord (CS)
'        Loop
'    End If
'    Call Main.CloseCS(CS)
'
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Next Day Air Early AM", "AllowUPSNextDatAirAM", "1DM", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Next Day Air", "AllowUPSNextDatAir", "1DA", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Next Day Air Saver", "AllowUPSNextDatAirSaver", "1DP", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS 2nd Day Air AM", "AllowUPS2ndDayAirAM", "2DM", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS 2nd Day Air", "AllowUPS2ndDayAir", "2DA", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS 3 Day Select", "AllowUPS3DaySelect", "3DS", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Ground", "AllowUPSGround", "GND", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Canada Standard", "AllowUPSCanadaStandard", "STD", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Worldwide Express", "AllowUPSCWorldwideExpress", "XPR", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Worldwide Express Plus", "AllowUPSCWorldwideExpressPlus", "XDM", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Worldwide Expedited", "AllowUPSCWorldwideExpedited", "XPD", CurrentName)
'    GetShippingSelect = GetShippingSelect & getShipOptionUPS("UPS Worldwide Express Saver", "AllowUPSCWorldwideSaver", "WXS", CurrentName)
'
'    If GetShippingSelect <> "" Then
'        GetShippingSelect = "<select name=""orderShipMethod""><option value="""" selected>Select One</option>" & GetShippingSelect & "</select>"
'    End If
'    '
'    '
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("GetShippingSelect"): Resume Next
'End Function
'
'
'
Private Function getCheckoutForm_ItemOrderForm(IncludeTotals As Boolean, AllowChanges As Boolean, allowOrdering As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim OrderPriceLevelDiscountPrice As Double
    Dim OrderPriceLevelDiscountDiscount As Double
    Dim cs As String
    Dim s As New FastStringClass
    Dim ItemPtr As Long
    Dim SubTotal As Double
    Dim copy As String
    Dim Style As String
    Dim StyleRight As String
    Dim StyleLeft As String
    Dim ItemSelect As String
    Dim QuanityString As String
    Dim Caption As String
    Dim qs As String
    Dim CSSelect As Long
    Dim ItemSelectOptions As String
    Dim AddButton As String
    '
    Const BakeName = "ItemSelect"
        s.Add ("<table ID=""scCheckoutTable"" class=""scCheckoutTable"" border=0 cellpadding=2 cellspacing=0 width=100%>")
        s.Add ("<tr>" _
            & "<td class=ccCatRowHeader style=""" & Style & """ align=center width=40>Qty.<br><img src=/cclib/images/spacer.gif width=40 height=1></td>" _
            & "<td class=ccCatRowHeader style=""" & Style & """ align=left width=99%>Product Description</td>" _
            & "<td class=ccCatRowHeader style=""" & Style & """ align=right width=70>Each<br><img src=/cclib/images/spacer.gif width=70 height=1></td>" _
            & "<td class=ccCatRowHeader style=""" & Style & """ align=right width=70>Total<br><img src=/cclib/images/spacer.gif width=70 height=1></td>" _
            & "</tr>")
        If main.OrderItemCount > 0 Then
            'Style = "color:black;background-color:white;border-bottom: 1px solid #a0a0a0;padding-left:5px;padding-right:5px;padding-top:10px;border-right: 1px solid #a0a0a0;"
            'StyleLeft = "color:black;background-color:white;border-bottom: 1px solid #a0a0a0;padding-left:5px;padding-right:5px;border-right: 1px solid #a0a0a0;"
            'StyleRight = "color:black;background-color:white;border-bottom: 1px solid #a0a0a0;padding-left:5px;padding-right:5px;padding-top:10px;"
            For ItemPtr = 0 To main.OrderItemCount - 1
                SubTotal = OrderDetails(ItemPtr).Price * OrderDetails(ItemPtr).quantity
                s.Add ("<tr>")
                If AllowChanges Then
                    QuanityString = "" _
                        & "<input TYPE=TEXT NAME=Q" & CStr(ItemPtr) & " VALUE=" & OrderDetails(ItemPtr).quantity & "  SIZE=2 MAXLENGTH=5>" _
                        & "<input type=hidden name=ID" & CStr(ItemPtr) & " value=" & OrderDetails(ItemPtr).itemId & ">" _
                        & "<input type=hidden name=D" & CStr(ItemPtr) & " value=" & OrderDetails(ItemPtr).Id & ">"
                Else
                   QuanityString = OrderDetails(ItemPtr).quantity
                End If
                Caption = OrderDetails(ItemPtr).itemName
                If OrderDetails(ItemPtr).Options <> "" Then
                    Caption = Caption & ", " & OrderDetails(ItemPtr).Options
                End If
                If IsPartOfShoppingCart Then
                    qs = main.RefreshQueryString
                    qs = ModifyQueryString(qs, RequestNameCatalogItemID, CStr(OrderDetails(ItemPtr).itemId))
                    qs = ModifyQueryString(qs, rnSrcShopFormId, shopFormCatalogDetails)
                    Caption = "<a href=""?" & qs & """>" & Caption & "</a>"
                End If
                s.Add ("<td class=""scTableRowLeft"" align=center valign=baseline>" & scSpanNormalStyle & QuanityString & "</SPAN></td>")
                s.Add ("<td class=""scTableRow"" align=left valign=baseline>" & scSpanNormalStyle & Caption & "</SPAN></td>")
                s.Add ("<td class=""scTableRow"" align=right valign=baseline>" & scSpanNormalStyle & FormatCurrency(OrderDetails(ItemPtr).Price) & "</SPAN></td>")
                s.Add ("<td class=""scTableRowRight"" align=right valign=baseline>" & scSpanNormalStyle & FormatCurrency(SubTotal) & "</SPAN></td>")
                s.Add ("</tr>")
            Next
        End If
        s.Add "</table >"
        s.Add main.GetFormInputHidden("CartRows", main.OrderItemCount, "CartRows")
        '
Dim OptionName As String
        '
        ' Build Add Button
        '
        If allowOrdering Then
            AddButton = ""
            If AllowChanges Then
                ItemSelect = main.ReadBake(BakeName)
                If ItemSelect = "" Then
                    ItemSelectOptions = ""
                    '
                    '   Options long text field creating error due to sort
                    '                    HideItemOrdering = main.GetCSBoolean(CSPointer, "HideOrderButton")

                    CSSelect = main.OpenCSContent("Items", "((HideOrderButton=0)or(HideOrderButton is null))", "name", , , , "Name,Options,ID")
                    Do While main.IsCSOK(CSSelect)
                        OptionName = main.GetCS(CSSelect, "Name") & " " & main.GetCS(CSSelect, "Options")
                        OptionName = Replace(OptionName, "'", "\'")
                        OptionName = Replace(OptionName, vbCrLf, " ")
                        OptionName = Replace(OptionName, vbCr, " ")
                        OptionName = Replace(OptionName, vbLf, " ")
                        ItemSelectOptions = ItemSelectOptions & "<option value=" & main.GetCSInteger(CSSelect, "ID") & ">" & OptionName & "</option>"
                        Call main.NextCSRecord(CSSelect)
                    Loop
                    Call main.CloseCS(CSSelect)
                    '
                    ItemSelect = "<SCRIPT>var ItemSelect='<SELECT name=row0id>" & ItemSelectOptions & "</Select>';</SCRIPT>"
                    Call main.SaveBake(BakeName, ItemSelect, "Items")
                End If
                AddButton = ""
                'AddButton = "" _
                '    & ItemSelect _
                '    & "<a href=""#"" onClick=""cartInsertRow(); return false;"" style=""text-decoration: none;"">Add Another Item to Cart</a>"
            End If
        End If
        '
        ' Subtotal
        '
        Style = "color:black;background-color:white;border-top:1px solid #a0a0a0;border-bottom:1px solid #a0a0a0;padding-left:5px;padding-right:5px;padding-top:10px;"
        s.Add ("<table class=""scCheckoutTable"" border=0 cellpadding=2 cellspacing=0 width=100%>")
        s.Add ("<tr>" _
            & "<td width=70 align=left style=""color:black;background-color:white;padding-top:10px;"">" & AddButton & "<br><img src=/cclib/images/spacer.gif width=70 height=1></td>" _
            & "<td width=99% align=right colspan=2 style=""color:black;background-color:white;padding-top:10px;"">" & scSpanNormalStyle & "<B>Product Total</B>&nbsp;</SPAN></td>" _
            & "<td width=70 align=right style=""" & Style & """>" & scSpanNormalStyle & FormatCurrency(OrderItemCharge) & "<br><img src=/cclib/images/spacer.gif width=70 height=1></SPAN></td>" _
            & "</tr>")
            '& "<td width=70 align=right class=ccPanelInput style=""padding-left:5px;padding-right:5px;"">" & scSpanNormalStyle & FormatCurrency(OrderItemCharge) & "<br><img src=/cclib/images/spacer.gif width=70 height=1></SPAN></td>"
        '
        ' Price Level Discount
        '
        If main.SiteProperty_BuildVersion > "3.3.580" Then
            If kmaEncodeBoolean(main.GetSiteProperty("OrderAllowPriceLevelDiscount", False)) Then
                OrderPriceLevelDiscountPrice = KmaEncodeNumber(main.GetSiteProperty("OrderPriceLevelDiscountPrice", "0"))
                OrderPriceLevelDiscountDiscount = KmaEncodeNumber(main.GetSiteProperty("OrderPriceLevelDiscountDiscount", "0"))
                s.Add ("<tr>" _
                    & "<td width=99% align=right colspan=3 style=""color:black;background-color:white;padding-top:10px;"">" & scSpanNormalStyle & "<B>" & OrderPriceLevelDiscountDiscount & "% Discount for orders over $" & OrderPriceLevelDiscountPrice & "</B>&nbsp;</SPAN></td>" _
                    & "<td width=70 align=right style=""" & Style & """>" & scSpanNormalStyle & FormatCurrency(OrderPriceLevelDiscount, , , vbFalse) & "<br><img src=/cclib/images/spacer.gif width=70 height=1></SPAN></td>" _
                    & "</tr>")
            End If
        End If
        If IncludeTotals Then
            '
            ' Tax
            '
            Style = "color:black;background-color:white;border-bottom:1px solid #a0a0a0;padding-left:5px;padding-right:5px;padding-top:10px;"
            s.Add ("<tr>" _
                & "<td style=""color:black;background-color:white;padding-left:5px;padding-right:5px;padding-top:10px;"" colspan=2>&nbsp;</td>" _
                & "<td style=""color:black;background-color:white;padding-top:10px;"" align=""right"">" & scSpanNormalStyle & "<B>Sales Tax</B>&nbsp;</SPAN></td>" _
                & "<td width=70 align=right style=""" & Style & """>" & scSpanNormalStyle & FormatCurrency(OrderTaxCharge) & "</SPAN></td>" _
                & "</tr>")
            '
            ' Shipping
            '
            If orderShipCharge > 0 Then
                copy = FormatCurrency(orderShipCharge)
            Else
                copy = "FREE"
            End If
            Style = "color:black;background-color:white;border-bottom:1px solid #a0a0a0;padding-left:5px;padding-right:5px;padding-top:10px;"
            s.Add ("<tr>" _
                & "<td style=""color:black;background-color:white;padding-left:5px;padding-right:5px;padding-top:10px;"" colspan=2>&nbsp;</td>" _
                & "<td style=""color:black;background-color:white;padding-top:10px;"" align=""right"">" & scSpanNormalStyle & "<B>Shipping</B>&nbsp;</SPAN></td>" _
                & "<td width=70 align=right style=""" & Style & """>" & scSpanNormalStyle & copy & "</SPAN></td>" _
                & "</tr>")
            '
            ' total
            '
            Style = "color:black;background-color:white;border-top:2px solid #a0a0a0;padding-left:5px;padding-right:5px;padding-top:10px;"
            s.Add ("<tr>" _
                & "<td style=""color:black;background-color:white;padding-left:5px;padding-right:5px;padding-top:10px;"" colspan=2>&nbsp;</td>" _
                & "<td style=""color:black;background-color:white;padding-top:10px;"" align=""right"">" & scSpanNormalStyle & "<B>Total</B></SPAN>&nbsp;</td>" _
                & "<td width=70 align=right style=""" & Style & """>" & scSpanNormalStyle & FormatCurrency(OrdertotalCharge) & "</SPAN></td>" _
                & "</tr>")
        End If
        '
        ' Bottom of table
        '
        s.Add ("</td></tr></table >")
        '
        getCheckoutForm_ItemOrderForm = s.Text
    
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetShippingSelect"): Resume Next
End Function
'
'
'
Private Function GetFormColumn(copy As String, Align As String, RowCount As Long) As String
    '
    Dim iAlign As String
    Dim Style As String
    '
    If (RowCount Mod 2) > 0 Then
        Style = "ccCatRowEven"
    Else
        Style = "ccCatRowOdd"
    End If
    '
    GetFormColumn = "<td valign=""top"" align=""" & Align & """ class=""" & Style & """>" & SpanClassAdminSmall & copy & "</SPAN></td>"
    End Function
''
''
''
'Private Function IsOKToAddUsername() As Boolean
'    Dim cs
'    '
'    If Main.IsAuthenticated Then
'        '
'        ' if you are authenticated to this account, do anything you wish
'        '
'        IsOKToAddUsername = True
'    Else
'        '
'        ' if you are not authenticated, you can only modify guest accounts (non-admin accounts with blank username/password, etc)
'        '
'        cs = Main.OpenCSContentRecord("people", Main.MemberID)
'        If Not Main.IsCSOK(cs) Then
'            IsOKToAddUsername = False
'        Else
'            IsOKToAddUsername = (Not Main.GetCSBoolean(cs, "admin")) And (Not Main.GetCSBoolean(cs, "developer")) And (Main.GetCSText(cs, "username") = "")
'        End If
'        Call Main.CloseCS(cs)
'    End If
'End Function
'
'
'
Private Function AddOrderGetID() As Long
    On Error GoTo ErrorTrap
    '
    Dim csSource As Long
    Dim csOrder As Long
    Dim SQL As String
    '
    csOrder = main.InsertCSContent("Orders")
    If main.IsCSOK(csOrder) Then
        AddOrderGetID = (main.GetCSInteger(csOrder, "ID"))
        orderId = AddOrderGetID
        '
        Call main.SetCS(csOrder, "VisitID", main.visitId)
        Call main.SetCS(csOrder, "VisitorID", main.VisitorID)
        Call main.SetCS(csOrder, "MemberID", CommerceAccountMemberID)
        Call main.SetCS(csOrder, "ShipDate", 0)
        Call main.SetCS(csOrder, "ShipConfirmed", 0)
        Call main.SetCS(csOrder, "ShipWeight", 0)
        Call main.SetCS(csOrder, "ShipMethod", "")
        Call main.SetCS(csOrder, "ItemCharge", 0)
        Call main.SetCS(csOrder, "ShipCharge", 0)
        Call main.SetCS(csOrder, "TaxCharge", 0)
        Call main.SetCS(csOrder, "TotalCharge", 0)
        Call main.SetCS(csOrder, "PayMethod", "")
        Call main.SetCS(csOrder, "PayDate", 0)
        Call main.SetCS(csOrder, "OrderCompleted", 0)
        Call main.SetCS(csOrder, "SpecialPricing", 0)
        '
        ' Save the new order to this visitor (anonymouse or commerceAdmin, both)
        '
        main.VisitorOrderID = AddOrderGetID
        SQL = "update ccvisitors set OrderID=" & AddOrderGetID & " where ID=" & main.VisitorID
        Call main.ExecuteSQL("default", SQL)
    End If
    Call main.CloseCS(csOrder)
    '
    ' add in fields from account, if available
    '
    Call syncOrderAccount(CommerceAccountMemberID, orderId, isSupportedAddress2)

    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("AddOrderGetID"): Resume Next
End Function
'
'
'
Private Sub syncOrderAccount(userId As Long, orderId As Long, isSupportedAddress2 As Boolean)
    Dim csSource As Long
    Dim csOrder As Long
    '
    csOrder = main.OpenCSContentRecord("orders", orderId)
    If main.IsCSOK(csOrder) Then
        csSource = main.OpenCSContentRecord("people", CommerceAccountMemberID)
        If main.IsCSOK(csSource) Then
            Call main.SetCS(csOrder, "BillName", main.GetCSText(csSource, "BillName"))
            Call main.SetCS(csOrder, "BillEmail", main.GetCSText(csSource, "BillEmail"))
            Call main.SetCS(csOrder, "BillPhone", main.GetCSText(csSource, "BillPhone"))
            Call main.SetCS(csOrder, "BillFax", main.GetCSText(csSource, "BillFax"))
            Call main.SetCS(csOrder, "BillCompany", main.GetCSText(csSource, "BillCompany"))
            Call main.SetCS(csOrder, "BillAddress", main.GetCSText(csSource, "BillAddress"))
            Call main.SetCS(csOrder, "BillCity", main.GetCSText(csSource, "BillCity"))
            Call main.SetCS(csOrder, "BillState", main.GetCSText(csSource, "BillState"))
            Call main.SetCS(csOrder, "BillCountry", main.GetCSText(csSource, "BillCountry"))
            Call main.SetCS(csOrder, "BillZip", main.GetCSText(csSource, "BillZip"))
            Call main.SetCS(csOrder, "BillConfirmed", 0)
            Call main.SetCS(csOrder, "ShipName", main.GetCSText(csSource, "ShipName"))
            Call main.SetCS(csOrder, "ShipPhone", main.GetCSText(csSource, "ShipPhone"))
            Call main.SetCS(csOrder, "ShipCompany", main.GetCSText(csSource, "ShipCompany"))
            Call main.SetCS(csOrder, "ShipAddress", main.GetCSText(csSource, "ShipAddress"))
            Call main.SetCS(csOrder, "ShipCity", main.GetCSText(csSource, "ShipCity"))
            Call main.SetCS(csOrder, "ShipState", main.GetCSText(csSource, "ShipState"))
            Call main.SetCS(csOrder, "ShipCountry", main.GetCSText(csSource, "ShipCountry"))
            Call main.SetCS(csOrder, "ShipZip", main.GetCSText(csSource, "ShipZip"))
            '
            If isSupportedAddress2 Then
                Call main.SetCS(csOrder, "BillAddress2", main.GetCSText(csSource, "BillAddress2"))
                Call main.SetCS(csOrder, "ShipAddress2", main.GetCSText(csSource, "ShipAddress2"))
            End If
        End If
        Call main.CloseCS(csSource)
    End If
    Call main.CloseCS(csOrder)
End Sub
''
''==============================================================================
''   Commerce Administrator Add Account
''==============================================================================
''
'Private Function GetAccountAdminAddAccount() As String
'    On Error GoTo ErrorTrap
'    '
'    Dim ps As String
'    Dim stream As New FastStringClass
'    '
'    ' Header Copy
'    '
'    GetAccountAdminAddAccount = getAccountForm_Edit(0)
'    '
'    Exit Function
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("GetAccountAdminAddAccount")
'    End Function
'
'==============================================================================
' Prints the Account menu form
'==============================================================================
'
Private Function GetshopFormAccountMenu() As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim qs As String
    '
    qs = main.RefreshQueryString
    If IsMemberCommerceAdmin Then
        '
        ' Commerce Administrator
        '
        If CommerceAccountMemberID = 0 Then
            s = "" _
                & "<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, "method", "logout") & """>Logout</a>" _
                & "&nbsp;&nbsp;|&nbsp;&nbsp;Edit" _
                & "&nbsp;&nbsp;|&nbsp;&nbsp;Orders" _
                & "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountSelect) & """>Select Account</a>" _
                & ""
        
                '& "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountAdd) & """>Add Account</a>"

        Else
            s = "" _
                & "<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, "method", "logout") & """>Logout</a>" _
                & "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountEdit) & """>Edit</a>" _
                & "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountOrderList) & """>Orders</a>" _
                & "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountSelect) & """>Select Account</a>" _
                & ""
                '& "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountAdd) & """>Add Account</a>"
        End If
    ElseIf Not main.IsAuthenticated Then
        '
        ' not logged in
        '
        s = "" _
            & "<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountLogin) & """>Login</a>" _
            & "&nbsp;&nbsp;|&nbsp;&nbsp;Edit" _
            & "&nbsp;&nbsp;|&nbsp;&nbsp;Orders" _
            & ""
    Else
        '
        ' Normal case
        '
        s = "" _
            & "<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, "method", "logout") & """>Logout</a>" _
            & "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountEdit) & """>Edit</a>" _
            & "&nbsp;&nbsp;|&nbsp;&nbsp;<A class=ccCatMenuLink href=""?" & ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountOrderList) & """>Orders</a>" _
            & ""
    End If
    '
    GetshopFormAccountMenu = "<div class=scSubMenu>" & s & "</div>"
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetshopFormAccountMenu"): Resume Next
    End Function
'
'==============================================================================
' Prints the form called out in AccountForm
'
' Note: call Init before calling this routine
'==============================================================================
'
Public Function GetAccountForm() As String
    On Error GoTo ErrorTrap
    '
    Dim contactPassword As String
    Dim RowCnt As Long
    Dim RecordID As Long
    Dim RowName As String
    Dim AllowCommerceSpecialPrice As Boolean
    Dim CommerceDiscount As Double
    Dim IsInGroup As Boolean
    Dim WasInGroup As Boolean
    Dim dateExpires As Date
    Dim groupName
    Dim stream As New FastStringClass
    Dim copy
    Dim ItemCount As Long
    Dim TempMessage As String
    Dim RedirectPage As String
    Dim cs As Long
    Dim firstName As String
    Dim lastName As String
    Dim CmrcCopy As New CmrcCopyClass
    Dim accountDstForm As Long
    Dim blockContactSave As Boolean
    Dim blockContactGroup As String
    Dim contactEmail As String
    Dim contactUsername As String
    Dim countryId As Long
    Dim countryName As String
    Dim common As New commonClass
    '
    Call main.AddPageTitle("Review your account")
    '
    ' Verify the Order's Member matches the CommerceAccount
    '
    Call LoadOrderGlobals(False)
    '
    ' check if this account can be updated
    '
    blockContactSave = False
    If main.IsAuthenticated Then
        blockContactGroup = Trim(main.GetSiteProperty("ShoppingCart Block Contact Update Group", ""))
        If blockContactGroup <> "" Then
            blockContactSave = common.isGroupMember(main, blockContactGroup, CommerceAccountMemberID)
        End If
    End If
    '
    '-----------------------------------------------------------------------------------
    ' read Input
    '-----------------------------------------------------------------------------------
    '
    AccountSourceForm = main.GetStreamNumber(rnSrcShopFormId)
    accountDstForm = main.GetStreamNumber(rnSrcShopFormId)
    AccountButton = main.GetStreamText(rnButton)
    AccountAction = main.GetStreamNumber("aca")
    AccountLoginUsername = main.GetStreamText("acu")
    AccountLoginPassword = main.GetStreamText("acp")
    '
    '-----------------------------------------------------------------------------------
    ' Process Forms
    '-----------------------------------------------------------------------------------
    '
    If AccountButton <> "" Then
        Select Case AccountSourceForm
            Case shopFormAccountLogin
                accountDstForm = processCommonForm_Login(shopFormAccountMenu, shopFormAccountLogin)
            Case shopFormAccountEdit
                Select Case AccountButton
                    Case ButtonCancel
                        accountDstForm = shopFormAccountMenu
                    Case ButtonSave
                        '
                        ' check for dup email
                        '
                        If Not blockContactSave Then
                            contactEmail = main.GetStreamText("MemberEmail")
                            If contactEmail <> "" Then
                                cs = main.OpenCSContent("people", "(id<>" & CommerceAccountMemberID & ")and(email=" & KmaEncodeSQLText(contactEmail) & ")")
                                If main.IsCSOK(cs) Then
                                    Call main.AddUserError("The contact email address you entered is already in user by another account. Please enter a different email address or logout and login with that account.")
                                End If
                                Call main.CloseCS(cs)
                            End If
                            '
                            contactPassword = main.GetStreamText("MemberPassword")
                            If contactPassword = "" Then
                                Call main.AddUserError("Please enter a valid Password.")
                            End If
                            '
                            If Not kmaEncodeBoolean(main.GetSiteProperty("ALLOWEMAILLOGIN", "")) Then
                                contactUsername = main.GetStreamText("MemberUsername")
                                If contactUsername = "" Then
                                    Call main.AddUserError("Please enter a valid Username.")
                                Else
                                    cs = main.OpenCSContent("people", "(id<>" & CommerceAccountMemberID & ")and(username=" & KmaEncodeSQLText(contactUsername) & ")")
                                    If main.IsCSOK(cs) Then
                                        Call main.AddUserError("The username you entered is already in user by another account. Please enter a different username or logout and login with that account.")
                                    End If
                                    Call main.CloseCS(cs)
                                End If
                            End If
                        End If
                        '
                        If Not main.IsUserError() Then
                            cs = main.OpenCSContentRecord("people", CommerceAccountMemberID)
                            If main.IsCSOK(cs) Then
                                '
                                ' security
                                '
                                If Not kmaEncodeBoolean(main.GetSiteProperty("ALLOWEMAILLOGIN", "")) Then
                                    Call main.SetCS(cs, "Username", contactUsername)
                                End If
                                Call main.SetCS(cs, "Password", main.GetStreamText("MemberPassword"))
                                '
                                ' contact
                                '
                                firstName = main.GetStreamText("MemberFirstName")
                                lastName = main.GetStreamText("MemberLastName")
                                If Not blockContactSave Then
                                    countryId = main.GetStreamInteger("MemberCountry")
                                    countryName = main.GetRecordName("countries", countryId)
                                    Call main.SetCS(cs, "FirstName", firstName)
                                    Call main.SetCS(cs, "LastName", lastName)
                                    Call main.SetCS(cs, "Name", Trim(firstName & " " & lastName))
                                    Call main.SetCS(cs, "Email", contactEmail)
                                    '
                                    Call main.SetCS(cs, "Phone", main.GetStreamText("MemberPhone"))
                                    Call main.SetCS(cs, "Fax", main.GetStreamText("MemberFax"))
                                    Call main.SetCS(cs, "Company", main.GetStreamText("MemberCompany"))
                                    Call main.SetCS(cs, "Address", main.GetStreamText("MemberAddress"))
                                    Call main.SetCS(cs, "City", main.GetStreamText("MemberCity"))
                                    Call main.SetCS(cs, "State", main.GetStreamText("MemberState"))
                                    Call main.SetCS(cs, "Zip", main.GetStreamText("MemberZip"))
                                    Call main.SetCS(cs, "Country", countryName)
                                End If
                                '
                                ' billing
                                '
                                countryId = main.GetStreamInteger("MemberBillCountry")
                                countryName = main.GetRecordName("countries", countryId)
                                Call main.SetCS(cs, "BillName", main.GetStreamText("MemberBillName"))
                                Call main.SetCS(cs, "BillEmail", main.GetStreamText("MemberBillEmail"))
                                Call main.SetCS(cs, "BillPhone", main.GetStreamText("MemberBillPhone"))
                                Call main.SetCS(cs, "BillFax", main.GetStreamText("MemberBillFax"))
                                Call main.SetCS(cs, "BillCompany", main.GetStreamText("MemberBillCompany"))
                                Call main.SetCS(cs, "BillAddress", main.GetStreamText("MemberBillAddress"))
                                Call main.SetCS(cs, "BillCity", main.GetStreamText("MemberBillCity"))
                                Call main.SetCS(cs, "BillState", main.GetStreamText("MemberBillState"))
                                Call main.SetCS(cs, "BillZip", main.GetStreamText("MemberBillZip"))
                                Call main.SetCS(cs, "BillCountry", countryName)
                                '
                                ' shipping
                                '
                                countryId = main.GetStreamInteger("MemberShipCountry")
                                countryName = main.GetRecordName("countries", countryId)
                                Call main.SetCS(cs, "ShipName", main.GetStreamText("MemberShipName"))
                                Call main.SetCS(cs, "ShipCompany", main.GetStreamText("MemberShipCompany"))
                                Call main.SetCS(cs, "ShipAddress", main.GetStreamText("MemberShipAddress"))
                                Call main.SetCS(cs, "ShipCity", main.GetStreamText("MemberShipCity"))
                                Call main.SetCS(cs, "ShipState", main.GetStreamText("MemberShipState"))
                                Call main.SetCS(cs, "ShipZip", main.GetStreamText("memberShipZip"))
                                Call main.SetCS(cs, "ShipCountry", countryName)
                                Call main.SetCS(cs, "ShipPhone", main.GetStreamText("MemberShipPhone"))
                                If isSupportedAddress2 Then
                                    Call main.SetCS(cs, "Address2", main.GetStreamText("MemberAddress2"))
                                    Call main.SetCS(cs, "BillAddress2", main.GetStreamText("MemberBillAddress2"))
                                    Call main.SetCS(cs, "ShipAddress2", main.GetStreamText("MemberShipAddress2"))
                                End If
                                '
                                ' if visitid = createkey, clear it now
                                '
                                If main.visitId = main.GetCSInteger(cs, "createkey") Then
                                    Call main.SetCS(cs, "createkey", 0)
                                End If
                            End If
                            Call main.CloseCS(cs)
                            '
                            ' Group Policies
                            '
                            If IsMemberCommerceAdmin And main.SiteProperty_BuildVersion > "3.3.443" Then
                                '
                                ' Group Based Commerce Discounts
                                '
                                cs = main.OpenCSContent("Groups", , "ID,Name", , , , "Caption,Name,AllowCommerceSpecialPrice,CommerceDiscount")
                                If Not main.IsCSOK(cs) Then
                                Else
                                    Do While main.IsCSOK(cs)
                    
                                        RecordID = main.GetCSInteger(cs, "ID")
                                        groupName = main.GetCSText(cs, "Name")
                                        RowName = "Group" & RecordID
                                        IsInGroup = main.GetStreamBoolean(RowName & "IsInGroup")
                                        WasInGroup = main.GetStreamBoolean(RowName & "WasInGroup")
                                        dateExpires = main.GetStreamDate(RowName & "DateExpires")
                                        If main.GetStreamBoolean(RowName & "block") Then
                                            '
                                            ' blocked - do not change
                                            '
                                        ElseIf IsInGroup And Not WasInGroup Then
                                            Call main.AddGroupMember(groupName, CommerceAccountMemberID, dateExpires)
                                        ElseIf WasInGroup And Not IsInGroup Then
                                            Call main.DeleteGroupMember(groupName, CommerceAccountMemberID)
                                        End If
                                        '
                                        RowCnt = RowCnt + 1
                                        main.NextCSRecord (cs)
                                    Loop
                                End If
                                Call main.CloseCS(cs)
                            End If
                        End If
                        accountDstForm = shopFormAccountEdit
                        ' end case
                    Case ButtonCancel
                        accountDstForm = shopFormCatalogIndex
                        ' end case
                    Case Else
                        Call HandleClassErrorAndResumeNext("AccountInit", "unknown button, form = " & AccountSourceForm & ", button = " & AccountButton)
                    End Select
                ' end case
            End Select
    End If
    '
    If (Not main.IsAuthenticated) Then
        '
        ' if you are not authenticated, you can either login, or go to the new account form
        ' login in the default
        '
        If accountDstForm <> shopFormAccountJoin Then
            accountDstForm = shopFormAccountLogin
        End If
    ElseIf IsMemberCommerceAdmin Then
        '
        ' Commerce administrator
        '
        If accountDstForm <= 0 Then
            If CommerceAccountMemberID = 0 Then
                '
                ' default to select if no account selected
                '
                accountDstForm = shopFormAccountSelect
            Else
                '
                ' default to edit if there is one selected
                '
                accountDstForm = shopFormAccountEdit
            End If
        End If
    End If
    '
    Select Case accountDstForm
        Case shopFormAccountOrderList
            stream.Add GetAccountForm_OrderList
        Case shopFormAccountOrderDetails
            stream.Add GetAccountForm_OrderDetails
        Case shopFormAccountLogin
            stream.Add getCommonForm_Login(shopFormAccountLogin, shopFormAccountLogin)
        Case shopFormAccountJoin
            stream.Add GetAccountForm_Join
        Case shopFormAccountSelect
            stream.Add GetAccountForm_SelectAccount
        Case shopFormAccountAdd
            stream.Add getAccountForm_Edit(0)
        Case Else
            stream.Add getAccountForm_Edit(CommerceAccountMemberID)
    End Select
    '
    ' Assemble page
    '
    GetAccountForm = "" _
        & vbCrLf & vbTab & "<div class=""ccCatBody"">" _
        & kmaIndent(stream.Text) _
        & vbCrLf & vbTab & "</div>" _
        & ""
    GetAccountForm = "" _
        & vbCrLf & vbTab & "<div class=""checkOutCon"">" _
        & kmaIndent(GetshopFormAccountMenu) _
        & kmaIndent(GetAccountForm) _
        & vbCrLf & vbTab & "</div>" _
        & ""
'    GetAccountForm = "" _
'        & vbCrLf & vbTab & "<div class=""checkOutCon"">" _
'        & GetshopFormAccountMenu _
'        & vbCrLf & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=100%><tr><td class=ccCatBody>" & stream.Text & "</td></tr></table >" _
'        & vbCrLf & vbTab & "</div>" _
'        & ""
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetAccountForm"): Resume Next
    End Function
'
'==============================================================================
' Prints the form
'==============================================================================
'
Private Function getAccountForm_Edit(accountUserId As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim countryName As String
    Dim countryId As Long
    Dim RowCnt As Long
    Dim GroupID As Long
    Dim RowName As String
    Dim copy As String
    Dim AllowCommerceSpecialPrice As Boolean
    Dim CommerceDiscount As Double
    Dim FormElement As String
    Dim CommerceAdmin As Long
    Dim IsInGroup As Boolean
    Dim dateExpires As Date
    Dim RowEnabled As Boolean
    Dim IsContentManager As Boolean
    Dim Caption As String
    Dim CommerceAdmin2 As Boolean
    Dim SQL As Long
    Dim SQL2 As String
    Dim CSPeople As Long
    Dim CSGroups As Long
    Dim CSCountry As Long
    Dim CSRule As Long
    Dim CountryOptions As String
    Dim stream As New FastStringClass
    Dim qs As String
    Dim isNewUser As Boolean
    Dim blockContactSave As Boolean
    Dim blockContactGroup As String
    Dim common As New commonClass
    Dim countryUsaId As Integer
    '
    blockContactSave = False
    countryUsaId = main.GetRecordID("countries", "United States")
    If main.IsAuthenticated Then
        blockContactGroup = Trim(main.GetSiteProperty("ShoppingCart Block Contact Update Group", ""))
        If blockContactGroup <> "" Then
            blockContactSave = common.isGroupMember(main, blockContactGroup, CommerceAccountMemberID)
        End If
    End If
    '
    isNewUser = False
    If accountUserId = 0 Then
        isNewUser = True
        accountUserId = main.InsertContentRecordGetID("People")
        CommerceAccountMemberID = accountUserId
        Call main.SetVisitProperty("CommerceAccountMemberID", CommerceAccountMemberID)
        Call LoadOrderGlobals(True)
    End If
    '
    ' Build Country Select options
    '
    CSCountry = main.OpenCSContent_Internal("Countries", , "name", , , , "NAME")
    Do While main.IsCSOK(CSCountry)
        CountryOptions = CountryOptions & "<option> " & main.GetCSText(CSCountry, "Name")
        main.NextCSRecord (CSCountry)
    Loop
    Call main.CloseCS(CSCountry)
    '
    ' Build PageStream
    '
    'QS = Main.RefreshQueryString
    'QS = ModifyQueryString(QS, rnSrcShopFormId, "-1", True)
    'Stream.Add "<div class=scSubMenu>Return to <a href=""?" & QS & """>Account</a> &gt; Edit</div>"
    '
    '
    '
    qs = main.RefreshQueryString
    qs = ModifyQueryString(qs, rnDstShopFormId, shopFormAccountEdit)
    Call stream.Add(main.GetFormStart(qs))
    Call stream.Add("<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormAccountEdit & """>")
    '
    ' Header
    '
    stream.Add DivBR
    If isNewUser Then
        stream.Add GetCommerceCopy(shopCopyNameAccountAddAccount, shopCopyDefaultAccountAddAccount)
    Else
        stream.Add GetCommerceCopy(shopCopyNameAccountEdit, shopCopyDefaultAccountEdit)
    End If
    stream.Add ("<p>Changes made to your account will be reflected on your next order.</p>")
    stream.Add DivBR
    '
    ' Errors
    '
    Call stream.Add(main.GetUserError)
    '
    ' Menus, Input and buttons
    '
    CSPeople = main.OpenCSContentRecord("People", accountUserId)
    If main.IsCSOK(CSPeople) Then
        '
        Call stream.Add("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")
        '
        ' security
        '
        Call stream.Add(GetEditFormRow(True, "Security", ""))
        If kmaEncodeBoolean(main.GetSiteProperty("ALLOWEMAILLOGIN", "")) Then
            Call stream.Add(GetEditFormRow(False, "Username", "(use your email as your username)", True))
        Else
            Call stream.Add(GetEditFormRow(False, "Username", main.GetFormInputText("MemberUsername", main.GetCSText(CSPeople, "username"))))
        End If
        Call stream.Add(GetEditFormRow(False, "Password", main.GetFormInputText("MemberPassword", main.GetCSText(CSPeople, "password"), , , , True)))
        '
        ' contact
        '
        Call stream.Add(GetEditFormRow(True, "Account Contact", ""))
        countryName = main.GetCSText(CSPeople, "Country")
        If countryName <> "" Then
            countryId = main.GetRecordID("countries", countryName)
        End If
        If countryId = 0 Then
            countryName = "United States"
            countryId = countryUsaId
        End If
        '
        If blockContactSave Then
            Call stream.Add(GetEditFormRow(False, "First Name", main.GetCSText(CSPeople, "Firstname"), True))
            Call stream.Add(GetEditFormRow(False, "Last Name", main.GetCSText(CSPeople, "Lastname"), True))
            Call stream.Add(GetEditFormRow(False, "Email", main.GetCSText(CSPeople, "email"), True))
            Call stream.Add(GetEditFormRow(False, "Phone", main.GetCSText(CSPeople, "Phone"), True))
            Call stream.Add(GetEditFormRow(False, "Fax", main.GetCSText(CSPeople, "Fax"), True))
            Call stream.Add(GetEditFormRow(False, "Company", main.GetCSText(CSPeople, "Company"), True))
            Call stream.Add(GetEditFormRow(False, "Address", main.GetCSText(CSPeople, "Address"), True))
            If isSupportedAddress2 Then
                Call stream.Add(GetEditFormRow(False, "Address Line 2", main.GetCSText(CSPeople, "Address2"), True))
            End If
            Call stream.Add(GetEditFormRow(False, "City", main.GetCSText(CSPeople, "City"), True))
            Call stream.Add(GetEditFormRow(False, "State", main.GetCSText(CSPeople, "State"), True))
            Call stream.Add(GetEditFormRow(False, "Postal Code", main.GetCSText(CSPeople, "Zip"), True))
            Call stream.Add(GetEditFormRow(False, "Country", countryName, True))
        Else
            Call stream.Add(GetEditFormRow(False, "First Name", main.GetFormInputText("MemberFirstName", main.GetCSText(CSPeople, "Firstname"))))
            Call stream.Add(GetEditFormRow(False, "Last Name", main.GetFormInputText("MemberLastName", main.GetCSText(CSPeople, "Lastname"))))
            Call stream.Add(GetEditFormRow(False, "Email", main.GetFormInputText("MemberEmail", main.GetCSText(CSPeople, "email"))))
            Call stream.Add(GetEditFormRow(False, "Phone", main.GetFormInputText("MemberPhone", main.GetCSText(CSPeople, "Phone"))))
            Call stream.Add(GetEditFormRow(False, "Fax", main.GetFormInputText("MemberFax", main.GetCSText(CSPeople, "Fax"))))
            Call stream.Add(GetEditFormRow(False, "Company", main.GetFormInputText("MemberCompany", main.GetCSText(CSPeople, "Company"))))
            Call stream.Add(GetEditFormRow(False, "Address", main.GetFormInputText("MemberAddress", main.GetCSText(CSPeople, "Address"))))
            If isSupportedAddress2 Then
                Call stream.Add(GetEditFormRow(False, "Address Line 2", main.GetFormInputText("MemberAddress2", main.GetCSText(CSPeople, "Address2"))))
            End If
            Call stream.Add(GetEditFormRow(False, "City", main.GetFormInputText("MemberCity", main.GetCSText(CSPeople, "City"))))
            Call stream.Add(GetEditFormRow(False, "State", main.GetFormInputText("MemberState", main.GetCSText(CSPeople, "State"))))
            Call stream.Add(GetEditFormRow(False, "Postal Code", main.GetFormInputText("MemberZip", main.GetCSText(CSPeople, "Zip"))))
            Call stream.Add(GetEditFormRow(False, "Country", main.GetFormInputSelect("MemberCountry", countryId, "Countries")))
        End If
        '
        Call stream.Add(GetEditFormRow(True, "Billing Contact", ""))
        '
        countryName = main.GetCSText(CSPeople, "BillCountry")
        If countryName <> "" Then
            countryId = main.GetRecordID("countries", countryName)
        End If
        If countryId = 0 Then
            countryName = "United States"
            countryId = countryUsaId
        End If
        Call stream.Add(GetEditFormRow(False, "Name", main.GetFormInputText("MemberBillName", main.GetCSText(CSPeople, "BillName"))))
        Call stream.Add(GetEditFormRow(False, "Email", main.GetFormInputText("MemberBillEmail", main.GetCSText(CSPeople, "BillEmail"))))
        Call stream.Add(GetEditFormRow(False, "Phone", main.GetFormInputText("MemberBillPhone", main.GetCSText(CSPeople, "BillPhone"))))
        Call stream.Add(GetEditFormRow(False, "Fax", main.GetFormInputText("MemberBillFax", main.GetCSText(CSPeople, "BillFax"))))
        Call stream.Add(GetEditFormRow(False, "Company", main.GetFormInputText("MemberBillCompany", main.GetCSText(CSPeople, "BillCompany"))))
        Call stream.Add(GetEditFormRow(False, "Address", main.GetFormInputText("MemberBillAddress", main.GetCSText(CSPeople, "BillAddress"))))
        If isSupportedAddress2 Then
            Call stream.Add(GetEditFormRow(False, "Address Line 2", main.GetFormInputText("MemberBillAddress2", main.GetCSText(CSPeople, "BillAddress2"))))
        End If
        Call stream.Add(GetEditFormRow(False, "City", main.GetFormInputText("MemberBillCity", main.GetCSText(CSPeople, "BillCity"))))
        Call stream.Add(GetEditFormRow(False, "State", main.GetFormInputText("MemberBillState", main.GetCSText(CSPeople, "BillState"))))
        Call stream.Add(GetEditFormRow(False, "Postal Code", main.GetFormInputText("MemberBillZip", main.GetCSText(CSPeople, "BillZip"))))
        Call stream.Add(GetEditFormRow(False, "Country", main.GetFormInputSelect("MemberBillCountry", countryId, "Countries")))
        '
        Call stream.Add(GetEditFormRow(True, "Shipping Contact", ""))
        '
        countryName = main.GetCSText(CSPeople, "ShipCountry")
        countryId = main.GetRecordID("countries", countryName)
        If countryId = 0 Then
            countryName = "United States"
            countryId = countryUsaId
        End If
        Call stream.Add(GetEditFormRow(False, "Name", main.GetFormInputText("MemberShipName", main.GetCSText(CSPeople, "ShipName"))))
        Call stream.Add(GetEditFormRow(False, "Phone", main.GetFormInputText("MemberShipPhone", main.GetCSText(CSPeople, "ShipPhone"))))
        Call stream.Add(GetEditFormRow(False, "Company", main.GetFormInputText("MemberShipCompany", main.GetCSText(CSPeople, "ShipCompany"))))
        Call stream.Add(GetEditFormRow(False, "Address", main.GetFormInputText("MemberShipAddress", main.GetCSText(CSPeople, "ShipAddress"))))
        If isSupportedAddress2 Then
            Call stream.Add(GetEditFormRow(False, "Address Line 2", main.GetFormInputText("MemberShipAddress2", main.GetCSText(CSPeople, "ShipAddress2"))))
        End If
        Call stream.Add(GetEditFormRow(False, "City", main.GetFormInputText("MemberShipCity", main.GetCSText(CSPeople, "ShipCity"))))
        Call stream.Add(GetEditFormRow(False, "State", main.GetFormInputText("MemberShipState", main.GetCSText(CSPeople, "ShipState"))))
        Call stream.Add(GetEditFormRow(False, "Postal Code", main.GetFormInputText("MemberShipZip", main.GetCSText(CSPeople, "ShipZip"))))
        Call stream.Add(GetEditFormRow(False, "Country", main.GetFormInputSelect("MemberShipCountry", countryId, "Countries")))
        'Call stream.Add(GetEditFormRow(True, "", ""))
'        '
'        ' Group Policies
'        '
'        If IsMemberCommerceAdmin And (main.SiteProperty_BuildVersion > "3.3.443") Then
'            '
'            ' Group Membership
'            '
'            Call stream.Add(GetEditFormRow(True, "Group Policies", ""))
'            CSGroups = main.OpenCSContent("Groups", , "ID,caption,Name", , , , "Caption,Name,AllowCommerceSpecialPrice,CommerceDiscount")
'            'CS = Main.OpenCSContent("Groups", , "ID,caption,Name", , , , "Caption,Name,AllowCommerceSpecialPrice,CommerceDiscount,CommerceAdmin")
'            If Not main.IsCSOK(CSGroups) Then
'                Call stream.Add(kmaStartTableRow & "<td>&nbsp;</td><td>There are no groups defined</td></tr>")
'            Else
'                Do While main.IsCSOK(CSGroups)
'                    GroupID = main.GetCSInteger(CSGroups, "ID")
'                    AllowCommerceSpecialPrice = main.GetCSBoolean(CSGroups, "AllowCommerceSpecialPrice")
'                    CommerceDiscount = main.GetCSNumber(CSGroups, "CommerceDiscount")
'                    RowName = "Group" & GroupID
'                    Caption = main.GetCSText(CSGroups, "Caption")
'                    If Caption = "" Then
'                        Caption = main.GetCSText(CSGroups, "Name")
'                        If Caption = "" Then
'                            Caption = "Group " & GroupID
'                        End If
'                    End If
'                    dateExpires = CDate(0)
'                    IsInGroup = False
'                    '
'                    SQL2 = "select ID,DateExpires from ccMemberRules R" _
'                        & " where (R.active<>0)" _
'                        & " and(R.MemberID=" & accountUserId & ")" _
'                        & " and(R.GroupID=" & GroupID & ")"
'                    CSRule = main.OpenCSSQL("default", SQL2)
'                    If main.IsCSOK(CSRule) Then
'                        IsInGroup = True
'                        dateExpires = main.GetCSDate(CSRule, "DateExpires")
'                    End If
'                    Call main.CloseCS(CSRule)
'                    '
'                    IsContentManager = False
'                    SQL2 = "select ID from ccGroupRules R" _
'                        & " where (R.active<>0)" _
'                        & " and(R.ContentID is not null)" _
'                        & " and(R.GroupID=" & GroupID & ")"
'                    CSRule = main.OpenCSSQL("default", SQL2)
'                    If main.IsCSOK(CSRule) Then
'                        IsContentManager = True
'                    End If
'                    Call main.CloseCS(CSRule)
'                    '
'                    RowEnabled = (Not CommerceAdmin2) And (Not IsContentManager)
'                    '
'                    If Not RowEnabled Then
'                        If IsInGroup Then
'                            FormElement = "<input type=checkbox name=null value=1 disabled checked>"
'                        Else
'                            FormElement = "<input type=checkbox name=null value=1 disabled>"
'                        End If
'                        FormElement = FormElement & "&nbsp;&nbsp;Not available<input type=hidden name=""" & RowName & "Block"" value=1>"
'                        Call stream.Add(GetEditFormRow(False, Caption, FormElement))
'                    Else
'                        FormElement = "" _
'                            & main.GetFormInputHidden(RowName & "WasInGroup", IsInGroup) _
'                            & main.GetFormInputCheckBox(RowName & "IsInGroup", IsInGroup) _
'                            & "&nbsp;&nbsp;Discount&nbsp;" & CommerceDiscount & "%"
'                        If AllowCommerceSpecialPrice Then
'                            FormElement = FormElement & ",&nbsp;&nbsp;Allow Custom Pricing"
'                        End If
'                        Call stream.Add(GetEditFormRow(False, Caption, FormElement))
'                    End If
'                    RowCnt = RowCnt + 1
'                    main.NextCSRecord (CSGroups)
'                Loop
'                Call stream.Add(main.GetFormInputHidden("GroupRowCnt", RowCnt))
'            End If
'            Call main.CloseCS(CSGroups)
'        End If
        '
        ' Fix table column width
        '
        Call stream.Add("<tr>")
        Call stream.Add("<td width=""200"" align=""right""><img border=""0"" src=""/ccLib/images/spacer.gif"" width=""100"" height=""1""></td>")
        Call stream.Add("<td width=""100%""><img border=""0"" src=""/ccLib/images/spacer.gif"" width=""100%"" height=""1""></td>")
        Call stream.Add("</tr>")
        Call stream.Add("</table >")
    End If
    Call main.CloseCS(CSPeople)
    '
    ' Buttons
    '
    Call stream.Add(main.GetPanelButtons(ButtonCancel & "," & ButtonSave, rnButton))
    Call stream.Add("</FORM>")
    '
    getAccountForm_Edit = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getAccountForm_Edit"): Resume Next
    End Function

'
'==============================================================================
' Prints the form
'==============================================================================
'
Private Function GetAccountForm_OrderList() As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim RowCount As Long
    Dim RowStyle As String
    Dim orderId As Long
    Dim stream As New FastStringClass
    Dim qs As String
    '
    ' Build PageStream
    '
    'QS = Main.RefreshQueryString
    'QS = ModifyQueryString(QS, rnSrcShopFormId, "-1", True)
    'Stream.Add "<div class=scSubMenu>Return to <a href=""?" & QS & """>Account</a> &gt; My Orders</div>"
    '
    ' Header
    '
    stream.Add DivBR
    stream.Add GetCommerceCopy(shopCopyNameAccountOrderList, shopCopyDefaultAccountOrderList)
    stream.Add DivBR
    '
    ' Errors
    '
    Call stream.Add(main.GetUserError)
    stream.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""1"" width=""100%"">")
    stream.Add ("<tr>")
    stream.Add ("<td class=ccCatRowHeader align=""left"" width=""10%"">" & "Details</td>")
    stream.Add ("<td class=ccCatRowHeader align=""left"" width=""30%"">" & "Date&nbsp;Ordered</td>")
    stream.Add ("<td class=ccCatRowHeader align=""left"" width=""30%"">" & "Date&nbsp;Shipped</td>")
    stream.Add ("<td class=ccCatRowHeader align=""right"" width=""30%"">" & "Tracking&nbsp;Number</td>")
    stream.Add ("</tr>")
    '
    ' List of orders
    '
    cs = main.OpenCSContent_Internal("Orders", "(OrderCompleted<>0)and(MemberID=" & CommerceAccountMemberID & ")", "ID DESC", , , , "DateCompleted,PayMethod,ShipMethod,ShipDate,ShipTracking")
    If Not main.IsCSOK(cs) Then
        stream.Add ("<tr><td colspan=""4"" class=""ccCatRowEven"">There are no orders on this account</td></tr>")
    Else
        RowCount = 0
        Do While main.IsCSOK(cs)
            orderId = main.GetCSInteger(cs, "ID")
            qs = main.RefreshQueryString
            qs = ModifyQueryString(qs, rnSrcShopFormId, shopFormAccountOrderDetails)
            qs = ModifyQueryString(qs, "ID", CStr(orderId))
            If (RowCount Mod 2) Then
                RowStyle = "class=""ccCatRowEven"""
            Else
                RowStyle = "class=""ccCatRowOdd"""
            End If
            stream.Add ("<tr>")
            stream.Add ("<td " & RowStyle & " align=""center""><A HREF=""?" & qs & """>details</A></td>")
            stream.Add ("<td " & RowStyle & " align=""left""><NOBR>" & KmaEncodeDate(main.GetCSText(cs, "DateCompleted")) & "</SPAN></NOBR></td>")
            stream.Add ("<td " & RowStyle & " align=""left""><NOBR>" & kmaEncodeText(main.GetCSText(cs, "ShipDate")) & "</SPAN></NOBR></td>")
            stream.Add ("<td " & RowStyle & " align=""right""><NOBR>" & kmaEncodeText(main.GetCSText(cs, "ShipTracking")) & "</SPAN></NOBR></td>")
            stream.Add ("</tr>")
            Call main.NextCSRecord(cs)
            RowCount = RowCount + 1
        Loop
    End If
    Call main.CloseCS(cs)
    stream.Add ("</table >")
    '
    'Stream.Add ("<P>Return to <A Href=""" & ModifyQueryString(GetShoppingCartLink(Main.ServerPage, Main.RefreshQueryString), rnSrcShopFormId, shopFormAccountMenu) & """>Account</A></p>")
    '
    GetAccountForm_OrderList = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetAccountForm_OrderList"): Resume Next
    End Function
'
'==============================================================================
' Prints the form
'==============================================================================
'
Private Function GetAccountForm_OrderDetails() As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim RowCount As Long
    Dim RowStyle As String
    Dim orderId As Long
    Dim stream As New FastStringClass
    Dim CSOrderDetails As Long
    Dim QSMain As String
    Dim QSList As String
    '
    'QSMain = Main.RefreshQueryString
    'QSMain = ModifyQueryString(QSMain, rnSrcShopFormId, "-1", True)
    'QSList = Main.RefreshQueryString
    'QSList = ModifyQueryString(QSList, rnSrcShopFormId, shopFormAccountOrderList, True)
    'Stream.Add "<div class=scSubMenu>Return to <a href=""?" & QSMain & """>Account</a> &gt; <a href=""?" & QSList & """>My Orders</a> &gt; My Orders</div>"
    '
    ' Header
    '
    stream.Add DivBR
    stream.Add GetCommerceCopy(shopCopyNameAccountOrderDetails, shopCopyDefaultAccountOrderDetails)
    stream.Add DivBR
    '
    ' Errors
    '
    Call stream.Add(main.GetUserError)
    '
    ' Details of this Order
    '
    orderId = main.GetStreamInteger("id")
    cs = main.OpenCSContent_Internal("Orders", "(ID=" & KmaEncodeSQLNumber(orderId) & ")AND(MemberID=" & KmaEncodeSQLNumber(CommerceAccountMemberID) & ")")
    If Not main.IsCSOK(cs) Then
        stream.Add ("<P>The order you selected could not be found.</p>")
    Else
        '
        RowStyle = "class=""ccPanelRowEven"""
        '
        Call stream.Add("<table border=""0"" cellspacing=""1"" cellpadding=""6"" width=""100%"">")
        stream.Add ("<tr>")
        stream.Add ("<td width=""150""><img SRC=""/ccLib/Images/spacer.gif"" WIDTH=""150"" HEIGHT=""1""></td>")
        stream.Add ("<td width=""100%""><img SRC=""/ccLib/Images/spacer.gif"" WIDTH=""1"" HEIGHT=""1""></td>")
        stream.Add ("</tr>")
        '
        Call stream.Add(GetEditFormRowCS(cs, "DateAdded"))
        Call stream.Add(GetEditFormRowCS(cs, "BillName"))
        Call stream.Add(GetEditFormRowCS(cs, "BillPhone"))
        Call stream.Add(GetEditFormRowCS(cs, "BillEmail"))
        Call stream.Add(GetEditFormRowCS(cs, "BillFax"))
        '
        Call stream.Add(GetEditFormRowCS(cs, "ShipDate"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipName"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipCompany"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipAddress"))
        If isSupportedAddress2 Then
            Call stream.Add(GetEditFormRowCS(cs, "ShipAddress2"))
        End If
        Call stream.Add(GetEditFormRowCS(cs, "ShipCity"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipState"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipZip"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipCountry"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipMethod"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipWeight"))
        Call stream.Add(GetEditFormRowCS(cs, "ShipTracking"))
        '
        Call stream.Add(GetEditFormRowCS(cs, "ItemCharge", FieldTypeCurrency))
        Call stream.Add(GetEditFormRowCS(cs, "ShipCharge", FieldTypeCurrency))
        Call stream.Add(GetEditFormRowCS(cs, "TaxCharge", FieldTypeCurrency))
        Call stream.Add(GetEditFormRowCS(cs, "TotalCharge", FieldTypeCurrency))
        Call stream.Add(GetEditFormRowCS(cs, "PayMethod"))
        Call stream.Add(GetEditFormRowCS(cs, "PayDate"))
        '
        'Call stream.Add(GetEditFormRowCS(cs, "CCType"))
        Call stream.Add(GetEditFormRowCS(cs, "DateCompleted"))
        '
        stream.Add ("</table >")
        '
        CSOrderDetails = main.OpenCSContent_Internal("Order Details", "OrderID=" & KmaEncodeSQLNumber(orderId))
        If Not main.IsCSOK(CSOrderDetails) Then
            '
            ' no items
            '
            stream.Add ("There are no item details for this order")
        Else
            '
            ' create item result table
            '
            
            RowCount = 0
            stream.Add ("<table border=0 cellspacing=1 cellpadding=6 width=100%>")
            stream.Add "<tr>"
            'stream.Add ("<td class=ccCatRowHeader align=right width=10%>Item<br>Number</td>")
            stream.Add ("<td class=ccCatRowHeader align=right width=10%>Quantity</td>")
            stream.Add ("<td class=ccCatRowHeader align=left width=90%>Name</td>")
            'stream.Add ("<td class=ccCatRowHeader align=right width=10%>Unit<br>Price</td>")
            'stream.Add ("<td class=ccCatRowHeader align=right width=10%><br>Weight</td>")
            stream.Add "</tr>"
            RowCount = RowCount + 1
            Do While main.IsCSOK(CSOrderDetails)
                stream.Add GetFormRowStart
                'stream.Add GetFormColumn(main.GetCSText(CSOrderDetails, "ItemNumber"), "right", RowCount)
                stream.Add GetFormColumn(main.GetCSText(CSOrderDetails, "Quantity"), "right", RowCount)
                stream.Add GetFormColumn(main.GetCSText(CSOrderDetails, "name"), "left", RowCount)
                'stream.Add GetFormColumn(Format(main.GetCSNumber(CSOrderDetails, "UnitPrice"), "0.00"), "right", RowCount)
                'stream.Add GetFormColumn(Format(main.GetCSNumber(CSOrderDetails, "weight"), "0.00"), "right", RowCount)
                stream.Add GetFormRowEnd
                RowCount = RowCount + 1
                Call main.NextCSRecord(CSOrderDetails)
            Loop
            stream.Add "</table >"
        End If
        Call main.CloseCS(CSOrderDetails)
    End If
    Call main.CloseCS(cs)
    '
 '   Stream.Add ("<P>Return to <A Href=""" & ModifyQueryString(GetShoppingCartLink(Main.ServerPage, Main.RefreshQueryString), rnSrcShopFormId, shopFormAccountOrderList) & """>Account Order List</A></p>")
    '
    GetAccountForm_OrderDetails = stream.Text
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetAccountForm_OrderDetails"): Resume Next
End Function
'
'==============================================================================
'   Account Join Form
'==============================================================================
'
Private Function GetAccountForm_Join() As String
    On Error GoTo ErrorTrap
    '
    Dim ps As String
    Dim stream As New FastStringClass
    '
    ' Header Copy
    '
    GetAccountForm_Join = "" _
        & DivBR _
        & main.GetUserError _
        & GetCommerceCopy(shopCopyNameAccountJoin, shopCopyDefaultAccountJoin) _
        & DivBR _
        & ""
    ' To Be Completed
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetAccountForm_Join")
    End Function
'
'==============================================================================
'   Commerce Administrator Select Account
'==============================================================================
'
Private Function GetAccountForm_SelectAccount() As String
    On Error GoTo ErrorTrap
    '
    Dim s As New FastStringClass
    Dim stream As New FastStringClass
    Dim SelectAccountName As String
    Dim SelectAccountEmail As String
    Dim SelectAccountPhone As String
    Dim SelectAccountMemberID As String
    Dim SelectAccountNumber As String
    Dim Button As String
    Dim cs As Long
    Dim SQL As String
    Dim SQLCriteria As String
    Dim SQLText As String
    Dim recordCnt As Long
    Dim RecordsPerPage As Long
    Dim RecordPtr As Long
    Dim RecordID As Long
    Dim QSBase As String
    Dim qs As String
    Dim SearchCaption As String
    Dim RowClass As String
    
    '
    ' Header Copy
    '
    s.Add "" _
        & DivBR _
        & main.GetUserError _
        & GetCommerceCopy(shopCopyNameAccountSelectAccount, shopCopyDefaultAccountSelectAccount) _
        & DivBR
    '
    ' Process search
    '
    Button = main.GetStreamText(rnButton)
    If Button = ButtonSearch Then
        SelectAccountName = main.GetStreamText("selectaccountname")
        SelectAccountEmail = main.GetStreamText("selectaccountemail")
        SelectAccountPhone = main.GetStreamText("SelectAccountPhone")
        SelectAccountNumber = main.GetStreamText("SelectAccountNumber")
        SelectAccountMemberID = kmaEncodeInteger(SelectAccountNumber)
        '
        If SelectAccountMemberID <> 0 Then
            SearchCaption = ", Account " & SelectAccountMemberID
            SQLText = KmaEncodeSQLText(SelectAccountName)
            SQLCriteria = SQLCriteria & "AND(ID=" & SelectAccountMemberID & ")"
        Else
            If SelectAccountName <> "" Then
                SearchCaption = SearchCaption & ", '" & SelectAccountName & "' in Name"
                SQLText = KmaEncodeSQLText(SelectAccountName)
                SQLCriteria = SQLCriteria & "AND((Firstname like '%" & Mid(SQLText, 2, Len(SQLText) - 2) & "%')or(LastName like '%" & Mid(SQLText, 2, Len(SQLText) - 2) & "%')or(Name like '%" & Mid(SQLText, 2, Len(SQLText) - 2) & "%'))"
            End If
            If SelectAccountEmail <> "" Then
                SearchCaption = SearchCaption & ", '" & SelectAccountEmail & "' in Email"
                SQLText = KmaEncodeSQLText(SelectAccountEmail)
                SQLCriteria = SQLCriteria & "AND(email like '%" & Mid(SQLText, 2, Len(SQLText) - 2) & "%')"
            End If
            If SelectAccountPhone <> "" Then
                SearchCaption = SearchCaption & ", '" & SelectAccountPhone & "' in Phone"
                SQLText = KmaEncodeSQLText(SelectAccountPhone)
                SQLCriteria = SQLCriteria & "AND(phone like '%" & Mid(SQLText, 2, Len(SQLText) - 2) & "%')"
            End If
        End If
        If SQLCriteria = "" Then
            '
            ' no criteria entered
            '
            s.Add ("<div>Please enter criteria to begin your search.</div>")
        Else
            SQL = "select count(*) as Cnt from ccMembers where " & Mid(SQLCriteria, 4)
            cs = main.OpenCSSQL("default", SQL)
            If main.IsCSOK(cs) Then
                recordCnt = main.GetCSInteger(cs, "cnt")
            End If
            Call main.CloseCS(cs)
            '
            If SearchCaption <> "" Then
                SearchCaption = "<P>Your search for " & Mid(SearchCaption, 3)
            End If
            
            If recordCnt = 0 Then
                s.Add SearchCaption & " returned no records.</p>"
            Else
                If recordCnt = 1 Then
                    s.Add SearchCaption & " returned 1 record.</p>"
                Else
                    s.Add SearchCaption & " returned " & recordCnt & " records.</p>"
                End If
                SQL = "select Name,FirstName,LastName,Email,ID,phone from ccMembers where " & Mid(SQLCriteria, 4) & " order by name"
                cs = main.OpenCSSQL("default", SQL)
                RecordsPerPage = 50
                RecordPtr = 0
                s.Add ("<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"" >")
                s.Add "<tr>"
                s.Add "<td width=10% align=right class=ccCatRowHeader>Account</td>"
                s.Add "<td width=40% align=left class=ccCatRowHeader>Name</td>"
                s.Add "<td width=25% align=left class=ccCatRowHeader>Phone</td>"
                s.Add "<td width=25% align=left class=ccCatRowHeader>Email</td>"
                s.Add "</tr>"
                QSBase = main.RefreshQueryString
                QSBase = ModifyQueryString(QSBase, rnSrcShopFormId, shopFormAccountEdit, True)
                Do While main.IsCSOK(cs) And RecordPtr < RecordsPerPage
                    If RecordPtr Mod 2 = 0 Then
                        RowClass = "Style="""" Class=""ccCatRowOdd"""
                    Else
                        RowClass = "Style="""" Class=""ccCatRowEven"""
                    End If
                    RecordID = main.GetCSInteger(cs, "ID")
                    s.Add "<tr>"
                    s.Add "<td " & RowClass & " align=right><a href=""?" & QSBase & "&SetAccountMemberid=" & RecordID & """>" & RecordID & "</a></td>"
                    s.Add "<td " & RowClass & " align=left>&nbsp;" & main.GetCSText(cs, "name") & " (" & main.GetCSText(cs, "LastName") & "," & main.GetCSText(cs, "FirstName") & ")</td>"
                    s.Add "<td " & RowClass & " align=left>&nbsp;" & main.GetCSText(cs, "phone") & "</td>"
                    s.Add "<td " & RowClass & " align=left>&nbsp;" & main.GetCSText(cs, "email") & "</td>"
                    s.Add "</tr>"
                    Call main.NextCSRecord(cs)
                    RecordPtr = RecordPtr + 1
                Loop
                s.Add ("</table >")
                Call main.CloseCS(cs)
            End If
        End If
    s.Add (DivBR)
    End If
    '
    ' Display search form
    '
    Dim Cell As String
    Cell = "" _
        & main.GetFormStart _
        & "<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & shopFormAccountSelect & """>" _
        & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"" >" _
        & getCheckoutForm_ShippingBilling_Row_Text("Account #", "SelectAccountNumber", SelectAccountNumber, False, False) _
        & getCheckoutForm_ShippingBilling_Row_Text("Name", "SelectAccountName", SelectAccountName, False, False) _
        & getCheckoutForm_ShippingBilling_Row_Text("Email", "SelectAccountEmail", SelectAccountEmail, False, False) _
        & getCheckoutForm_ShippingBilling_Row_Text("Phone", "SelectAccountPhone", SelectAccountPhone, False, False) _
        & "<tr><td width=140><img src=/ccLib/Images/spacer.gif width=140 height=1></td><td width=99%></td></tr>" _
        & "</table >" _
        '& "<div><B>Search Form</B></div>" _
        & ""
    s.Add (main.GetPanel(Cell))
    '
    ' buttons
    '
    s.Add main.GetPanelButtons(ButtonSearch, rnButton)
    s.Add "</form>"
    '
    GetAccountForm_SelectAccount = s.Text
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetAccountForm_SelectAccount")
    End Function
''
''==============================================================================
''   Calculate Shipping Charge for a given shipping method
''       orderShipWeight and OrderItemCharge must be calculated already
''       returns the charge in dollars
''==============================================================================
''
'Private Function CalculateShipCharge(MethodID As Long) As Double
'    On Error GoTo ErrorTrap
'    '
'    Dim sql As String
'    Dim RS As Recordset
'    Dim BaseCharge As Double
'    Dim ChargePerWeight As Double
'    Dim ChargePerCharge As Double
'    '
'    CalculateShipCharge = 0
'    If IsNumeric(MethodID) Then
'        Set RS = Main.OpenTable("orderShipmethods", "ID=" & MethodID)
'        If Not (RS Is Nothing) Then
'            '
'            BaseCharge = KmaEncodeNumber(Main.GetRSField(RS, "BaseCharge"))
'            '
'            ChargePerWeight = KmaEncodeNumber(Main.GetRSField(RS, "ChargePerweight"))
'            '
'            ChargePerCharge = KmaEncodeNumber(Main.GetRSField(RS, "ChargePercharge"))
'            '
'            CalculateShipCharge = BaseCharge + (orderShipWeight * ChargePerWeight) + (OrderItemCharge * ChargePerCharge)
'            RS.Close
'            End If
'        Set RS = Nothing
'        End If
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Set RS = Nothing
'    Call HandleClassErrorAndResumeNext("CalculateShipCharge"): Resume Next
'    End Function

Private Function getCountryAbbreviation(Value As String) As String
    On Error GoTo ErrorTrap
    
    Dim cs As Long
    Dim criteria As String
    
    If IsNumeric(Value) Then
        criteria = "ID=" & Value
    Else
        criteria = "Name=" & main.EncodeSQLText(Value)
    End If
    
    cs = main.OpenCSContent("Countries", criteria, , , , , "Abbreviation")
    If main.CSOK(cs) Then
        getCountryAbbreviation = main.GetCS(cs, "Abbreviation")
    End If
    Call main.CloseCS(cs)
    
    Exit Function

ErrorTrap:
    Call HandleClassErrorAndResumeNext("CalculateShipCharge"): Resume Next
End Function

Private Function getShipPriceUPS(serviceCode, curWeight As String, curZip As String, curCountry As String) As Double
    On Error GoTo ErrorTrap

    Dim stream As String
    Dim appURL As String
    Dim objXML As New XMLHTTP60
    Dim respVal As String
    Dim respArr() As String
    Dim upsPrice As String
    Dim rateChart As String
        
    Select Case main.GetSiteProperty("ShippingDropOff")
        Case 1
            rateChart = "Regular%20Daily%20Pickup"
        Case 2
            rateChart = "On%20Call%20Air"
        Case 3
            rateChart = "One%20Time%20Pickup"
        Case 4
            rateChart = "Letter%20Center"
        Case Else
            rateChart = "Customer%20Counter"
    End Select
        
    'appURL = "http://www.ups.com/using/services/rave/qcost_dss.cgi?AppVersion=1.2&AcceptUPSLicenseAgreement=YES&ResponseType=application/x-ups-rss&ActionCode=3&RateChart=Regular%20Daily%20Pickup&ResidentialInd=yes&PackagingType=00&"
    appURL = "http://www.ups.com/using/services/rave/qcost_dss.cgi?AppVersion=1.2&AcceptUPSLicenseAgreement=YES&ResponseType=application/x-ups-rss&ActionCode=3&RateChart=" & rateChart & "&ResidentialInd=yes&PackagingType=00&"
    appURL = appURL & "ServiceLevelCode=" & serviceCode
    appURL = appURL & "&ShipperPostalCode=" & main.GetSiteProperty("ShippingOriginatingZip")
    appURL = appURL & "&14_origCountry=" & getCountryAbbreviation(main.GetSiteProperty("ShippingOriginatingCountry"))
    appURL = appURL & "&ConsigneePostalCode=" & curZip
    appURL = appURL & "&ConsigneeCountry=" & getCountryAbbreviation(curCountry)
    appURL = appURL & "&PackageActualWeight=" & curWeight

    Call objXML.Open("POST", appURL)
    Call objXML.send
    respVal = objXML.responseText
    
    If (InStr(1, respVal, "Success") <> 0) And (InStr(1, respVal, "%") <> 0) Then
        respArr = Split(respVal, "%")
        getShipPriceUPS = FormatCurrency(respArr(14), 2)
    End If

    Exit Function

ErrorTrap:
    Call HandleError("shipResponder", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function

Private Function isSubscriptionOrder() As Boolean
    On Error GoTo ErrorTrap

    Dim cs As Long
    Dim SQL As String
    Dim orderId As Long
    
    orderId = 0
    SQL = "select OrderID from ccvisitors where ID=" & main.VisitorID
    cs = main.OpenCSSQL("default", SQL)
    If main.IsCSOK(cs) Then
        orderId = main.GetCSInteger(cs, "OrderID")
    End If
    Call main.CloseCS(cs)
    
    SQL = "SELECT Items.GroupID " _
        & "From OrderDetails, Items " _
        & "Where (OrderID = " & orderId & ") " _
        & "and (OrderDetails.ItemID=Items.ID) " _
        & "and (Items.GroupID is Not Null)"
    
    cs = main.OpenCSSQL("Default", SQL)
    isSubscriptionOrder = main.CSOK(cs)
    Call main.CloseCS(cs)

    Exit Function

ErrorTrap:
    Call HandleError("isSubscriptionOrder", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
'
'------------------------------------------------------------------------------------------------------------
'   Returns the paypal form
'
'   This is a PayPal HTML Buy button.
'
'------------------------------------------------------------------------------------------------------------
'
Private Function getCheckoutForm_Billing_PayPalForm(mainobj As Object, ByVal echoString As String, ByVal amount As String, payPalItemName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim pos As Long
    Dim o As String
    Dim thankYouPageLink As String
    Dim paymentConfirmLink As String
    Dim paymentCancelLink As String
    Dim qs As String
    Dim ServerLink As String
    Dim RQS As String
    Dim currency_code As String
    '
    RQS = mainobj.RefreshQueryString
    ServerLink = mainobj.ServerLink
    pos = InStr(1, ServerLink, "?")
    If pos > 0 Then
        ServerLink = Left(ServerLink, pos - 1)
    End If
    '
    qs = ModifyQueryString(RQS, rnSrcShopFormId, shopFormCheckoutReceipt, True)
    'thankYouPageLink = ServerLink & "?" & qs
    '
    thankYouPageLink = ServerLink & "?" & qs
    paymentConfirmLink = mainobj.ServerProtocol & mainobj.ServerHost & "/PayPalConfirm?echo=" & echoString
    paymentCancelLink = mainobj.ServerProtocol & mainobj.ServerHost
    '
    currency_code = "USD"
    '
    o = o & "<form id=""PayPal"" name=""PayPal"" action=""" & mainobj.GetSiteProperty("PayPalPostURL", "https://www.paypal.com/cgi-bin/webscr") & """ method=""post"" style=""display:inline;"">"
    o = o & "<input type=""hidden"" name=""cmd"" value=""_xclick"">"
    o = o & "<input type=""hidden"" name=""business"" value=""" & mainobj.GetSiteProperty("PayPalBusinessName", "") & """>"
    o = o & "<input type=""hidden"" name=""item_name"" value=""" & payPalItemName & """>"
    o = o & "<input type=""hidden"" name=""currency_code"" value=""" & mainobj.GetSiteProperty("PayPalCurrencyCode", "USD") & """>"
    o = o & "<input type=""hidden"" name=""amount"" value=""" & amount & """>"
    o = o & "<input type=""hidden"" name=""return"" value=""" & thankYouPageLink & """>"
    o = o & "<input type=""hidden"" name=""notify_url"" value=""" & paymentConfirmLink & """>"
    o = o & "<input type=""hidden"" name=""shopping_url"" value=""" & paymentCancelLink & """>"
    o = o & "</form>"
    getCheckoutForm_Billing_PayPalForm = o
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Billing_PayPalForm"): Resume Next
End Function
''
''------------------------------------------------------------------------------------------------------------
''   Returns the paypal form
''------------------------------------------------------------------------------------------------------------
''
'Private Function getCheckoutForm_Billing_PayPalForm() As String
'    On Error GoTo ErrorTrap
'    '
'    Dim O As String
'    '
'    O = O & "<form id=""PayPal"" name=""PayPal"" action=""" & main.GetSiteProperty("PayPalPostURL", "https://www.sandbox.paypal.com/cgi-bin/webscr") & """ method=""post"" style=""display:inline;"">"
'    O = O & "<input type=""hidden"" name=""cmd"" value=""_xclick"">"
'    O = O & "<input type=""hidden"" name=""business"" value=""" & main.GetSiteProperty("PayPalBusinessName", "") & """>"
'    O = O & "<input type=""hidden"" name=""item_name"" value=""" & orderId & """>"
'    O = O & "<input type=""hidden"" name=""currency_code"" value=""USD"">"
'    O = O & "<input type=""hidden"" name=""amount"" value=""" & OrdertotalCharge & """>"
'    O = O & "</form>"
'    getCheckoutForm_Billing_PayPalForm = O
'    '
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("getCheckoutForm_Billing_PayPalForm"): Resume Next
'End Function
'
'
'
'
Private Function getCountrySelect(main As Object, currentCountry As String, htmlName As String, htmlStateSelectId As String, dstDiv As String, stateSelectId As String) As String
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Dim o As String
    Dim Name As String
    Dim lcaseCurrentCountry As String
    Dim countryId As Long
    '
    If currentCountry <> "" Then
        countryId = main.GetRecordID("countries", currentCountry)
    End If
    If countryId = 0 Then
        currentCountry = "United States"
        countryId = main.GetRecordID("countries", "United States")
    End If
    '
    lcaseCurrentCountry = LCase(currentCountry)
    cs = main.OpenCSContent("countries", , "name", , , , "name")
    Do While main.IsCSOK(cs)
        Name = main.GetCSText(cs, "name")
        If LCase(Name) = lcaseCurrentCountry Then
            o = o & vbCrLf & vbTab & "<option selected>" & main.GetCSText(cs, "name") & "</option>"
        Else
            o = o & vbCrLf & vbTab & "<option>" & main.GetCSText(cs, "name") & "</option>"
        End If
        Call main.NextCSRecord(cs)
    Loop
    Call main.CloseCS(cs)
    '
    If o = "" Then
        o = "<option>United States</option>"
    End If
    o = "" _
        & vbCrLf & vbTab & "<select onChange=""var e,s;s='';e=document.getElementById('" & htmlStateSelectId & "');if(e){s=e.value};cj.ajax.addon('cartGetStateSelect','country='+this.value+'&state='+s+'&element=" & stateSelectId & "','','" & dstDiv & "');cartResetShippingSelect();"" id=""" & htmlName & """ name=""" & htmlName & """>" _
        & kmaIndent(o) _
        & vbCrLf & vbTab & "</select>"
    getCountrySelect = o
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Billing_PayPalForm"): Resume Next
End Function
'
'
'
Private Function isShippingConfigured() As Boolean
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    '
    isShippingConfigured = False
    cs = main.OpenCSContent("Order Ship Methods", "(name<>'')and(name is not null)", , , , , "ID")
    isShippingConfigured = isShippingConfigured Or main.IsCSOK(cs)
    Call main.CloseCS(cs)
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSNextDatAirAM", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSNextDatAir", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSNextDatAirSaver", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPS2ndDayAirAM", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPS2ndDayAir", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPS3DaySelect", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSGround", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSCanadaStandard", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSCWorldwideExpress", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSCWorldwideExpressPlus", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSCWorldwideExpedited", "0"))
    isShippingConfigured = isShippingConfigured Or kmaEncodeBoolean(main.GetSiteProperty("AllowUPSCWorldwideSaver", "0"))

    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCheckoutForm_Billing_PayPalForm"): Resume Next
End Function
'
'==============================================================================
'
'==============================================================================
'
Private Function getCommonForm_Login(srcForm As Long, dstForm As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim ps As String
    Dim stream As New FastStringClass
    Dim newAccountForm As String
    Dim loginForm As String
    Dim loginAction As String
    Dim pos As Long
    Dim posStart As Long
    Dim posEnd As Long
    Dim Ptr As Long
    '
    '
    '
    'Call main.AddRefreshQueryString(rnSrcShopFormId, srcForm)
    Call main.AddRefreshQueryString(rnDstShopFormId, dstForm)
    '
    newAccountForm = "" _
        & vbCrLf & vbTab & "<div class=""scAccountCaption"">Email</div>" _
        & vbCrLf & vbTab & "<div class=""scAccountInput""><input type=""text"" name=""" & rnAccountUsername & """ value=""" & main.MemberEmail & """></div>" _
        & vbCrLf & vbTab & "<div>&nbsp;</div>" _
        & vbCrLf & vbTab & "<div class=""scButtonCon""><input type=""submit"" name=""button"" value=""" & scButtonCreateAccount & """></div>" _
        & ""
    '
    newAccountForm = "" _
        & vbCrLf & vbTab & main.GetFormStart() _
        & kmaIndent(newAccountForm) _
        & vbCrLf & vbTab & "<input type=""hidden"" name=""" & rnSrcShopFormId & """ value=""" & srcForm & """>" _
        & vbCrLf & vbTab & "</form>" _
        & ""
    '
    '
    '
    loginAction = main.RefreshQueryString
    loginForm = main.GetFormLogin
    pos = InStrRev(loginForm, "action=", , vbTextCompare)
    Do While (Ptr < 10) And (pos <> 0)
        pos = pos + 1
        posStart = InStr(pos, loginForm, """")
        If posStart > 0 Then
            posEnd = InStr(posStart + 1, loginForm, """")
            If posEnd > 0 Then
                loginForm = Mid(loginForm, 1, posStart - 1) & """?" & loginAction & """" & Mid(loginForm, posEnd + 1)
            End If
        End If
        pos = InStrRev(loginForm, "action=", pos, vbTextCompare)
        Ptr = Ptr + 1
    Loop
    '
    ' Header Copy
    '
    getCommonForm_Login = "" _
        & DivBR _
        & main.GetUserError _
        & vbCrLf & vbTab & "<fieldset>" _
        & vbCrLf & vbTab & "<legend>New Account</legend>" _
        & GetCommerceCopy(shopCopyNameAccountNewCustomer, shopCopyDefaultAccountNewCustomer) _
        & newAccountForm _
        & vbCrLf & vbTab & "</fieldset>" _
        & "<div>&nbsp</div>" _
        & vbCrLf & vbTab & "<fieldset>" _
        & vbCrLf & vbTab & "<legend>Log in</legend>" _
        & GetCommerceCopy(shopCopyNameAccountLogin, shopCopyDefaultAccountLogin) _
        & kmaIndent(loginForm) _
        & vbCrLf & vbTab & "</fieldset>" _
        & "<div>&nbsp</div>" _
        & vbCrLf & vbTab & "<fieldset>" _
        & vbCrLf & vbTab & "<legend>Forget Your Password?</legend>" _
        & GetCommerceCopy(shopCopyNameAccountSendPassword, shopCopyDefaultAccountSendPassword) _
        & kmaIndent(main.GetSendPasswordForm) _
        & vbCrLf & vbTab & "</fieldset>" _
        & ""
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCommonForm_Login")
    End Function
'
'
'
Private Function processCommonForm_Login(successFormId As Long, failureFormId As Long) As Long
    On Error GoTo ErrorTrap
    '
    Dim returnId As Long
    Dim accountEmail As String
    Dim SQL As String
    Dim cs As String
    Dim emailDup As String
    '
    returnId = failureFormId
    'returnId = shopFormCheckoutAccount
    If main.IsAuthenticated Then
        '
        ' successful login
        '
        returnId = successFormId
        'returnId = shopFormCheckoutShippingBilling
    ElseIf main.GetStreamText("button") = scButtonCreateAccount Then
        '
        ' attempt to create account
        '
        accountEmail = main.GetStreamText(rnAccountUsername)
        If accountEmail = "" Then
            Call main.AddUserError("You must enter a valid email to create an account.")
        ElseIf Len(accountEmail) < 6 Then
            Call main.AddUserError("You must enter a valid email to create an account.")
        ElseIf (InStr(1, accountEmail, "@") <= 0) Then
            Call main.AddUserError("You must enter a valid email to create an account.")
        ElseIf (InStr(1, accountEmail, ".") <= 0) Then
            Call main.AddUserError("You must enter a valid email to create an account.")
        Else
            SQL = "select id from ccmembers where email=" & KmaEncodeSQLText(accountEmail)
            cs = main.OpenCSSQL("", SQL)
            emailDup = main.IsCSOK(cs)
            Call main.CloseCS(cs)
            If emailDup Then
                Call main.AddUserError("The email you entered is already in use. If this is your email, please request your password below.")
            Else
                '
                ' this is a valid email, logout, login and save this email to this account
                '
                Call main.LogoutMember
                SQL = "update ccmembers set createkey=" & main.visitId & ",name='',firstname='',lastname='',email=" & KmaEncodeSQLText(accountEmail) & " where id=" & main.memberId
                Call main.ExecuteSQL("", SQL)
                Call main.LoginMemberByID(main.memberId)
                CommerceAccountMemberID = main.memberId
                Call LoadOrderGlobals(True)
                returnId = successFormId
            End If
        End If
    End If
    '
    processCommonForm_Login = returnId
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("processCommonForm_Login")
End Function


