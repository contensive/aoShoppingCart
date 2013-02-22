VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdminControlClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'========================================================================
' This file and its contents are copyright by Kidwell McGowan Associates.
'========================================================================
'
'============================================================================================================================
' Global Objects
'============================================================================================================================
'
Private main As Object
'
Const WidthSmall = 5
Const WidthWide = 40
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
    Call HandleError("AdminControlClass", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("AdminControlClass", MethodName, Description, ErrorNumber)
End Sub
'
'========================================================================
'   Init()
'   Read in all Request value here so we can build the refresh string
'========================================================================
'
Public Sub Init(MainPointer As Object)
    Set main = MainPointer
    On Error GoTo ErrorTrap
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("Init"): Resume Next
    End Sub
'
'
'
Public Function GetContent(optionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim Copy As String
    Dim FormElement As String
    Dim Button As String
    Dim SaveAction As Boolean
    Dim HelpCopy As String
    Dim FieldValue As String
    Dim PaymentProcessMethod As Long
    Dim CCID As Long
    Dim cs As Long
    Dim AllowCommerceSpecialPrice As Boolean
    Dim AllowCommerceAddButton As Boolean
    Dim CommerceDiscount As Double
    Dim RowName As String
    Dim RowCnt As Long
    Dim RecordID As Long
    Dim CommerceAdmin As Boolean
    'Dim AdminUI As New AdminUIClass
    Dim AdminUI As Object
    Dim Caption As String
    Dim ButtonListLeft As String
    Dim ButtonListRight As String
    Dim Description As String
    Dim ContentSummary As String
    Dim content As New FastStringClass
    Dim common As New commonClass

    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    
    '
    ' Premissions Check
    '
Call AppendLogFile("IN")
    If Not main.IsAdmin Then
        '
        ' Must be admin to use this tool
        '
        ButtonListLeft = ButtonCancel
        content.Add AdminUI.GetFormBodyAdminOnly(main)
    Else
        '
        ' Process Requests
        '
        Button = main.GetStreamText(RequestNameButton)
        SaveAction = (Button = ButtonSave) Or (Button = ButtonOK)
        '
        If Button = ButtonCancel Then
            Call main.Redirect(kmaEncodeAppRootPath(main.GetSiteProperty("AdminURL"), main.ServerVirtualPath, main.ServerAppRootPath, main.ServerHost))
        End If
        '
        content.Add main.GetFormInputHidden(RequestNameRefreshBlock, main.GetFormSN)
        content.Add (AdminUI.EditTableOpen)
        '
        ' Catalog Properties
        '
        Call content.Add(kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Catalog</B>" & kmaEndTableCell & kmaEndTableRow)
'        '
'        HelpCopy = "Set the format of the catalog listing display."
'        If Not SaveAction Then
'            FieldValue = main.GetSiteProperty("CatalogIndexFormat")
'        Else
'            FieldValue = main.GetStreamText("CatalogIndexFormat")
'            Call main.SetSiteProperty("CatalogIndexFormat", FieldValue)
'        End If
'        Copy = "<select size=1 name=CatalogIndexFormat>" _
'            & "<option value=0>Normal Item Listing</option>" _
'            & "<option value=1>Specials Only Listing</option>" _
'            & "</select>"
'        Copy = Replace(Copy, "value=" & FieldValue, "selected value=" & FieldValue) _
'            & "<div class=""ssHelpText"">Use specials only format to create a catalog with only the current specials.</div>"
'        Call Content.Add(GetContent_GetHTMLElement(Copy, "Catalog Listing Format", HelpCopy, False, False))
        '
        Call content.Add("" _
            & kmaStartTableRow _
            & "<td>&nbsp;</td>" _
            & "<td class=""ssHelpPanel"" colspan=2>" _
            & "<div class=""ssHelpText"">This group of settings applies to the catalog.</div>" _
            & "</td>" _
            & "</tr>" _
            & "")
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthSmall, "CataloglistingColumns", FieldTypeInteger, SaveAction, "1")) _
            & "<div class=""ssHelpText"">This is the number of columns displayed when many items are listed out, such as search results and when an entire category is selected.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Catalog Listing Columns", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthSmall, "CatalogRecordSpan", FieldTypeInteger, SaveAction, "1")) _
            & "<div class=""ssHelpText"">This is the number of items displayed per page, such as search results and when an entire category is selected.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Catalog Listing Items", HelpCopy, False, False))
        '
        HelpCopy = ""
        If Not SaveAction Then
            FieldValue = main.GetSiteProperty("CatalogItemOrder")
        Else
            FieldValue = main.GetStreamText("CatalogItemOrder")
            Call main.SetSiteProperty("CatalogItemOrder", FieldValue)
        End If
        Copy = main.GetFormInputSelect("CatalogItemOrder", FieldValue, "Sort Methods") _
            & "<div class=""ssHelpText"">How Items are sorted on the catalog.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Catalog Item Sort Order", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAllowPricing", FieldTypeBoolean, SaveAction, True)) _
            & "<div class=""ssHelpText"">When checked, pricing is displayed with an item in the catalog. Pricing is required to allow ordering.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow prices in catalog", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAllowOrdering", FieldTypeBoolean, SaveAction, True)) _
            & "<div class=""ssHelpText"">When checked, items in the catalog can be added to a shopping cart and purchased through the order process. To only allow certain users to order from the catalog, see 'Group Policies' below.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow ordering for everyone", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAliasRetailPrice", FieldTypeText, SaveAction, "Retail Price")) _
            & "<div class=""ssHelpText"">For comparison pricing only. 'Retail Price', or 'MSRP'. Leave blank if not needed.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "'Retail price' caption", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAliasRegularPrice", FieldTypeText, SaveAction, "Regular Price")) _
            & "<div class=""ssHelpText"">This is the normal price for the item. If there are no other sales or group discounts, this is what the customer will pay.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "'Regular price' caption", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAllowSpecialPrice", FieldTypeBoolean, SaveAction, False)) _
            & "<div class=""ssHelpText"">When checked, items with a 'custom price' will display the custom pricing. Use the 'Custom Price' caption to label your pricing scheme, for example 'Member Price' or 'Seasonal Price'. See 'Group Policies' for an explaination of custom pricing.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow custom prices", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAliasSpecialPrice", FieldTypeText, SaveAction, "Member Price")) _
            & "<div class=""ssHelpText"">When Allow custom prices is checked, this caption is used to identify the custom price. See 'Group Policies' for an explaination of custom pricing.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "'Custom price' caption", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAliasSalesPrice", FieldTypeText, SaveAction, "Sale Price")) _
            & "<div class=""ssHelpText"">When an item includes a sale price, and is marked on sale, this is the caption that displays with the sale price.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "'Sale price' caption", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAliasYourPrice", FieldTypeText, SaveAction, "Your Price")) _
            & "<div class=""ssHelpText"">The lowest possible price is calculated for each item, based on regular price, sales price and any Group Discounts. This caption is used to display that price.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "'Your price' caption", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogMfgPhrase", FieldTypeText, SaveAction, "Mfg by ")) _
            & "<div class=""ssHelpText"">When included, this caption proceeds the name of the source of this item, as selected for each item.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "'Manufactured by' caption", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthSmall, "CatalogThumbnailWidth", FieldTypeInteger, SaveAction, "100")) _
            & "<div class=""ssHelpText"">This is the width of all thumbnail images displayed in the catalog.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Catalog Thumbnail Width", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthSmall, "CatalogImageWidth", FieldTypeInteger, SaveAction, "100")) _
            & "<div class=""ssHelpText"">This is the width of all images displayed in the catalog.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Catalog Image Width", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = "" _
            & GetContent_GetCopyElement("Catalog Home", shopCopyNameCatalogHome, shopCopyDefaultCatalogHome) _
            & GetContent_GetCopyElement("Search Tab: Home", shopCopyNameSearchHome, shopCopyDefaultSearchHome) _
            & GetContent_GetCopyElement("Account Tab: Edit Account", shopCopyNameAccountEdit, shopCopyDefaultAccountEdit) _
            & GetContent_GetCopyElement("Account Tab: Order List", shopCopyNameAccountOrderList, shopCopyDefaultAccountOrderList) _
            & GetContent_GetCopyElement("Account Tab: Select Account (admin only)", shopCopyNameAccountSelectAccount, shopCopyDefaultAccountSelectAccount) _
            & GetContent_GetCopyElement("Account Tab: Add Account (admin only)", shopCopyNameAccountAddAccount, shopCopyDefaultAccountAddAccount) _
            & ""
        Call content.Add(GetContent_GetHTMLElement(Copy, "Catalog Text", HelpCopy, False, False))
        '
        ' Checkout
        '
        Call content.Add(kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Checkout</B>" & kmaEndTableCell & kmaEndTableRow)
        Call content.Add(kmaStartTableRow & "<td>&nbsp;</td><td class=""ssHelpPanel"" colspan=2><div class=""ssHelpText"">" _
            & "Checkout settings control the process of purchasing the items in a shopping cart. If you will be using PayPal or an online credit card processor, see the settings below. Configure onlineyour online payment processor with <a href=""?addonguid=%7B1FC72AE3-19CD-4530-9B44-4D6C51E292CB%7D"">Payment Process Settings</a>" _
            & "</div></td></tr>")
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "EMAILORDERFROMADDRESS", FieldTypeText, SaveAction, "")) _
            & "<div class=""ssHelpText"">This is the email address the appears as the 'From Address' on receipts sent to customers.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Email from-address for order receipt", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "EMAILORDERNOTIFYADDRESS", FieldTypeText, SaveAction, "")) _
            & "<div class=""ssHelpText"">This is the email address when the notification of the sale will be sent. This should be the administrator who ships the order.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Email notification to-address", HelpCopy, False, False))
        ''
        'HelpCopy = "<div class=""ssHelpText"">" _
        '    & "<p>This feature controls how a customer's contact information is saved. Each user in the site has a people record. That people record has three sets of address information: contact, shipping and billing.</p>" _
        '    & "<p>For shoppers who have NOT logged in, their contact name, phone and email is always updated from the billing address they provide during the checkout process.</p>" _
        '    & "<p>If this checkbox is checked, shoppers who have logged in will have their contact name, phone and email  updated from the billing address they provide.</p>" _
        '    & "<p>If un-checked, the contact name, phone and email for shoppers who have logged in will not be changed.</p>" _
        '    & "<p>Typically you would uncheck this box if your people records are input and controlled though another process. Check this box if the catalog is the primary source of new user accounts.</p>" _
        '    & "</div>"
        'copy = (GetContent_GetPropertyElement(WidthWide, "orderUpdateMemberContactFromBilling", FieldTypeBoolean, SaveAction, False)) _
        '    & HelpCopy
        'Call content.Add(GetContent_GetHTMLElement(copy, "Update User Contact from Billing Data", HelpCopy, False, False))
        '
        ' Free Shipping
        '
        HelpCopy = ""
        Copy = "" _
            & (GetContent_GetPropertyElement(WidthWide, "OrderAllowFreeShippingPrice", FieldTypeBoolean, SaveAction, False)) & "&nbsp;Allow Price Level Free Shipping" _
            & "<BR>" _
            & "<BR>" & (GetContent_GetPropertyElement(WidthWide, "OrderFreeShippingPrice", FieldTypeText, SaveAction, "")) & "&nbsp;Price Level" _
            & "<BR>" _
            & "<BR><div  class=""ssHelpText"">When allowed, orders over the Free Shipping Price Level will be given free shipping. The total price is before tax and shipping calculations, but includes all other discounts. Free shipping applies only to those shipping methods with 'Free Shipping at a Price Level' checked.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Price Level Free Shipping", HelpCopy, False, False))
        '
        ' Discount Price Level
        '
        HelpCopy = ""
        Copy = "" _
            & (GetContent_GetPropertyElement(WidthWide, "OrderAllowPriceLevelDiscount", FieldTypeBoolean, SaveAction, False)) & "&nbsp;Allow Price Level Discount" _
            & "<BR>" _
            & "<BR>" & (GetContent_GetPropertyElement(WidthWide, "OrderPriceLevelDiscountPrice", FieldTypeText, SaveAction, "")) & "&nbsp;Price Level" _
            & "<BR>" _
            & "<BR>" & (GetContent_GetPropertyElement(WidthWide, "OrderPriceLevelDiscountDiscount", FieldTypeText, SaveAction, "")) & "&nbsp;Discount Percentage" _
            & "<BR>" _
            & "<BR><div class=""ssHelpText"">When checked, orders over a specified price total will be given a discount. The total price is before tax and shipping calculations, but includes all other discounts. For example, for a 10% discount on orders over $200, enter 200 in Price Level and 10 in Discount Percentage.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Price Level Discount", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = "" _
            & GetContent_GetCopyElement("Shopping Cart", shopCopyNameReview, shopCopyDefaultReview) _
            & GetContent_GetCopyElement("Shopping Cart: Empty", shopCopyNameReviewEmpty, shopCopyDefaultReviewEmpty) _
            & GetContent_GetCopyElement("Contact Page: Top", shopCopyNameContact, shopCopyDefaultContact) _
            & GetContent_GetCopyElement("Contact Page: Contact Address Section", shopCopyNameContactAddress, shopCopyDefaultContactAddress) _
            & GetContent_GetCopyElement("Contact Page: Billing Address Section", shopCopyNameBillAddress, shopCopyDefaultBillAddress) _
            & GetContent_GetCopyElement("Contact Page: Shipping Address Section", shopCopyNameShipAddress, shopCopyDefaultShipAddress) _
            & GetContent_GetCopyElement("Contact Page: Shipping Method Section", shopCopyNameShipMethod, shopCopyDefaultShipMethod) _
            & GetContent_GetCopyElement("Contact Page: Payment Method Section", shopCopyNamePaymentMethod, shopCopyDefaultPaymentMethod) _
            & GetContent_GetCopyElement("Review Page: Top", shopCopyNamePurchase, shopCopyDefaultPurchase) _
            & GetContent_GetCopyElement("Receipt Page", shopCopyNameReceipt, shopCopyDefaultReceipt) _
            & ""
        Call content.Add(GetContent_GetHTMLElement(Copy, "Order Process Text", HelpCopy, False, False))
        '
        ' ----- Paypal payment Processing
        '
        Call content.Add(kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Accepting Paypal Payments</B>" & kmaEndTableCell & kmaEndTableRow)
        '
        Call content.Add(kmaStartTableRow & "<td>&nbsp;</td><td class=""ssHelpPanel"" colspan=2><div class=""ssHelpText"">" _
            & "With PayPal Express checkout, your visitors will see a PalPal link that takes them to the PayPal site to process charges. After the purchase is complete, they return to your site. To use PayPal, you must create a PayPal account on the PayPal site and update these fields." _
            & "</div></td></tr>")
        '
        HelpCopy = "Check this box if you want to use PayPal Express checkout."
            '& "<ol style=""list-style-position:outside;margin:10px 0 0 20px;"">" _
            '& "<li style=""margin:10px 0 0 0;"">Create a PayPal 'seller' account on the PayPal site.</li>" _
            '& "<li style=""margin:10px 0 0 0;"">Set the PayPal Business name below to your PayPal login username.</li>" _
            '& "<li style=""margin:10px 0 0 0;"">The PayPal Post URL below should be one of the following:<br>for testing - https://www.sandbox.paypal.com/cgi-bin/webscr<br>for production - https://www.paypal.com/cgi-bin/webscr</li>" _
            '& "</ol>" _
            '& ""
            '& "<li style=""margin:10px 0 0 0;"">On your PayPal seller account, set the Return URL. This is where the visitor will go when return after paying. From your PayPal account go to Profile - More Options - Seller Preferences - Website payment Preferences and set the Return URL to your URL.</li>" _
            '& "<li style=""margin:10px 0 0 0;"">On your PayPal seller account, set the Instant Payment Notification.This is a special webpage that Paypal will use to signal your website that the order is complete. Set it to 'http://(your website)/PayPalConfirm'.</li>"
        Copy = GetContent_GetPropertyElement(WidthWide, "OrderAllowPayPal", FieldTypeBoolean, SaveAction, False) _
            & "<div class=""ssHelpText"">" & HelpCopy & "</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow payment by PayPal", "", False, False))
        '
        HelpCopy = "When you create your PayPal seller account, this is the username they assign to you."
        Copy = GetContent_GetPropertyElement(WidthWide, "PayPalBusinessName", FieldTypeText, SaveAction, "PayPalBusinessName") _
            & "<div class=""ssHelpText"">" & HelpCopy & "</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "PayPal Business Name", "", False, False))
        '
        HelpCopy = "The PayPal Post URL below should be one of the following:" _
            & "<br>&nbsp;&nbsp;&nbsp;for testing - https://www.sandbox.paypal.com/cgi-bin/webscr" _
            & "<br>&nbsp;&nbsp;&nbsp;for production - https://www.paypal.com/cgi-bin/webscr</li>"
        Copy = GetContent_GetPropertyElement(WidthWide, "PayPalPostURL", FieldTypeText, SaveAction, "https://www.paypal.com/cgi-bin/webscr") _
            & "<div class=""ssHelpText"">" & HelpCopy & "</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "PayPal Post URL", "", False, False))
        '
        ' ----- payment
        '
        Call content.Add(kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Accepting Online Payments</B>" & kmaEndTableCell & kmaEndTableRow)
        '
        Call content.Add(kmaStartTableRow & "<td>&nbsp;</td><td class=""ssHelpPanel"" colspan=2><div class=""ssHelpText"">" _
            & "Online payments are accepted right on your site. You can choose from several options that allow visitors to send checks, prepay, etc. You may also choose to allow visitors to process credit cards online with the Allow online payment (Credit Card) option. When online credit card payments are selected, you must also configure a Payment Processor such as AuthorizeNet through the Payment Processing Add-on. For each option, you can customize instructions with the Payment Text which appears on the payment page and the Receipt Text which appears on the receipt." _
            & "<br><br>If you select any of the credit card options, check a box in front of the credit card(s) you wish to accept." _
            & "</div></td></tr>")
        '
        HelpCopy = ""
        Copy = GetContent_GetPropertyElement(WidthWide, "ORDERALLOWCODCERTIFIEDFUNDS", FieldTypeBoolean, SaveAction, False) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCODCertifiedFunds), "COD certified funds") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCODCertifiedFunds), "COD certified funds")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow COD certified funds payment", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = GetContent_GetPropertyElement(WidthWide, "ORDERALLOWCODCOMPANYCHECK", FieldTypeBoolean, SaveAction, False) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCODCompanyCheck), "COD pre-approved payment") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCODCompanyCheck), "COD pre-approved payment")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow COD pre-approved payment", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWCOMPANYCHECKS", FieldTypeBoolean, SaveAction, False)) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCHECKCOMPANY), "Company Check") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCHECKCOMPANY), "Company Check")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow company check payment", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWCREDIT", FieldTypeBoolean, SaveAction, False)) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCREDIT), "Credit") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCREDIT), "Credit")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow billing payment", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWPERSONALCHECKS", FieldTypeBoolean, SaveAction, False)) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCHECK), "Personal Check") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCHECK), "Personal Check")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow personal check payment", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWCREDITCARDBYFAX", FieldTypeBoolean, SaveAction, False)) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCreditCardByFax), "Credit Card by Fax") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCreditCardByFax), "Credit Card by Fax")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow fax payment (credit card)", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWCREDITCARDBYPHONE", FieldTypeBoolean, SaveAction, False)) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCreditCardByPhone), "Credit Card by Phone") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCreditCardByPhone), "Credit Card by Phone")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow phone payment (credit card)", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWCREDITCARDONLINE", FieldTypeBoolean, SaveAction, False)) _
            & GetContent_GetCopyElement("Payment Text", shopCopyNamePaymentDetailPrefix & PayString(PayTypeCreditCardOnline), "Credit Card Online") _
            & GetContent_GetCopyElement("Receipt Text", shopCopyNameReceptDetailPrefix & PayString(PayTypeCreditCardOnline), "Credit Card Online")
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow online payment (credit card)", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWDISCOVER", FieldTypeBoolean, SaveAction, False)) _
            & "<div class=""ssHelpText"">When Allow Online Payment is checked, this checkbox will cause the Discover icon to appear in the checkout form.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Credit card: Discover", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWMC", FieldTypeBoolean, SaveAction, False)) _
            & "<div class=""ssHelpText"">When Allow Online Payment is checked, this checkbox will cause the MasterCard icon to appear in the checkout form.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Credit card: Master Card", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWVISA", FieldTypeBoolean, SaveAction, False)) _
            & "<div class=""ssHelpText"">When Allow Online Payment is checked, this checkbox will cause the Visa icon to appear in the checkout form.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Credit card: Visa", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ORDERALLOWAMEX", FieldTypeBoolean, SaveAction, False)) _
            & "<div class=""ssHelpText"">When Allow Online Payment is checked, this checkbox will cause the American Express icon to appear in the checkout form.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Credit card: American Express", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = GetContent_GetPropertyElement(WidthWide, "OrderSaveCreditCard", FieldTypeBoolean, SaveAction, False) _
            & "<div class=""ssHelpText"">When checked, the credit card will be saved in the order record. For security reasons, you should leave this blank unless you need to save the credit card number. For instance, if you choose not to process cards online, you should select this checkbox to save the whole credit card number for processing later.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Save credit card number in Db", HelpCopy, False, False))
        '
        HelpCopy = "This field only applies to configurations with Online Credit Card processing enabled. Leave this field blank when taking orders online. Entering text here will let you successfully complete an order on your site using without processing a credit card."
        Copy = GetContent_GetPropertyElement(WidthWide, "OrderTestCCNumber", FieldTypeText, SaveAction, "") _
            & "<div class=""ssHelpText"">Enter a number sequence you will use to test your order process. When the credit card matches this number, the order will be completed without any online processing. Leave this blank to disable this feature.</div>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Testing Credit Card Number", HelpCopy, False, False))
        '
        ' ----- shipping
        '
        Call content.Add(kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Shipping</B>" & kmaEndTableCell & kmaEndTableRow)
        '
        Call content.Add(kmaStartTableRow & "<td>&nbsp;</td><td class=""ssHelpPanel"" colspan=2><div class=""ssHelpText"">" _
            & "All items that require shipping (in the shipping tab) will enable this section of the checkout. Checking the boxes next to the shipping sources you would like to include. These calculate shipping automatically. Create new shipping methods bases on the item price and weight with the Add-new-shipping-method link. Shipping methods will be sorted by the Alpha Sort field, followed by the UPS methods." _
            & "</div></td></tr>")
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "ShippingOriginatingZip", FieldTypeText, SaveAction, ""))
        Call content.Add(GetContent_GetHTMLElement(Copy, "From Zip", HelpCopy, False, False))
        '
        HelpCopy = ""
        If Not SaveAction Then
            FieldValue = main.GetSiteProperty("ShippingOriginatingCountry")
            If FieldValue = "" Then
                FieldValue = main.GetRecordID("countries", "United States")
                If FieldValue = "0" Then
                    'Set common = New commonClass
                    Call common.VerifyShippingStateCountry(main)
                End If
            End If
        Else
            FieldValue = main.GetStreamText("ShippingOriginatingCountry")
            Call main.SetSiteProperty("ShippingOriginatingCountry", FieldValue)
        End If
        Copy = main.GetFormInputSelect("ShippingOriginatingCountry", FieldValue, "Countries")
        Call content.Add(GetContent_GetHTMLElement(Copy, "From Country", HelpCopy, False, False))
        '
        HelpCopy = ""
        If Not SaveAction Then
            FieldValue = main.GetSiteProperty("ShippingDropOff", "5")
            If FieldValue = "" Then
                FieldValue = "5"
                Call main.SetSiteProperty("ShippingDropOff", FieldValue)
            End If
        Else
            FieldValue = main.GetStreamText("ShippingDropOff")
            Call main.SetSiteProperty("ShippingDropOff", FieldValue)
        End If
        Copy = main.GetFormInputSelectList("ShippingDropOff", FieldValue, "Regular Daily Pickup,On Call Air,One Time Pickup,Letter Center,Customer Counter")
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS DropOff/Pickup", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSNextDatAirAM", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Next Day Air Early AM", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSNextDatAir", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Next Day Air", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSNextDatAirSaver", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Next Day Air Saver", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPS2ndDayAirAM", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS 2nd Day Air AM", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPS2ndDayAir", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS 2nd Day Air", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPS3DaySelect", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS 3 Day Select", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSGround", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Ground", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSCanadaStandard", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Canada Standard", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSCWorldwideExpress", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Worldwide Express", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSCWorldwideExpressPlus", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Worldwide Express Plus", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSCWorldwideExpedited", FieldTypeBoolean, SaveAction, False))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Worldwide Expedited", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = (GetContent_GetPropertyElement(WidthWide, "AllowUPSCWorldwideSaver", FieldTypeBoolean, SaveAction, ""))
        Call content.Add(GetContent_GetHTMLElement(Copy, "UPS Worldwide Saver", HelpCopy, False, False))
        '
        HelpCopy = ""
        Copy = ""
        CCID = main.GetContentID("Order Ship Methods")
        cs = main.OpenCSContent("Order Ship Methods", "", "", False)
Dim itemId As Long
Dim itemName As String
Dim itemDescription As String
        Do While main.IsCSOK(cs)
            itemId = main.GetCSInteger(cs, "ID")
            itemName = main.GetCSText(cs, "Name")
            If Not main.GetCSBoolean(cs, "active") Then
                itemName = "(inactive) " & itemName
            End If
            itemDescription = main.GetCSText(cs, "Description")
            Copy = Copy _
                & "<tr>" _
                & "<td width=""20"" align=center><A href=""?cid=" & CCID & "&id=" & itemId & "&af=4""><img src=/cclib/images/IconContentEdit.gif border=0 alt=""Edit content"" valign=absmiddle></a></td>" _
                & "<td width=""99%""align=left>" & itemName & "<br><img src=/cclib/images/spacer width=150 height=1></td>" _
                & "</tr>"
            If (itemName <> itemDescription) And (itemDescription <> "") Then
                Copy = Copy _
                    & "<tr>" _
                    & "<td width=""20"">&nbsp;</td>" _
                    & "<td width=""99%""align=left>" & itemDescription & "</td>" _
                    & "</tr>"
            End If
            Copy = Copy & "<tr><td colspan=""2""><hr></td></tr>"
            Call main.NextCSRecord(cs)
        Loop
        Call main.CloseCS(cs)
        If Copy = "" Then
            Copy = Copy & "<tr><td width=100% align=left>There are no shipping methods defined</td></tr>"
        End If
        Copy = "" _
            & vbCrLf & vbTab & kmaStartTable(4, 0, 0) _
            & kmaIndent(Copy) _
            & vbCrLf & vbTab & "</table>" _
            & ""
        Copy = Copy _
            & vbCrLf & vbTab & kmaStartTable(4, 0, 0) _
            & vbCrLf & vbTab & "<tr><td width=100% align=left><a href=""?cid=" & CCID & "&af=4&aa=2&depth=1""><img src=/cclib/images/IconContentAdd.gif border=0></a>Add a new shipping method</td></tr>" _
            & vbCrLf & vbTab & "</table >" _
            & ""
        Call content.Add(GetContent_GetHTMLElement(Copy, "Other Shipping Methods", HelpCopy, False, False))
        '
        If main.SiteProperty_BuildVersion > "3.3.443" Then
            '
            ' Group Based Commerce Discounts
            '
            HelpCopy = "Use Group Policies to apply specific discounts and privileges based on membership to assigned groups."
            Call content.Add("" _
                & kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Group Policies</B>" _
                & main.GetHelpLink(0, "Group Policies", HelpCopy) _
                & kmaEndTableCell & kmaEndTableRow)
            Call content.Add("" _
                & kmaStartTableRow _
                & "<td>&nbsp;</td>" _
                & "<td class=""ssHelpPanel"" colspan=2>" _
                & "<div class=""ssHelpText"">Use Group Policies to apply special discounts and privileges available to only those people in specific groups. There are two kinds of Group Discounts, Custom Price Discount and Percentage Discount. When Allow Custom Price is checked, members of that group will be elegible for the 'custom price' in any item. When a non-zero Percentage Discount is entered below, anyone in that group received that discount off of the 'Regular Price'. This discount does not apply to 'Sales Price', or 'Custom Price'. The final price for each item, labeled with the caption 'Your Price', will be the lowest of the calculated price for each pricing strategy.</div>" _
                & "</td>" _
                & "</tr>" _
                & "")
            
Dim fieldValueInt As Long
Dim fieldValueText As String

            HelpCopy = "test test test"
            If SaveAction Then
                fieldValueInt = main.GetStreamInteger(rnContactUpdateGroupId)
                fieldValueText = main.GetRecordName("groups", fieldValueInt)
                Call main.SetSiteProperty("ShoppingCart Block Contact Update Group", fieldValueText)
            Else
                fieldValueText = main.GetSiteProperty("ShoppingCart Block Contact Update Group", "")
            End If
            If fieldValueText <> "" Then
                fieldValueInt = main.GetRecordID("groups", fieldValueText)
            End If
            Copy = main.GetFormInputSelect(rnContactUpdateGroupId, fieldValueInt, "Groups", , "None") _
                & "<div class=""ssHelpText"">If a group is selected, anyone logged in and in this group will not be able to update their contact information. Use this setting to protect user settings that may be controlled by another source such as a synchonization from another database.</div>"
            Call content.Add(GetContent_GetHTMLElement(Copy, "Block Contact Update Group", HelpCopy, False, False))
            '
            cs = main.OpenCSContent("Groups", , "ID,caption,Name", , , , "Caption,Name,AllowCommerceSpecialPrice,CommerceDiscount,AllowCommerceOrdering")
            'CS = Main.OpenCSContent("Groups", , "ID,caption,Name", , , , "Caption,Name,AllowCommerceSpecialPrice,CommerceDiscount,CommerceAdmin")
            If Not main.IsCSOK(cs) Then
                Call content.Add(kmaStartTableRow & "<td>&nbsp;</td><td colspan=2>There are no groups defined</td>" & kmaEndTableRow)
            Else
                Do While main.IsCSOK(cs)
                    RecordID = main.GetCSInteger(cs, "ID")
                    RowName = "Group" & RecordID
                    Copy = main.GetCSText(cs, "Caption")
                    If Copy = "" Then
                        Copy = main.GetCSText(cs, "Name")
                        If Copy = "" Then
                            Copy = "Group " & RecordID
                        End If
                    End If
                    If SaveAction Then
                        '
                        'CommerceAdmin = Main.GetStreamNumber(RowName & "Admin")
                        'Call Main.SetCS(CS, "CommerceAdmin", CommerceAdmin)
                        ''
                        AllowCommerceSpecialPrice = main.GetStreamBoolean(RowName & "Price")
                        Call main.SetCS(cs, "AllowCommerceSpecialPrice", AllowCommerceSpecialPrice)
                        '
                        CommerceDiscount = main.GetStreamNumber(RowName & "Discount")
                        Call main.SetCS(cs, "CommerceDiscount", CommerceDiscount)
                        '
                        AllowCommerceAddButton = main.GetStreamBoolean(RowName & "AllowAdd")
                        Call main.SetCS(cs, "AllowCommerceOrdering", AllowCommerceAddButton)
                    Else
                        AllowCommerceSpecialPrice = main.GetCSBoolean(cs, "AllowCommerceSpecialPrice")
                        CommerceDiscount = main.GetCSNumber(cs, "CommerceDiscount")
                        '
                        AllowCommerceAddButton = main.GetCSBoolean(cs, "AllowCommerceOrdering")
                        '
                        'CommerceAdmin = Main.GetCSBoolean(CS, "CommerceAdmin")
                    End If
                    FormElement = "" _
                        & "<div class="""">" & main.GetFormInputText(RowName & "Discount", CommerceDiscount, 1, 5) & "&nbsp;%&nbsp;Discount</div>" _
                        & "<div class="""">" & main.GetFormInputCheckBox(RowName & "Price", AllowCommerceSpecialPrice) & "Allow&nbsp;Custom&nbsp;Price</div>" _
                        & "<div class="""">" & main.GetFormInputCheckBox(RowName & "AllowAdd", AllowCommerceAddButton) & "Allow ordering (if ""Allow ordering for everyone"" is not checked above)</div>" _
                        & main.GetFormInputHidden(RowName, RecordID)
                    'FormElement = "" _
                    '    & Main.GetFormInputCheckBox(RowName & "Admin", CommerceAdmin) & "Commerce&nbsp;Administrator" _
                    '    & ",&nbsp;&nbsp;" & Main.GetFormInputCheckBox(RowName & "Price", AllowCommerceSpecialPrice) & "Allow&nbsp;Custom&nbsp;Price" _
                    '    & ",&nbsp;&nbsp;Discount&nbsp;" & Main.GetFormInputText(RowName & "Discount", CommerceDiscount, 1, 5) & "%" _
                    '    & Main.GetFormInputHidden(RowName, RecordID)
                    Call content.Add(GetContent_GetHTMLElement(FormElement, Copy))
                    RowCnt = RowCnt + 1
                    main.NextCSRecord (cs)
                Loop
                Call content.Add(main.GetFormInputHidden("GroupRowCnt", RowCnt))
            End If
            Call main.CloseCS(cs)
        End If
        '
        ' ----- sales tax
        '
        Call content.Add(kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Sales Tax</B>" & kmaEndTableCell & kmaEndTableRow)
        '
        Call content.Add(kmaStartTableRow & "<td>&nbsp;</td><td class=""ssHelpPanel"" colspan=2><div class=""ssHelpText"">" _
            & "Sales tax will be automatically included on the order based on the state of the shipping address. Individual state sales tax rates are set in the record for the state. The checkout process calculates the tax required but it is up to you to handle all state tax filings." _
            & "</div></td></tr>")
        '
        HelpCopy = ""
        Copy = "<a href=""?cid=" & main.GetContentID("states") & """>Edit States</a>"
        Call content.Add(GetContent_GetHTMLElement(Copy, "Edit State Tax Rates", HelpCopy, False, False))
        '
        ' ----- Inventory Management Properties
        '
        HelpCopy = "Use Inventory Management to notify you when catalog item quantities are below their individual limits."
        Call content.Add(kmaStartTableRow & "<td colspan=""3"" class=ccAdminEditSubHeader><B>Inventory Management</B>" & main.GetHelpLink(0, "Inventory Management", HelpCopy) & kmaEndTableCell & kmaEndTableRow)
        Call content.Add(kmaStartTableRow & "<td>&nbsp;</td><td class=""ssHelpPanel"" colspan=2><div class=""ssHelpText"">" _
            & "Use Inventory Management to notify you when an items quantity is low. For each item, enter the Inventory on hand, and Inventory low. Each time an item is purchased, the quantity on hand is decreased. When the quantity is below the Inventory Low number, a notification is sent by email." _
            & "</div></td></tr>")
        '
        HelpCopy = "Enable inventory management. When disabled, no inventory checks will be made."
        Copy = (GetContent_GetPropertyElement(WidthWide, "CatalogAllowInventory", FieldTypeBoolean, SaveAction, "false"))
        Call content.Add(GetContent_GetHTMLElement(Copy, "Allow Inventory Management", HelpCopy, False, False))
        '
        HelpCopy = "The email address where the low inventory notifications will be sent. The Email will be from the system's 'General Email From Address', in the Email Control screen"
        Copy = (GetContent_GetPropertyElement(WidthWide, "EmailCatalogInventoryAdmin", FieldTypeText, SaveAction, ""))
        Call content.Add(GetContent_GetHTMLElement(Copy, "Inventory Notification Email", HelpCopy, False, False))
        '
        ' end of settings
        '
        content.Add (AdminUI.EditTableClose)
        '
        s = "" _
            & vbCrLf & vbTab & "<div class=""shopSettingsCon"">" _
            & kmaIndent(content.Text) _
            & vbCrLf & vbTab & "</div>" _
            & ""
        '
        ButtonListLeft = ButtonCancel & "," & ButtonSave & "," & ButtonOK
    End If
    '
    Caption = "Shopping Cart Settings"
    Description = "This tool is used to control the Contensive Shopping Cart Add-on. Many of these controls are also available one at a time through Developer Site Properties."
    GetContent = AdminUI.GetBody(main, Caption, ButtonListLeft, ButtonListRight, False, False, Description, ContentSummary, 0, s)
    Exit Function
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetContent"): Resume Next
End Function
'
'
'
Private Function GetContent_GetHTMLElement(HTMLFieldString As String, Caption As String, Optional HelpMessage As String, Optional FieldRequired As Boolean, Optional AllowActiveEdit As Boolean, Optional ignore0 As String) As String
    On Error GoTo ErrorTrap
    '
    Dim FastString As FastString.FastStringClass
    Dim Copy As String
    '
    Set FastString = New FastStringClass
    Call FastString.Add("<tr>")
    Copy = Caption
    If Copy = "" Then
        Copy = "&nbsp;"
    End If
    FastString.Add ("<td class=ccAdminEditCaption style=""vertical-align:top;"">" & Copy)
    If main.VisitProperty_AllowHelpIcon Then
        If HelpMessage <> "" Then
            FastString.Add "&nbsp;" & main.GetHelpLink(0, Caption, HelpMessage)
        Else
            FastString.Add "&nbsp;<img src=/cclib/images/spacer.gif width=18 height=22>"
        End If
    End If
    FastString.Add ("<img src=/ccLib/Images/spacer.gif width=1 height=15></td>")
    '
    Copy = HTMLFieldString
    If Copy = "" Then
        Copy = "&nbsp;"
    End If
    Call FastString.Add("<td class=ccAdminEditField colspan=2 style=""padding-left:20px;"">" & Copy & "</td>")
    Call FastString.Add("</tr>")
    GetContent_GetHTMLElement = FastString.Text
    Set FastString = Nothing
    Exit Function
    '
ErrorTrap:
    Set FastString = Nothing
    Call HandleClassErrorAndResumeNext("GetContent_GetHTMLElement"): Resume Next
    End Function
'
'
'
Private Function GetContent_GetPropertyElement(Width As Long, Name As String, fieldType As Long, ProcessRequest As Boolean, DefaultValue As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CurrentValue As String
    '
    If ProcessRequest Then
        CurrentValue = main.GetStreamText(Name)
        Call main.SetSiteProperty(Name, CurrentValue)
    Else
        CurrentValue = main.GetSiteProperty(Name, DefaultValue)
        End If
    Select Case fieldType
        Case FieldTypeBoolean
            GetContent_GetPropertyElement = main.GetFormInputCheckBox(Name, CurrentValue)
        Case Else
            GetContent_GetPropertyElement = main.GetFormInputText(Name, CurrentValue, 1, Width)
        End Select
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetContent_GetPropertyElement"): Resume Next
    End Function
'
'
'
Private Function GetContent_GetCopyElement(Caption As String, CopyName As String, defaultCopy As String) As String
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Dim RecordID As Long
    Dim EditIcon As String
    Dim Copy As String
    '
    Const ContentName = "Copy Content"
    '
    cs = main.OpenCSContent(ContentName, "Name=" & KmaEncodeSQLText(CopyName))
    If main.IsCSOK(cs) Then
        RecordID = main.GetCSInteger(cs, "ID")
        Copy = main.GetCSText(cs, "copy")
    End If
    Call main.CloseCS(cs)
    '
    If RecordID = 0 Then
        Copy = defaultCopy
        cs = main.InsertCSRecord(ContentName)
        If main.IsCSOK(cs) Then
            RecordID = main.GetCSInteger(cs, "ID")
            Call main.SetCS(cs, "name", CopyName)
            Call main.SetCS(cs, "copy", defaultCopy)
        End If
        Call main.CloseCS(cs)
    End If
    '
    If RecordID <> 0 Then
        EditIcon = "<A href=""?cid=" & main.GetContentID(ContentName) & "&id=" & RecordID & "&af=4""><img src=/cclib/images/IconContentEdit.gif border=0 alt=""Edit content"" valign=absmiddle></a>"
    Else
        EditIcon = "<A href=""?cid=" & main.GetContentID(ContentName) & "&af=4&aa=2&depth=1&wc=" & kmaEncodeURL("name=" & CopyName) & """><img src=/cclib/images/IconContentEdit.gif border=0 alt=""Edit content"" valign=absmiddle></a>"
    End If
    If Copy = "" Then
        Copy = "&nbsp;"
    End If
'    '
'    GetContent_GetCopyElement = "" _
'        & kmaStartTable(4, 0, 0) & "<tr>" _
'        & "<td width=150 align=right>" & Caption & "<br><img src=/cclib/images/spacer width=150 height=1></td>" _
'        & "<td width=20 align=center>" & EditIcon & "</td>" _
'        & "<td>" & Copy & "&nbsp;</td>" _
'        & "</tr></table>"
    GetContent_GetCopyElement = "" _
        & vbCrLf & vbTab & "<div class=""ssEditIcon"">" & EditIcon & "</div>" _
        & vbCrLf & vbTab & "<div class=""ssEditCaption"">" & Caption & "</div>" _
        & vbCrLf & vbTab & "<div class=""ssEditValue"">" & Copy & "</div>" _
        & ""
    GetContent_GetCopyElement = "" _
        & vbCrLf & vbTab & "<div class=""ssEditRow"">" _
        & kmaIndent(GetContent_GetCopyElement) _
        & vbCrLf & vbTab & "</div>" _
        & ""
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetContent_GetCopyElement"): Resume Next
    '
End Function



