VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "commonClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
'
'
Friend Function completeOrder(main As Object, RefCode As String, ActionCode As String, AuthCode As String, OrderCCType As String, OrderCCNumber As String, OrderCCExpires As String, OrderCCName As String, OrderCCZip As String, CommerceAccountMemberID As Long, orderId As Long)
    On Error GoTo ErrorTrap
    '
    Dim CompletedOrdersCID As Long
    Dim cs As Long
    Dim orderVisitorId As Long
    Dim orderVisitId As Long
    '
    main.OrderAuthorize = False
    main.OrderAuthorized = True
    '
    Call main.AddGroupMember("Online Purchasers", CommerceAccountMemberID)
    Call UpdateInventory(main, orderId)
    Call UpdateSubscriptionGroups(main, orderId, CommerceAccountMemberID)
    Call SendNotice(main, orderId)
    Call SendReceipt(main, orderId)
    '
    ' Save the completed order
    '
    cs = main.OpenCSContentRecord("Orders", orderId)
    If main.IsCSOK(cs) Then
        orderVisitorId = main.GetCSInteger(cs, "visitorid")
        orderVisitId = main.GetCSInteger(cs, "visitid")
        'Call Main.SetCS(cs, "Complete", True)
        Call main.SetCS(cs, "OrderCompleted", True)
        Call main.SetCS(cs, "DateCompleted", Now())
        Call main.SetCS(cs, "CCAuthCode", AuthCode)
        Call main.SetCS(cs, "CCActionCode", ActionCode)
        Call main.SetCS(cs, "CCRefCode", RefCode)
        Call main.SetCS(cs, "MemberID", CommerceAccountMemberID)
        If kmaEncodeBoolean(main.GetSiteProperty("OrderSaveCreditCard", "0", True)) Then
            Call main.SetCS(cs, "CCType", OrderCCType)
            Call main.SetCS(cs, "CCNumber", OrderCCNumber)
            Call main.SetCS(cs, "CCName", OrderCCName)
            Call main.SetCS(cs, "CCExpires", OrderCCExpires)
            Call main.SetCS(cs, "CCZip", OrderCCZip)
        Else
            Call main.SetCS(cs, "CCType", "[not saved]")
            Call main.SetCS(cs, "CCNumber", "[not saved]")
            Call main.SetCS(cs, "CCName", "[not saved]")
            Call main.SetCS(cs, "CCExpires", "[not saved]")
            Call main.SetCS(cs, "CCZip", "[not saved]")
        End If
    End If
    Call main.CloseCS(cs)
    '
    CompletedOrdersCID = main.GetContentID("Orders Completed")
    If orderId > 0 And CompletedOrdersCID > 0 Then
        Call main.ExecuteSQL("Default", "Update Orders set ContentControlID=" & KmaEncodeSQLNumber(CompletedOrdersCID) & " where ID=" & orderId & ";")
    End If
    '
    ' Clear the visitors cart reference
    '
    main.VisitorOrderID = 0
    Call main.SaveVisitor
    Call main.ExecuteSQL("Default", "Update ccvisitors set orderid=0 where id=" & orderVisitorId & ";")
    '
    ' Clear the cart
    '
    orderId = 0
    main.OrderItemCount = 0
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("completeOrder"): Resume Next
End Function
'
'==============================================================================
' Update inventory
'   Go through the order and remove all items from the item tables.
'   Check the inventory level of all items in the current order,
'   if they are below the OrderQuantityLow, email the administrator
'   Aministrator shold then run a CatalogInventory() to create a full
'   Order (by source)
'==============================================================================
'
Private Sub UpdateInventory(main As Object, orderId As Long)
    On Error GoTo ErrorTrap
    '
    Dim quantity As Long
    Dim SQL As String
    Dim cs As Long
    Dim csItems As Long
    Dim ItemNumber As Long
    Dim ItemQuantityOnHand As Long
    Dim ItemQuantityLow As Long
    Dim Message As String
    Dim Recipient As String
    Dim CatalogAllowInventory As Boolean
    Dim fromAddress As String
    Dim itemId As Long
    Dim itemName As String
    '
    CatalogAllowInventory = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowInventory", "false"))
    If CatalogAllowInventory Then
        '
        ' go through all items in order and remove from inventory
        '
        cs = main.OpenCSContent("Order Details", "orderId=" & KmaEncodeSQLNumber(orderId))
        Do While main.IsCSOK(cs)
            itemId = main.GetCSInteger(cs, "itemid")
            quantity = main.GetCSInteger(cs, "quantity")
            csItems = main.OpenCSContentRecord("items", itemId)
            If main.IsCSOK(csItems) Then
                ItemQuantityOnHand = main.GetCSInteger(csItems, "QuantityOnhand")
                ItemQuantityOnHand = ItemQuantityOnHand - quantity
                ItemQuantityLow = main.GetCSInteger(csItems, "QuantityLow")
                itemName = main.GetCSText(csItems, "name")
                '
                ' update items table
                '
                SQL = "update items set QuantityOnhand=" & ItemQuantityOnHand & " where ID=" & itemId & ";"
                Call main.ExecuteSQL("", SQL)
                If ItemQuantityOnHand <= ItemQuantityLow Then
                    '
                    ' send low inventory notice
                    '
                    Message = "During Order processing, the following item was found to be running low. Please re-order" & vbCrLf
                    Message = Message & vbCrLf
                    Message = Message & "    " & itemName & vbCrLf
                    Message = Message & "    Inventory Onhand: " & ItemQuantityOnHand & vbCrLf
                    Message = Message & "    Inventory Limit: " & ItemQuantityLow & vbCrLf
                    Recipient = main.GetSiteProperty("EmailCatalogInventoryAdmin")
                    fromAddress = main.GetSiteProperty("EmailFromAddress")
                    Call main.SendEmail(Recipient, fromAddress, "Inventory Low Notice: " & itemName, Message, , False, False)
                End If
            End If
            Call main.CloseCS(csItems)
            Call main.NextCSRecord(cs)
        Loop
        Call main.CloseCS(cs)
    End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("UpdateInventory"): Resume Next
End Sub
'
'==============================================================================
' Update subscription groups
'==============================================================================
'
Private Sub UpdateSubscriptionGroups(main As Object, orderId As Long, CommerceAccountMemberID As Long)
    On Error GoTo ErrorTrap
    '
    Dim dateExpires As Date
    Dim csDetails As Long
    Dim csItems As Long
    Dim GroupID As Long
    Dim ExpirePeriod As Long
    Dim itemId As Long
    Dim cs As Long
    '
    ' go through all items add group member as needed
    '
    csDetails = main.OpenCSContent("Order Details", "orderId=" & KmaEncodeSQLNumber(orderId))
    Do While main.IsCSOK(csDetails)
        itemId = main.GetCSInteger(csDetails, "itemid")
        csItems = main.OpenCSContentRecord("items", itemId)
        If main.IsCSOK(csItems) Then
            GroupID = main.GetCSInteger(csItems, "GroupID")
            If GroupID <> 0 Then
                ExpirePeriod = main.GetCSInteger(csItems, "GroupExpirationPeriod")
                If ExpirePeriod = 0 Then
                    Call main.AddGroupMember(main.GetGroupByID(GroupID), CommerceAccountMemberID)
                Else
                    If main.IsGroupListMember(CStr(GroupID), CommerceAccountMemberID) Then
                        '
                        ' They are already a member, extend membership
                        '
                        cs = main.OpenCSContent("Member Rules", "((memberid=" & CommerceAccountMemberID & ")and(groupid=" & GroupID & "))")
                        If main.IsCSOK(cs) Then
                            dateExpires = main.GetCSDate(cs, "dateExpires")
                            Call main.SetCS(cs, "dateExpires", dateExpires + ExpirePeriod)
                        End If
                        Call main.CloseCS(cs)
                    Else
                        '
                        '
                        ' new membership
                        Call main.AddGroupMember(main.GetGroupByID(GroupID), CommerceAccountMemberID, Now() + ExpirePeriod)
                    End If
                End If
            End If
        End If
        Call main.CloseCS(csItems)
        Call main.NextCSRecord(csDetails)
    Loop
    Call main.CloseCS(csDetails)
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("UpdateSubscriptionGroups"): Resume Next
End Sub
'
'==============================================================================
'   SendNotice
'
'   send a notice to alert to a new order
'==============================================================================
'
Private Sub SendNotice(main As Object, orderId As Long)
    On Error GoTo ErrorTrap
    '
    Dim Recipient As String
    Dim Sender As String
    Dim Subject As String
    Dim MailServer As String
    Dim Message As String
    Dim Result As String
    Dim Mailer As String
    Dim defaultEmail As String
    '
    defaultEmail = main.SiteProperty_EmailAdmin
    If defaultEmail = "" Then
        defaultEmail = "info@" & main.ServerHost
        defaultEmail = Replace(defaultEmail, "@www.", "@")
    End If
    Recipient = main.GetSiteProperty("EmailOrderNotifyAddress", defaultEmail)
    If InStr(kmaEncodeText(Recipient), "@") = 0 Then
        Call HandleClassErrorAndResumeNext("SendNotice", "Order Process Notification email was not sent because EmailOrderNotifyAddress SiteProperty is not valid")
    Else
        Sender = main.GetSiteProperty("EmailOrderFromAddress", defaultEmail)
        If Sender = "" Then
            Sender = defaultEmail
        End If
        Subject = main.ServerDomain & " Online Order " & orderId & " Pending"
        Message = GetEmailReceiptSection(main, orderId)
        Call main.SendEmail(Recipient, Sender, Subject, Message, , False, True)
    End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("SendNotice"): Resume Next
    End Sub
'
'==============================================================================
'   SendReceipt
'
'   send a notice to alert to a new order
'==============================================================================
'
Private Sub SendReceipt(main As Object, orderId As Long)
    On Error GoTo ErrorTrap
    '
    Dim firstchr As String
    Dim pos As Long
    Dim toAddress As String
    Dim fromAddress As String
    Dim Subject As String
    Dim Message As String
    Dim Result As String
    Dim CSPointer As Long
    Dim DefaultEmailName As String
    Dim csOrder As Long
    Dim orderBillEmail As String
    Dim defaultEmail As String
    '
    csOrder = main.OpenCSContentRecord("Orders", orderId)
    If main.IsCSOK(csOrder) Then
        orderBillEmail = main.GetCSText(csOrder, "BillEmail")
    End If
    Call main.CloseCS(csOrder)
    '
    toAddress = orderBillEmail
    If InStr(toAddress, "@") <> 0 Then
        defaultEmail = main.SiteProperty_EmailAdmin
        If defaultEmail = "" Then
            defaultEmail = "info@" & main.ServerHost
            defaultEmail = Replace(defaultEmail, "@www.", "@")
        End If
        '
        fromAddress = main.GetSiteProperty("EmailOrderFromAddress", defaultEmail)
        If fromAddress = "" Then
            fromAddress = defaultEmail
        End If
        If InStr(kmaEncodeText(fromAddress), "@") = 0 Then
            Call HandleError("CommerceClass", "SendReceipt", KmaErrorInternal, App.EXEName, "Order Receipt email was not sent because EmailOrderFromAddress SiteProperty is not valid", False, True)
        Else
            DefaultEmailName = "Default Commerce Order Receipt"
            CSPointer = main.OpenCSContent_Internal("System Email", "name=" & KmaEncodeSQLText(DefaultEmailName))
            If main.IsCSOK(CSPointer) Then
                '
                ' get email
                '
                Subject = main.GetCSText(CSPointer, "Subject")
                Message = main.GetCS(CSPointer, "Copyfilename")
                '
                ' remove an initial date from the old default email body - was a legacy mistake
                '
                firstchr = Mid(Message, 1, 1)
                If firstchr <> "" Then
                    If IsNumeric(firstchr) Then
                        pos = InStr(1, Message, "<br>Dear Customer", vbTextCompare)
                        If pos > 0 Then
                            Message = Mid(Message, pos + 4)
                            Call main.SetCS(CSPointer, "copyfilename", Message)
                        End If
                    End If
                End If
            Else
                '
                ' No email found, create default
                '
                Call main.CloseCS(CSPointer)
                CSPointer = main.InsertCSContent("System Email")
                If Not main.IsCSOK(CSPointer) Then
                    '
                    ' Could not create system email
                    '
                    Call HandleClassErrorAndResumeNext("SendReceipt", "Could not create System Email [" & DefaultEmailName & "]")
                Else
                    Subject = "Thank you for your purchase at " & main.ServerDomain
                    Message = "" _
                        & "Dear Customer," _
                        & "<br>" _
                        & "<br>" & "Thank you for your purchase at " & main.ServerDomain & "." _
                        & "<br>" _
                        & "<br>" & "Your order will be processed promptly. Please keep this email as a record of your purchase. If you have any questions about your order, please return to " & main.ServerDomain & " for our phone number and call during regular business hours." _
                        & "<br>" _
                        & "<br>" & "We hope you enjoyed shopping at " & main.ServerDomain & ". If you have any comments or suggestions, please feel free to contact us at " & fromAddress & "." _
                        & "<br>" _
                        & "" _
                        & ""
                    Call main.SetCS(CSPointer, "active", 1)
                    Call main.SetCS(CSPointer, "Name", DefaultEmailName)
                    Call main.SetCS(CSPointer, "Subject", Subject)
                    Call main.SetCS(CSPointer, "fromAddress", fromAddress)
                    Call main.SetCS(CSPointer, "copyfilename", Message)
                End If
            End If
            Call main.CloseCS(CSPointer)
            '
            ' add the date and append the reciept details
            '
            Message = "" _
                & FormatDateTime(Now(), vbLongDate) _
                & "<br>" _
                & "<br>" & Message _
                & "<br>" & GetEmailReceiptSection(main, orderId)
            Call main.SendEmail(toAddress, fromAddress, Subject, Message, , False, True)
        End If
    End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("SendReceipt"): Resume Next
End Sub
'
'
'
Private Function GetEmailReceiptSection(main As Object, orderId As Long)
    On Error GoTo ErrorTrap
    '
    Dim csItem As Long
    Dim csDetails As Long
    Dim itemList As String
    Dim shippingRequired As Boolean
    Dim itemId As Long
    Dim itemReceiptCopy As String
    Dim detailOptions As String
    Dim ItemNumber As Long
    Dim itemName As String
    Dim ItemQuantity As Long
    Dim ItemUnitPrice As Double
    Dim ItemTotalPrice As Double
    Dim s As String
    Dim receiptCopy As String
    Dim cs As Long
    Dim isSupportedAddress2 As Boolean
    '
    ' create item list
    '
    isSupportedAddress2 = main.IsContentFieldSupported("people", "billAddress2")
    '
    csDetails = main.OpenCSContent("order details", "orderid=" & orderId)
    itemList = ""
    Do While main.IsCSOK(csDetails)
        itemId = main.GetCSInteger(csDetails, "itemid")
        csItem = main.OpenCSContentRecord("items", itemId)
        If main.IsCSOK(csItem) Then
            itemName = main.GetCSText(csItem, "name")
            If main.GetCSBoolean(csItem, "allowReceiptCopy") Then
                itemReceiptCopy = main.GetCSText(csItem, "receiptCopy")
            End If
            shippingRequired = shippingRequired Or main.GetCSBoolean(csItem, "ShippingRequired")
        End If
        Call main.CloseCS(csItem)
        detailOptions = main.GetCSText(csDetails, "options")
        If detailOptions <> "" Then
            itemName = itemName & ", " & detailOptions
        End If
        ItemQuantity = main.GetCSText(csDetails, "quantity")
        ItemUnitPrice = main.GetCSText(csDetails, "unitPrice")
        ItemTotalPrice = ItemQuantity * ItemUnitPrice
        itemList = itemList & "<br>    -----------------------------------------------------"
        itemList = itemList & "<br>    (x" & ItemQuantity & ") " & itemName & " @" & ItemUnitPrice & " each"
        receiptCopy = receiptCopy & itemReceiptCopy
        Call main.NextCSRecord(csDetails)
    Loop
    Call main.CloseCS(csDetails)
    itemList = itemList & "<br>    -----------------------------------------------------"
    '
    cs = main.OpenCSContentRecord("Orders", orderId)
    If main.IsCSOK(cs) Then
        s = "The following order was made " & Now()
        s = s & "<br>Order # " & orderId
        s = s & "<br>"
        s = s & "<br>Billing Contact"
        s = s & "<br>"
        s = s & "<br>           " & main.GetCSText(cs, "BillName")
        s = s & "<br>    Email: " & main.GetCSText(cs, "BillEmail")
        s = s & "<br>    Phone: " & main.GetCSText(cs, "BillPhone")
        s = s & "<br>      Fax: " & main.GetCSText(cs, "BillFax")
        s = s & "<br>  Company: " & main.GetCSText(cs, "BillCompany")
        s = s & "<br>  Address: " & main.GetCSText(cs, "BillAddress")
        If isSupportedAddress2 Then
            s = s & "<br>  Address: " & main.GetCSText(cs, "BillAddress2")
        End If
        s = s & "<br>     City: " & main.GetCSText(cs, "BillCity")
        s = s & "<br>    State: " & main.GetCSText(cs, "BillState")
        s = s & "<br> Zip Code: " & main.GetCSText(cs, "BillZip")
        s = s & "<br>  Country: " & main.GetCSText(cs, "BillCountry")
        s = s & "<br>"
        If Not shippingRequired Then
            s = s & "<br>No shipping is required"
            s = s & "<br>"
        Else
            s = s & "<br>Shipping"
            s = s & "<br>"
            s = s & "<br>    " & main.GetCSText(cs, "ShipName")
            If main.GetCSText(cs, "ShipCompany") <> "" Then
                s = s & "<br>    " & main.GetCSText(cs, "ShipCompany")
            End If
            s = s & "<br>    " & main.GetCSText(cs, "ShipAddress")
            If isSupportedAddress2 Then
                s = s & "<br>    " & main.GetCSText(cs, "ShipAddress2")
            End If
            s = s & "<br>    " & main.GetCSText(cs, "ShipCity") & ", " & main.GetCSText(cs, "ShipState") & " " & main.GetCSText(cs, "ShipZip")
            s = s & "<br>    " & main.GetCSText(cs, "ShipCountry")
            s = s & "<br>    " & main.GetCSText(cs, "ShipPhone")
            s = s & "<br>"
            s = s & "<br>Shipping Method: " & main.GetCSText(cs, "ShipMethod")
        End If
        s = s & "<br>"
        s = s & "<br>Order"
        s = s & "<br>"
        s = s & itemList
        s = s & "<br>    Ship=" & FormatCurrency(main.GetCSNumber(cs, "ShipCharge"))
        s = s & "<br>    Tax=" & FormatCurrency(main.GetCSNumber(cs, "TaxCharge"))
        s = s & "<br>    Total=" & FormatCurrency(main.GetCSNumber(cs, "TotalCharge"))
        s = s & "<br>"
        s = s & "<br>Payment Method: " & main.GetCSText(cs, "PayMethod")
        '
        If receiptCopy <> "" Then
            s = "" _
                & "<div>    -----------------------------------------------------</div>" _
                & "<div>" & receiptCopy & "</div>" _
                & "<div>    -----------------------------------------------------</div>" _
                & s _
                & ""
        End If
    End If
    Call main.CloseCS(cs)
    GetEmailReceiptSection = s
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetEmailReceiptSection")
End Function
'
' Get the Menu for FormInputHTML
'
Private Sub VerifyCountry(main As Object, Name As String, Abbreviation As String)
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Dim Active As Boolean
    '
    cs = main.OpenCSContent("Countries", "name=" & KmaEncodeSQLText(Name))
    If Not main.IsCSOK(cs) Then
        Call main.CloseCS(cs)
        cs = main.InsertCSRecord("Countries")
        If main.IsCSOK(cs) Then
            Call main.SetCSField(cs, "ACTIVE", True)
        End If
    End If
    If main.IsCSOK(cs) Then
        Call main.SetCSField(cs, "NAME", Name)
        Call main.SetCSField(cs, "Abbreviation", Abbreviation)
        If LCase(Name) = "united states" Then
            Call main.SetCS(cs, "DomesticShipping", "1")
        End If
    End If
    Call main.CloseCS(cs)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("VerifyCountry")
End Sub
'
'=========================================================================================
'
'=========================================================================================
'
Private Sub VerifyState(main As Object, Name As String, Abbreviation As String, SaleTax As Double, CountryID As Long, FIPSState As String)
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Const ContentName = "States"
    '
    cs = main.OpenCSContent(ContentName, "name=" & KmaEncodeSQLText(Name), , False)
    If Not main.IsCSOK(cs) Then
        Call main.CloseCS(cs)
        cs = main.InsertCSRecord(ContentName)
        Call main.SetCSField(cs, "NAME", Name)
        Call main.SetCSField(cs, "ACTIVE", True)
        Call main.SetCSField(cs, "Abbreviation", Abbreviation)
        Call main.SetCSField(cs, "CountryID", CountryID)
        Call main.SetCSField(cs, "FIPSState", FIPSState)
    End If
    Call main.CloseCS(cs)
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("VerifyState")
End Sub
'
'=========================================================================================
'
'=========================================================================================
'
Friend Sub VerifyShippingStateCountry(main As Object)
    On Error GoTo ErrorTrap
    '
    Dim CountryID As Long
    '
    Call VerifyCountry(main, "United States", "US")
    CountryID = main.GetRecordID("Countries", "United States")
    '
    Call VerifyState(main, "Alaska", "AK", 0#, CountryID, "")
    Call VerifyState(main, "Alabama", "AL", 0#, CountryID, "")
    Call VerifyState(main, "Arizona", "AZ", 0#, CountryID, "")
    Call VerifyState(main, "Arkansas", "AR", 0#, CountryID, "")
    Call VerifyState(main, "California", "CA", 0#, CountryID, "")
    Call VerifyState(main, "Connecticut", "CT", 0#, CountryID, "")
    Call VerifyState(main, "Colorado", "CO", 0#, CountryID, "")
    Call VerifyState(main, "Delaware", "DE", 0#, CountryID, "")
    Call VerifyState(main, "District of Columbia", "DC", 0#, CountryID, "")
    Call VerifyState(main, "Florida", "FL", 0#, CountryID, "")
    Call VerifyState(main, "Georgia", "GA", 0#, CountryID, "")
    
    Call VerifyState(main, "Hawaii", "HI", 0#, CountryID, "")
    Call VerifyState(main, "Idaho", "ID", 0#, CountryID, "")
    Call VerifyState(main, "Illinois", "IL", 0#, CountryID, "")
    Call VerifyState(main, "Indiana", "IN", 0#, CountryID, "")
    Call VerifyState(main, "Iowa", "IA", 0#, CountryID, "")
    Call VerifyState(main, "Kansas", "KS", 0#, CountryID, "")
    Call VerifyState(main, "Kentucky", "KY", 0#, CountryID, "")
    Call VerifyState(main, "Louisiana", "LA", 0#, CountryID, "")
    Call VerifyState(main, "Massachusetts", "MA", 0#, CountryID, "")
    Call VerifyState(main, "Maine", "ME", 0#, CountryID, "")
    
    Call VerifyState(main, "Maryland", "MD", 0#, CountryID, "")
    Call VerifyState(main, "Michigan", "MI", 0#, CountryID, "")
    Call VerifyState(main, "Minnesota", "MN", 0#, CountryID, "")
    Call VerifyState(main, "Missouri", "MO", 0#, CountryID, "")
    Call VerifyState(main, "Mississippi", "MS", 0#, CountryID, "")
    Call VerifyState(main, "Montana", "MT", 0#, CountryID, "")
    Call VerifyState(main, "North Carolina", "NC", 0#, CountryID, "")
    Call VerifyState(main, "Nebraska", "NE", 0#, CountryID, "")
    Call VerifyState(main, "New Hampshire", "NH", 0#, CountryID, "")
    Call VerifyState(main, "New Mexico", "NM", 0#, CountryID, "")
    
    Call VerifyState(main, "New Jersey", "NJ", 0#, CountryID, "")
    Call VerifyState(main, "New York", "NY", 0#, CountryID, "")
    Call VerifyState(main, "Nevada", "NV", 0#, CountryID, "")
    Call VerifyState(main, "North Dakota", "ND", 0#, CountryID, "")
    Call VerifyState(main, "Ohio", "OH", 0#, CountryID, "")
    Call VerifyState(main, "Oklahoma", "OK", 0#, CountryID, "")
    Call VerifyState(main, "Oregon", "OR", 0#, CountryID, "")
    Call VerifyState(main, "Pennsylvania", "PA", 0#, CountryID, "")
    Call VerifyState(main, "Rhode Island", "RI", 0#, CountryID, "")
    Call VerifyState(main, "South Carolina", "SC", 0#, CountryID, "")
    
    Call VerifyState(main, "South Dakota", "SD", 0#, CountryID, "")
    Call VerifyState(main, "Tennessee", "TN", 0#, CountryID, "")
    Call VerifyState(main, "Texas", "TX", 0#, CountryID, "")
    Call VerifyState(main, "Utah", "UT", 0#, CountryID, "")
    Call VerifyState(main, "Vermont", "VT", 0#, CountryID, "")
    Call VerifyState(main, "Virginia", "VA", 0.045, CountryID, "")
    Call VerifyState(main, "Washington", "WA", 0#, CountryID, "")
    Call VerifyState(main, "Wisconsin", "WI", 0#, CountryID, "")
    Call VerifyState(main, "West Virginia", "WV", 0#, CountryID, "")
    Call VerifyState(main, "Wyoming", "WY", 0#, CountryID, "")
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("VerifyShippingStateCountry")
    Resume Next
End Sub


'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("CommerceClass", MethodName, Description, ErrorNumber)
End Sub
'
Friend Function getAllowOrdering(main As Object, CommerceAccountMemberID As Long) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim SQLNow As String
    '
    If kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowOrdering", True)) Then
        getAllowOrdering = True
    Else
        SQLNow = main.EncodeSQLDate(Now)
        SQL = "select G.ID" _
            & " From ccGroups G left join ccMemberRules R on R.GroupID=G.ID" _
            & " where R.MemberID=" & CommerceAccountMemberID _
            & " and(G.active<>0)and((R.DateExpires is null)or(R.DateExpires>" & SQLNow & "))" _
            & " and (AllowCommerceOrdering=1)"
        cs = main.OpenCSSQL("default", SQL)
        getAllowOrdering = main.IsCSOK(cs)
        Call main.CloseCS(cs)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getAllowOrdering")
End Function
'
Friend Function isGroupMember(main As Object, groupName As String, userId As Long) As Boolean
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim SQLNow As String
    '
    isGroupMember = False
    If main.IsAuthenticated Then
        SQLNow = main.EncodeSQLDate(Now)
        SQL = "select r.ID" _
            & " from ccGroups g" _
            & " left join ccMemberRules r on r.GroupID=g.ID" _
            & " where r.MemberID=" & userId _
            & " and(g.active<>0)and((r.DateExpires is null)or(r.DateExpires>" & SQLNow & "))" _
            & " and (g.name=" & KmaEncodeSQLText(groupName) & ")"
        cs = main.OpenCSSQL("default", SQL)
        isGroupMember = main.IsCSOK(cs)
        Call main.CloseCS(cs)
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getAllowOrdering")
End Function

