VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "payPalConfirmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
Private mainobj As Object
Private Csv As Object
'
'===========================================================================================================
'   copied from meetingSmart
'===========================================================================================================
'
Public Function Execute(CsvObject As Object, mainObject As Object, optionString As String, FilterInput As String)
    '
    Set mainobj = mainObject
    'Set MainObject = CsvObject
    '
    Dim memberId As Long
    Dim orderId As Long
    'Dim DiscountCode As String
    'Dim Email As String
    'Dim AttendeeID As Long
    'Dim AttendeeTypeID As Long
    'Dim itinerary As String
    'Dim notificationGroup As String
    'Dim meetingName As String
    'Dim meetingEmailSubject As String
    'Dim meetingEmailBody As String
    'Dim meetingEmailBodyTemplate As String
    'Dim meetingEarlyRegistrationDate As Date
    'Dim meetingEmailFromAddress As String
    'Dim CSPointer As Long
    'Dim meetingID As Long
    'Dim adminObj As New AdminClass
    'Dim okToCompleteRegistration As Boolean
    'Dim registrationVisitID As Long
    Dim visitId As Long
    Dim PayPalAuthCode As String
    Dim PayPalPaymentStatus As String
    Dim PayPalPendingReason As String
    Dim PayPalPaymentType As String
    Dim PayPalShipName As String
    Dim PayPalShipAddress As String
    Dim PayPalShipCity As String
    Dim PayPalShipState As String
    Dim PayPalShipZip As String
    Dim PayPalBillEmail As String
    Dim echoString As String
    Dim echoVar() As String
    Dim amountPaid As Double
    Dim amountDue As Double
    'Dim amountApplied As Double
    'Dim attendeeAmountDue As Double
    'Dim attendeeAmountPaid As Double
    'Dim attendeeMeetingCharge As Double
    'Dim attendeeSessionCharge As Double
    Dim common As New commonClass
    Dim csOrder As Long
    '
    '-------------------------------------
    ' This is from the paypal confirm in shopping cart
    '
    echoString = mainobj.GetStreamText("echo")
    If echoString = "" Then
        Call AppendLogFile("PayPal Confirmation hit received but item_name was empty. It should be the echoString with visitId and AmountPaid.")
    Else
        echoVar = Split(echoString, ",")
        If UBound(echoVar) <> 1 Then
            Call AppendLogFile("PayPal Confirmation hit received but echoString was invalid [" & echoString & "]. It should be the echoString with visitId and AmountPaid.")
        Else
            visitId = kmaEncodeInteger(echoVar(0))
            amountPaid = KmaEncodeNumber(echoVar(1))
            If visitId = 0 Then
                Call AppendLogFile("PayPal Confirmation hit received but echostring returned an invalid visitId (first arg of echostring). [" & echoString & "]")
            ElseIf amountPaid = 0 Then
                Call AppendLogFile("PayPal Confirmation hit received but echostring returned an invalid amountPaid (second arg of echostring). [" & echoString & "]")
            Else
                '
                ' good visitid and good amountPaid
                '
                PayPalAuthCode = mainobj.GetStreamText("txn_id")
                PayPalPaymentStatus = mainobj.GetStreamText("payment_status")
                PayPalPendingReason = mainobj.GetStreamText("pending_reason")
                PayPalPaymentType = "PayPal " & mainobj.GetStreamText("payment_type")
                '
                PayPalShipName = mainobj.GetStreamText("first_name") & " " & mainobj.GetStreamText("last_name")
                PayPalShipAddress = mainobj.GetStreamText("address_street")
                PayPalShipCity = mainobj.GetStreamText("address_city")
                PayPalShipState = mainobj.GetStreamText("address_state")
                PayPalShipZip = mainobj.GetStreamText("address_zip")
                PayPalBillEmail = mainobj.GetStreamText("payer_email")
                '
                orderId = 0
                'csOrder = mainobj.OpenCSContent("orders", "(visitid=" & visitId & ")", "id desc")
                csOrder = mainobj.OpenCSContent("orders", "(visitid=" & visitId & ")and((complete is null)or(complete=0))and(totalCharge=" & KmaEncodeSQLNumber(amountPaid) & ")", "id desc")
                If Not mainobj.IsCSOK(csOrder) Then
                    Call AppendLogFile("PayPal Confirmation hit received but the visit returned in the echostring could not be opened (second arg of echostring). [" & echoString & "]")
                Else
                    orderId = mainobj.GetCSInteger(csOrder, "id")
                    memberId = mainobj.GetCSInteger(csOrder, "memberid")
                End If
                Call mainobj.CloseCS(csOrder)
                If orderId <> 0 Then
                    'Set common = New commonClass
                    Call common.completeOrder(mainobj, PayPalPaymentStatus, PayPalPendingReason, PayPalAuthCode, "", "", "", "", "", memberId, orderId)
                End If
                ''
                '' Finish paypal close
                ''
                'mainobj.BlockClosePageCopyright = True
                'mainobj.BlockClosePageLink = True
                'Call mainobj.GetClosePage(False, False)
            End If
        End If
    End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("PayPalConfirm", "Execute", Err.Number, Err.Source, Err.Description, True, False)
End Function


