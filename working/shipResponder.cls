VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "shipResponder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private main As Object
Private Csv As Object

Public Function Execute(CsvObject As Object, mainObject As Object, optionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    Set Csv = CsvObject
    Set main = mainObject
    
    Dim stream As String
    Dim curName As String
    Dim curWeight As Double
    Dim curCharge As Double
    Dim shipZip As String
    Dim shipCountry As String
    
    curName = main.GetStreamText("shipMethod")
    curWeight = main.GetStreamNumber("shipWeight")
    curCharge = main.GetStreamNumber("shipCharge")
    shipCountry = main.GetStreamText("shipCountry")
    shipZip = main.GetStreamText("shipZip")
    Execute = GetShippingSelect(curName, curWeight, curCharge, shipZip, shipCountry)
    Exit Function
    
ErrorTrap:
    Call HandleError("shipResponder", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function

Private Function GetShippingSelect(CurrentName As String, orderShipWeight As Double, OrderItemCharge As Double, orderShipZip As String, orderShipCountry As String)
    On Error GoTo ErrorTrap
    '
    Dim msg As String
    Dim ErrorMessage As String
    Dim ShippingOriginatingCountry As String
    Dim ShippingOriginatingCountryAbbr As String
    Dim ShippingOriginatingZip As String
    Dim cs As String
    Dim Name As String
    Dim ChargeValue As Double
    Dim Time As String
    Dim Label As String
    Dim Domestic As Boolean
    Dim International As Boolean
    Dim ShippingDaysMin As Long
    Dim ShippingDaysMax As Long
    Dim AllowPriceLevelFreeShipping As Boolean
    Dim SiteAllowPriceLevelFreeShipping As Boolean
    Dim SiteFreeShippingPrice As Double
    Dim upsVal As String
    Dim appURL As String
    Dim objXML As New XMLHTTP60
    Dim selectOptions As String
    Dim criteria As String
    Dim orderIsDomestic As Boolean
    '
    criteria = "(name<>'')and(name is not null)"
    If (LCase(orderShipCountry) = "united states") Or (LCase(orderShipCountry) = "us") Or (LCase(orderShipCountry) = "usa") Or (LCase(orderShipCountry) = "united states of america") Then
        orderIsDomestic = True
        If Len(orderShipZip) > 5 Then
            orderShipZip = Left(orderShipZip, 5)
        End If
        criteria = criteria & "and(Domestic=1)"
    Else
        orderIsDomestic = False
        criteria = criteria & "and(International=1)"
    End If
    cs = main.OpenCSContent("Order Ship Methods", criteria)
    If main.IsCSOK(cs) Then
        SiteAllowPriceLevelFreeShipping = (main.SiteProperty_BuildVersion > "3.3.580") And kmaEncodeBoolean(main.GetSiteProperty("OrderAllowFreeShippingPrice", "0"))
        SiteFreeShippingPrice = KmaEncodeNumber(main.GetSiteProperty("OrderFreeShippingPrice", "0"))
        Do While main.IsCSOK(cs)
            Label = ""
            If SiteAllowPriceLevelFreeShipping Then
                AllowPriceLevelFreeShipping = main.GetCSInteger(cs, "AllowPriceLevelFreeShipping")
            End If
            Name = main.GetCSText(cs, "Name")
            ChargeValue = CalcShippingCharge(main.GetCSNumber(cs, "BaseCharge"), main.GetCSNumber(cs, "ChargePerWeight"), main.GetCSNumber(cs, "ChargePerCharge"), orderShipWeight, OrderItemCharge, AllowPriceLevelFreeShipping, SiteFreeShippingPrice)
            '
            If ChargeValue < 0.01 Then
                Label = "FREE SHIPPING"
            Else
                Label = "est " & FormatCurrency(CStr(ChargeValue))
            End If
            '
            ShippingDaysMin = main.GetCSInteger(cs, "DaysMin")
            ShippingDaysMax = main.GetCSInteger(cs, "DaysMax")
            If (ShippingDaysMin > 0) Or (ShippingDaysMax > 0) Then
                Label = Label & ", " & ShippingDaysMin
                If ShippingDaysMin = ShippingDaysMax Then
                    If ShippingDaysMin = 1 Then
                        Label = Label & " business day"
                    Else
                        Label = Label & " business days"
                    End If
                Else
                    Label = Label & " to " & ShippingDaysMax & " business days"
                End If
            End If
            '
            Domestic = main.GetCSBoolean(cs, "Domestic")
            International = main.GetCSBoolean(cs, "International")
            If Domestic And International Then
                Label = Label & ", Domestic and International"
            ElseIf Domestic Then
                Label = Label & ", Domestic only"
            ElseIf Domestic Then
                Label = Label & ", International only"
            End If
            '
            Label = Name & " (" & Label & ")"
            '
            If Name = CurrentName Then
                selectOptions = selectOptions & "<option value=""" & Name & """ selected>" & Label & "</option>"
            Else
                selectOptions = selectOptions & "<option value=""" & Name & """>" & Label & "</option>"
            End If
            '
            main.NextCSRecord (cs)
        Loop
    End If
    Call main.CloseCS(cs)
    ShippingOriginatingZip = main.GetSiteProperty("ShippingOriginatingZip", "20175")
    If ShippingOriginatingZip = "" Then
        ShippingOriginatingZip = "20175"
        Call main.SetSiteProperty("ShippingOriginatingZip", ShippingOriginatingZip)
    End If
    ShippingOriginatingCountry = main.GetSiteProperty("ShippingOriginatingCountry", "")
    If ShippingOriginatingCountry = "" Or (Not IsNumeric(ShippingOriginatingCountry)) Then
        ShippingOriginatingCountry = main.GetRecordID("countries", "United States")
        Call main.SetSiteProperty("ShippingOriginatingCountry", ShippingOriginatingCountry)
    End If
    ShippingOriginatingCountryAbbr = getCountryAbbreviation(ShippingOriginatingCountry)
    If ShippingOriginatingCountryAbbr = "" Then
        ShippingOriginatingCountryAbbr = "US"
    End If
    If orderIsDomestic Then
        selectOptions = selectOptions & getShipOptionUPS("UPS Next Day Air Early AM", "AllowUPSNextDatAirAM", "1DM", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS Next Day Air", "AllowUPSNextDatAir", "1DA", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS Next Day Air Saver", "AllowUPSNextDatAirSaver", "1DP", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS 2nd Day Air AM", "AllowUPS2ndDayAirAM", "2DM", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS 2nd Day Air", "AllowUPS2ndDayAir", "2DA", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS 3 Day Select", "AllowUPS3DaySelect", "3DS", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS Ground", "AllowUPSGround", "GND", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
    Else
        selectOptions = selectOptions & getShipOptionUPS("UPS Canada Standard", "AllowUPSCanadaStandard", "STD", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS Worldwide Express", "AllowUPSCWorldwideExpress", "XPR", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS Worldwide Express Plus", "AllowUPSCWorldwideExpressPlus", "XDM", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS Worldwide Expedited", "AllowUPSCWorldwideExpedited", "XPD", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
        selectOptions = selectOptions & getShipOptionUPS("UPS Worldwide Express Saver", "AllowUPSCWorldwideSaver", "WXS", CurrentName, CStr(orderShipWeight), orderShipZip, orderShipCountry, ShippingOriginatingZip, ShippingOriginatingCountryAbbr, ErrorMessage)
    End If
    '
    If selectOptions = "" Then
        GetShippingSelect = "<p>No shipping options are available. Please verify your shipping country and zip/postal code.</p>"
        'If orderShipZip = "" Then
        '    GetShippingSelect = "<p>No shipping options are available. Please verify your shipping country and postal code.</p>"
        'ElseIf orderShipCountry = "" Then
        '    GetShippingSelect = "<p>No shipping options are available. Please verify your shipping country and zip/postal code.</p>"
        'Else
        '    GetShippingSelect = "<p>No shipping options are available. Please verify your shipping country and zip/postal code.</p>"
        'End If
    Else
        GetShippingSelect = "" _
            & "<select name=""orderShipMethod"">" _
            & "<option value="""">Select One</option>" _
            & selectOptions _
            & "</select>"
    End If
    If main.IsDeveloper() And (ErrorMessage <> "") Then
        msg = main.GetAdminHintWrapper(ErrorMessage)
        msg = Replace(msg, "administrator", "Developer", , , vbTextCompare)
        GetShippingSelect = GetShippingSelect & msg
    End If

    Exit Function
ErrorTrap:
    Call HandleError("shipResponder", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function

Private Function getShipOptionUPS(Caption, propertyName, serviceCode, curVal, curWeight As String, curZip As String, curCountry As String, ShippingOriginatingZip As String, ShippingOriginatingCountryAbbr As String, ByRef return_ErrorMsg As String) As String
    On Error GoTo ErrorTrap

    Dim ConsigneeCountry As String
    Dim stream As String
    Dim appURL As String
    Dim objXML As New XMLHTTP60
    Dim respVal As String
    Dim respArr() As String
    Dim upsPrice As String
    Dim rateChart As String

    If kmaEncodeBoolean(main.GetSiteProperty(propertyName)) Then
        If curZip <> "" Then
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
            ConsigneeCountry = getCountryAbbreviation(curCountry)
            '
            ' fix legacy abbreviation, must be ISO
            '
            If LCase(ConsigneeCountry) = "usa" Then
                ConsigneeCountry = "US"
            End If
            If LCase(ShippingOriginatingCountryAbbr) = "usa" Then
                ShippingOriginatingCountryAbbr = "US"
            End If
            
            'appURL = "http://www.ups.com/using/services/rave/qcost_dss.cgi?AppVersion=1.2&AcceptUPSLicenseAgreement=YES&ResponseType=application/x-ups-rss&ActionCode=3&RateChart=Regular%20Daily%20Pickup&ResidentialInd=yes&PackagingType=00&"
            appURL = "http://www.ups.com/using/services/rave/qcost_dss.cgi?AppVersion=1.2&AcceptUPSLicenseAgreement=YES&ResponseType=application/x-ups-rss&ActionCode=3&ResidentialInd=yes&PackagingType=00&"
            appURL = appURL & "&RateChart=" & rateChart
            appURL = appURL & "&ServiceLevelCode=" & serviceCode
            appURL = appURL & "&ShipperPostalCode=" & ShippingOriginatingZip
            appURL = appURL & "&14_origCountry=" & ShippingOriginatingCountryAbbr
            appURL = appURL & "&ConsigneePostalCode=" & curZip
            appURL = appURL & "&ConsigneeCountry=" & ConsigneeCountry
            appURL = appURL & "&PackageActualWeight=" & curWeight
        
            Call objXML.Open("POST", appURL)
            Call objXML.send
            respVal = objXML.responseText
            
            If (InStr(1, respVal, "Success") <> 0) And (InStr(1, respVal, "%") <> 0) Then
                respArr = Split(respVal, "%")
                upsPrice = FormatCurrency(respArr(14), 2)
                
                stream = "<option value=""" & Caption & """ "
                If curVal = Caption Then
                    stream = stream & "selected "
                End If
                stream = stream & ">" & Caption & " - " & upsPrice & "</option>"
            Else
                return_ErrorMsg = return_ErrorMsg & "<div>UPS returned an error for shipping method [" & propertyName & "], [" & kmaEncodeHTML(respVal) & "]</div>"
            End If
        End If
    End If
    
    getShipOptionUPS = stream

    Exit Function

ErrorTrap:
    Call HandleError("shipResponder", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function

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
    Call HandleError("shipResponder", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
'
'==============================================================================
'   Calculate Shipping Charge for a given shipping method
'       orderShipWeight and OrderCharge must be calculated already
'       returns the charge in dollars
'==============================================================================
'
Private Function CalcShippingCharge(BaseCharge As Double, ChargePerWeight As Double, ChargePerCharge As Double, orderShipWeight As Double, OrderCharge As Double, AllowPriceLevelFreeShipping As Boolean, FreeShippingPrice As Double) As Double
    On Error GoTo ErrorTrap
    
    If AllowPriceLevelFreeShipping And (OrderCharge > FreeShippingPrice) Then
        CalcShippingCharge = 0
    Else
        CalcShippingCharge = BaseCharge + (orderShipWeight * ChargePerWeight) + (OrderCharge * ChargePerCharge)
    End If
    
    Exit Function

ErrorTrap:
    Call HandleError("shipResponder", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
