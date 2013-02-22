VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CmrcPricingClass"
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
    Call HandleErrorAndResumeNext("CmrcPricingClass", MethodName, Description, ErrorNumber)
End Sub
'
'==============================================================================
'   Calculate Order
'       Calculate the order details and order, set the orderdetail structure, and save them all
'       After Calucate Order, anything in the OrderDetails and Orders struct can be displayed
'==============================================================================
'
Public Sub GetPricingPolicies(main As Object, CommerceAccountMemberID As Long, AllowGroupPolicies As Boolean, MemberDiscount As Double, MemberAllowSpecialPricing As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim CatalogAllowSpecialPrice As Boolean
    Dim GroupDiscount As Double
    Dim SQLNow As String
    Dim cs As Long
    Dim RecordDiscount As Double
    '
    ' Determine MemberDiscount and MemberAllowSpecialPricing
    '
    MemberDiscount = 0
    MemberAllowSpecialPricing = False
    If AllowGroupPolicies And main.SiteProperty_BuildVersion >= "3.3.444" Then
        SQLNow = main.EncodeSQLDate(Now)
        SQL = "select G.AllowCommerceSpecialPrice,G.CommerceDiscount" _
            & " From ccGroups G left join ccMemberRules R on R.GroupID=G.ID" _
            & " where R.MemberID=" & CommerceAccountMemberID _
            & " and(G.active<>0)and((R.DateExpires is null)or(R.DateExpires>" & SQLNow & "))"
        cs = main.OpenCSSQL("default", SQL)
        Do While main.IsCSOK(cs)
            MemberAllowSpecialPricing = MemberAllowSpecialPricing Or main.GetCSBoolean(cs, "AllowCommerceSpecialPrice")
            GroupDiscount = main.GetCSNumber(cs, "CommerceDiscount")
            If (GroupDiscount > MemberDiscount) And (GroupDiscount <= 100) Then
                MemberDiscount = GroupDiscount
            End If
            main.NextCSRecord (cs)
        Loop
        Call main.CloseCS(cs)
    End If
    
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetPricingPolicies"): Resume Next
    End Sub
'
'==============================================================================
'
'==============================================================================
'
Public Sub GetRightsPolicies(main As Object, memberId As Long, AllowGroupPolicies As Boolean, ByRef return_isCommerceAdmin As Boolean)
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim CatalogAllowSpecialPrice As Boolean
    Dim GroupDiscount As Double
    Dim SQLNow As String
    Dim cs As Long
    Dim RecordDiscount As Double
    '
    ' Determine MemberDiscount and MemberAllowSpecialPricing
    '
    return_isCommerceAdmin = False
    If AllowGroupPolicies And main.SiteProperty_BuildVersion >= "3.3.444" Then
        return_isCommerceAdmin = main.IsAdmin()
    End If
    
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetRightsPolicies"): Resume Next
    End Sub
'
'
'
Public Function GetYourPrice(GoodPriceRegular As Boolean, PriceRegular As Double, GoodPriceSale As Boolean, PriceSale As Double, GoodPriceSpecial As Boolean, PriceSpecial As Double, DiscountFactor As Double)
    On Error GoTo ErrorTrap
    '
    ' Determine Your Price
    '
    If GoodPriceRegular Then
        GetYourPrice = DiscountFactor * PriceRegular
    End If
    If GoodPriceSale And PriceSale < GetYourPrice Then
        GetYourPrice = PriceSale
    End If
    If GoodPriceSpecial And PriceSpecial < GetYourPrice Then
        GetYourPrice = PriceSpecial
    End If

    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetYourPrice"): Resume Next
End Function
'
'
'
Public Function GetPricingCopy(main As Object, csItem As Long, CaptionYourPrice As String, CaptionRegularPrice As String, CaptionSalePrice As String, CaptionRetailPrice As String, CaptionSpecialPrice As String, AllowSpecialPrice As Boolean, CommerceAccountMemberID As Long, AllowGroupPolicies As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim PriceDetails As String
    Dim PriceDetailCnt As Long
    Dim s As String
    Dim PriceRegular As Double
    Dim PriceSpecial As Double
    Dim PriceSale As Double
    
    Dim PriceRegularText As String
    Dim PriceSpecialText As String
    Dim PriceSaleText As String
    
    Dim PriceYour As Double
    Dim PriceDiscount As Double
    Dim GoodPriceSale As Boolean
    Dim GoodPriceRegular As Boolean
    Dim GoodPriceSpecial  As Boolean
    Dim MemberDiscount As Double
    Dim MemberAllowSpecialPricing As Boolean
    Dim DiscountFactor As Double
    '
    PriceRegularText = main.GetCSText(csItem, "UnitPrice")
    PriceRegular = KmaEncodeNumber(PriceRegularText)
    '
    PriceSpecialText = main.GetCSText(csItem, "SpecialPrice")
    PriceSpecial = KmaEncodeNumber(PriceSpecialText)
    '
    PriceSaleText = main.GetCSText(csItem, "SalePrice")
    PriceSale = KmaEncodeNumber(PriceSaleText)
    '
    Call GetPricingPolicies(main, CommerceAccountMemberID, AllowGroupPolicies, MemberDiscount, MemberAllowSpecialPricing)
    DiscountFactor = (100# - MemberDiscount) / 100#
    '
    GoodPriceSale = (main.GetCSBoolean(csItem, "OnSale")) And (PriceSaleText <> "") And (CaptionSalePrice <> "")
    GoodPriceRegular = (PriceRegularText <> "") And (CaptionRegularPrice <> "")
    GoodPriceSpecial = MemberAllowSpecialPricing And (PriceSpecialText <> "") And (AllowSpecialPrice) And (CaptionSpecialPrice <> "")
    '
    ' Determine PriceYour
    '
    PriceYour = GetYourPrice(GoodPriceRegular, PriceRegular, GoodPriceSale, PriceSale, GoodPriceSpecial, PriceSpecial, DiscountFactor)
    '
    ' Lead with Your Price
    '
    s = s & "<div class=ccCatYourPrice>" & CaptionYourPrice & " " & FormatCurrency(PriceYour) & "</div>"
    PriceDetailCnt = 0
    PriceDetails = ""

    '
    ' Sale Price, if valid
    '
    If GoodPriceSale Then
        PriceDetailCnt = PriceDetailCnt + 1
        If PriceYour = PriceSale Then
            PriceDetails = PriceDetails & "<div class=ccCatSalePrice>" & CaptionSalePrice & " " & FormatCurrency(PriceSale) & "</div>"
        Else
            PriceDetails = PriceDetails & "<div class=ccCatSalePrice style=""text-decoration: line-through;"">" & CaptionSalePrice & " " & FormatCurrency(PriceSale) & "</div>"
        End If
    End If
    '
    ' Special Price, if valid
    '
    If GoodPriceSpecial Then
        PriceDetailCnt = PriceDetailCnt + 1
        If PriceYour = PriceSpecial Then
            PriceDetails = PriceDetails & "<div class=ccCatListCopy>" & CaptionSpecialPrice & " " & FormatCurrency(PriceSpecial) & "</div>"
        Else
            PriceDetails = PriceDetails & "<div class=ccCatListCopy style=""text-decoration: line-through;"">" & CaptionSpecialPrice & " " & FormatCurrency(PriceSpecial) & "</div>"
        End If
    End If
    '
    ' Regular Price, if valid
    '
    If GoodPriceRegular Then
        PriceDetailCnt = PriceDetailCnt + 1
        If PriceYour = PriceRegular Then
            PriceDetails = PriceDetails & "<div class=ccCatListCopy>" & CaptionRegularPrice & " " & FormatCurrency(PriceRegular) & "</div>"
        Else
            PriceDetails = PriceDetails & "<div class=ccCatListCopy style=""text-decoration: line-through;"">" & CaptionRegularPrice & " " & FormatCurrency(PriceRegular) & "</div>"
        End If
    End If
    '
    ' Discounted Regular Price, if valid
    '
    If GoodPriceRegular And (DiscountFactor <> 1#) Then
        PriceDetailCnt = PriceDetailCnt + 1
        PriceDiscount = PriceRegular * DiscountFactor
        If PriceYour = PriceDiscount Then
            PriceDetails = PriceDetails & "<div class=ccCatListCopy>Discounted " & CaptionRegularPrice & " " & FormatCurrency(PriceDiscount) & "</div>"
        Else
            PriceDetails = PriceDetails & "<div class=ccCatListCopy style=""text-decoration: line-through;"">Discounted " & CaptionRegularPrice & " " & FormatCurrency(PriceDiscount) & "</div>"
        End If
    End If
    '
    If PriceDetailCnt > 1 Then
        s = s & PriceDetails
    End If
    GetPricingCopy = s
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetPricingCopy"): Resume Next
End Function


