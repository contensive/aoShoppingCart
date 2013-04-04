VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CatalogClass2"
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
' do not use csv - this class has publics called by other classes. Since the init only initializes main, csv is left
' nothing, and csv. calls will fail.
'Private Csv As Object
'
' The member whos account will be used for pricing, etc - set to main.memberid during init
'
Public CommerceAccountMemberID As Long
Public AllowGroupPolicies As Boolean
Public IsMemberCommerceAdmin As Boolean
'
' ----- global scope variables
'
Private ClassInitialized As Boolean
Private ItemsContentName As String
Private CategoryContentName As String
Private allowOrdering As Boolean
'
' ------------------------------------------------------------------------
' Catalog globals - must be globals so all routines can get to them
' ------------------------------------------------------------------------
'
Private SearchCategoryID As Long
Private itemId As Long
Private ItemAge As Long
Private ItemFeatured As Boolean
'
Private RecordTop As Long
Private RecordSpan As Long
Private SearchKeywords As String
Private SearchManufacturerID As Long
'
Private CatalogCategoryCount As Long
Private CatalogCategoryIndent As Long
Private CatalogRowcolor As String
'
' Set from site properties
'
'Private CatalogIndexFormat As Long
'
' ------------------------------------------------------------------------
' CatalogForm Data
' ------------------------------------------------------------------------
'
Private catalogForm As Long         ' determines which catalog form is displayed next
'
' ------------------------------------------------------------------------
' Common Recordset data so subs can share data
' ------------------------------------------------------------------------
'
Private RecordPtr As Long
Private RowEven As Boolean
Private RowIndent As Long
'
Const RecordSpanDefault = 20
'
Private Type RowDescriptorType
    CatalogAllowPricing As Boolean
    CatalogThumbnailWidth As Long
    CatalogImageWidth As Long
    CatalogAllowInventory As Boolean
    CatalogAllowOrdering As Boolean
    CatalogAliasRetailPrice As String
    CatalogAliasRegularPrice As String
    CatalogAllowSpecialPrice As Boolean
    CatalogAliasSpecialPrice As String
    CatalogAliasSalePrice As String
    CatalogAliasYourPrice As String
    CatalogListingColumns As Long
End Type
'
'===========================================================================================================
'
'===========================================================================================================
'
Public Function Execute(CsvObject As Object, mainObject As Object, optionString As String, FilterInput As String)
    '
    Call Init(mainObject)
    Execute = GetCatalog2(main.GetStreamInteger(rnSrcShopFormId), main.GetStreamInteger(rnDstShopFormId))
    '
    Exit Function
ErrorTrap:
    Call HandleError("CatalogClass2", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
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
    Dim Copy As String
    Dim Pricing As New CmrcPricingClass
    Dim common As New commonClass
    
    '
    ' --- common storage
    '
    ItemsContentName = "items"
    CategoryContentName = "Item Categories"
    CommerceAccountMemberID = main.memberId
    AllowGroupPolicies = main.IsAuthenticated
    Call Pricing.GetRightsPolicies(main, main.memberId, main.IsAuthenticated, IsMemberCommerceAdmin)
    If IsMemberCommerceAdmin Then
        CommerceAccountMemberID = main.GetVisitProperty("CommerceAccountMemberID", 0)
    End If
    '
    '   if catalog does not allow everyone to order - see if current user can order by way of a group
    '
    allowOrdering = common.getAllowOrdering(main, CommerceAccountMemberID)
    'main.WriteStream ("allowOrdering: " & allowOrdering)
    '
    ' ----- Get form
    '
    catalogForm = main.GetStreamInteger(rnSrcShopFormId)
    If catalogForm <> 0 Then
        Call main.AddRefreshQueryString(rnSrcShopFormId, catalogForm)
    End If
    '
    ' ----- Get and Track category ID
    '
    SearchCategoryID = main.GetStreamInteger(RequestNameCatalogCategoryID)
    If (SearchCategoryID > 0) Then
        Call main.AddRefreshQueryString(RequestNameCatalogCategoryID, SearchCategoryID)
        Call main.IncrementContentField(CategoryContentName, SearchCategoryID, "Viewings")
    End If
    '
    ' ----- Get and track item ID
    '
    itemId = main.GetStreamInteger(RequestNameCatalogItemID)
    If (itemId > 0) Then
        Call main.AddRefreshQueryString(RequestNameCatalogItemID, itemId)
        Call main.IncrementContentField(ItemsContentName, itemId, "Viewings")
    End If
    '
    ' ----- Get age attribute
    '
    ItemAge = main.GetStreamInteger(RequestNameCatalogItemAge)
    'If (ItemAge > 0) Then
    '    Call Main.AddRefreshQueryString(RequestNameCatalogItemAge, ItemAge)
    'End If
    '
    ' ----- read the featured value
    '
    ItemFeatured = main.GetStreamBoolean(RequestNameCatalogFeatured)
    'If ItemFeatured Then
    '    Call Main.AddRefreshQueryString(RequestNameCatalogFeatured, ItemFeatured)
    'End If
    '
    ' ----- Catalog Top
    '
    RecordTop = main.GetStreamInteger(RequestNameCatalogRecordTop)
    'If RecordTop > 0 Then
    '    Call Main.AddRefreshQueryString(RequestNameCatalogRecordTop, RecordTop)
    'End If
    '
    ' ----- Catalog Span
    '
    RecordSpan = main.GetStreamInteger(RequestNameCatalogSpan)
    If RecordSpan = 0 Then
        RecordSpan = main.GetSiteProperty("CatalogRecordSpan", RecordSpanDefault)
    End If
    If RecordSpan <> RecordSpanDefault Then
        Call main.AddRefreshQueryString(RequestNameCatalogSpan, RecordSpan)
    End If
    '
    ' ----- Catalog Query
    '
    SearchKeywords = main.GetStreamText(RequestNameCatalogKeywords)
    'If (SearchKeywords <> "") Then
    '    Call Main.AddRefreshQueryString(RequestNameCatalogKeywords, SearchKeywords)
    'End If
    '
    ' ----- Catalog Manufacturer Query
    '
    SearchManufacturerID = main.GetStreamInteger(RequestNameCatalogSource)
    '
    
    If main.RenderedPageID <> 0 Then
        Call main.SetSiteProperty("Catalog Location", "?bid=" & main.RenderedPageID)
    End If
    
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("Init"): Resume Next
    '
    End Sub
'
'
'
Public Function GetContent(optionString As String) As String
    GetContent = GetCatalog2(main.GetStreamInteger(rnSrcShopFormId), main.GetStreamInteger(rnDstShopFormId))
End Function
''=============================================================================
''   GetCatalog2_IndexFormRow( CSCategoryPointer )
''=============================================================================
''
'Private Function GetCatalog2_IndexFormRow(CSCategoryPointer As Long) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim SQL As String
'    Dim StatusOK As Boolean
'    Dim RowID As Long
'    Dim RowName As String
'    Dim CSPointer As Long
'    Dim CategoryContentID As Long
'    Dim CategoryContentName As String
'    Dim qs As String
'    Dim stream As String
'    '
'    If main.StreamOpen And main.IsCSOK(CSCategoryPointer) Then
'        '
'        ' ----- print out the category
'        '
'        RowID = (main.GetCSInteger(CSCategoryPointer, "id"))
'        RowName = kmaEncodeText(main.GetCSEncodedField(CSCategoryPointer, "name"))
'        CategoryContentID = (main.GetCSInteger(CSCategoryPointer, "contentcontrolid"))
'        CategoryContentName = kmaEncodeText(main.GetContentNameByID(CategoryContentID))
'        '
'        RowIndent = RowIndent + 1
'        stream = stream & "<tr><td>"
'        stream = stream & main.GetCSRecordEditLink(CSCategoryPointer)
'        If RowIndent > 1 Then
'            stream = stream & ("<img alt="""" src=""/ccLib/Images/spacer.gif"" height=""1"" width=""" & 15 * (RowIndent - 1) & """ >")
'        Else
'            stream = stream & ("<img alt="""" src=""/ccLib/Images/spacer.gif"" height=""1"" width=""15"" >")
'        End If
'        qs = main.RefreshQueryString
'        qs = ModifyQueryString(qs, rnSrcShopFormId, shopFormCatalogListing, True)
'        qs = ModifyQueryString(qs, RequestNameCatalogCategoryID, CStr(RowID), True)
'        stream = stream & ("&#187;&nbsp;<A href=""" & main.ServerPage & "?" & qs & """>" & RowName & "</A>")
'        stream = stream & ("</td></tr>")
'        '
'        ' ----- Now print all the child categories to this category
'        '
'        CSPointer = main.OpenCSContent_Internal(CategoryContentName, "(categoryid=" & RowID & ")")
'        Do While main.IsCSOK(CSPointer)
'            main.NextCSRecord (CSPointer)
'            stream = stream & GetCatalog2_IndexFormRow(CSPointer)
'        Loop
'        Call main.CloseCS(CSPointer)
'        RowIndent = RowIndent - 1
'    End If
'    '
'    GetCatalog2_IndexFormRow = stream
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("GetCatalog2_IndexFormRow"): Resume Next
'    '
'    End Function
''
''=============================================================================
'' Print Catalog General Index Form
''   Page with just the index
''=============================================================================
''
'Private Function GetCatalog2_IndexFormGeneral(ItemContentName As String, CategoryContentName As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim SQL As String
'    Dim Count As Long
'    Dim StatusOK As Boolean
'    Dim ItemCount As Long
'    Dim CSPointer As Long
'    Dim Id As Long
'    Dim QueryString As String
'    Dim BreadCrumbPrefix As String
'    Dim BreadCrumbDelimiter As String
'    Dim stream As String
'    Dim qs As String
'    '
'    If main.StreamOpen Then
'        qs = main.RefreshQueryString
'        '
'        CSPointer = main.OpenCSContent_Internal(CategoryContentName, "(( categoryid is null )OR( categoryid=0 ))")
'        If main.IsCSOK(CSPointer) Then
'            stream = stream & kmaStartTable(2, 0, 0)
'            RowEven = True
'            Do While main.IsCSOK(CSPointer)
'                stream = stream & GetCatalog2_IndexFormRow(CSPointer)
'                main.NextCSRecord (CSPointer)
'            Loop
'            stream = stream & kmaEndTable
'        Else
'            '
'            ' blank catalog place holder
'            '
'            stream = stream & "<P>&nbsp;</p>"
'        End If
'        Call main.CloseCS(CSPointer)
'        If main.IsLinkAuthoring(CategoryContentName) Then
'            stream = stream & kmaStartTable(2, 0, 0) & kmaStartTableRow & kmaStartTableCell()
'            stream = stream & main.GetCSRecordAddLink(CSPointer)
'            stream = stream & kmaEndTableCell & kmaEndTableRow & kmaEndTable
'        End If
'    End If
'    GetCatalog2_IndexFormGeneral = stream
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("GetCatalog2_IndexFormGeneral"): Resume Next
'    '
'End Function
''
''=============================================================================
'' Print Catalog General Index Form
''   Page with just the index
''=============================================================================
''
'Private Function GetCatalog2_IndexForm(ItemContentName As String, CategoryContentName As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim s As String
'    Dim copy As New CmrcCopyClass
'    '
'    If main.StreamOpen Then
'        '
'        s = "" _
'            & copy.GetCopy(main, shopCopyNameCatalogHome, shopCopyDefaultCatalogHome) _
'            & GetCatalog2_ListingForm(ItemContentName, CategoryContentName, 0)
'    End If
'    '
'    GetCatalog2_IndexForm = s
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("GetCatalog2_IndexForm"): Resume Next
'    '
'    End Function
''
''=============================================================================
'' Print Catalog Specials Index form
''   List of categories, plus specials and most popular
''=============================================================================
''
'Private Function GetCatalog2_IndexFormSpecials(ItemContentName As String, CategoryContentName As String) As String
'    On Error GoTo ErrorTrap
'    '
'    Dim SQL As String
'    'Dim RSCategories As Recordset
'    Dim Count As Long
'    Dim StatusOK As Boolean
'    Dim ItemCount As Long
'    Dim CSPointer As Long
'    Dim criteria As String
'    Dim s As String
'    Dim RowDescriptor As RowDescriptorType
'    '
'    If Not main.StreamOpen Then
'        Exit Function
'    End If
'    '
'    RowDescriptor = GetRowDescriptor()
'    's = s & "<div class=scSubMenu>Categoriesx</div>"
'    s = s & "<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"" ><tr>"
'    '
'    ' left side goes the index
'    '
'    s = s & "<td width=""50%"" valign=""top"" align=""left"">"
'    s = s & "Categories<BR>" & GetCatalog2_IndexFormGeneral(ItemContentName, CategoryContentName)
'    s = s & "</td>"
'    '
'    ' right side, featured items
'    '
'    s = s & ("<td width=""50%"" valign=""top"" align=""center"">")
'    ItemCount = 0
'    CSPointer = main.OpenCSContent_Internal(ItemContentName, "Featured<>" & SQLFalse)
'    If main.IsCSOK(CSPointer) Then
'        s = s & ("<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" class=""ccPanel"">")
'        s = s & ("<tr><td>" & scSpanNormalStyle & "Featured Items" & "</SPAN></td></tr>")
'        s = s & ("</table >")
'        Do While main.IsCSOK(CSPointer)
'            s = s & (GetCatalog2_ListingForm_GetRow_NextRecord(CSPointer, RowDescriptor))
'            ItemCount = ItemCount + 1
'        Loop
'    End If
'    Call main.CloseCS(CSPointer)
'    '
'    ' most popular
'    '
'    SQL = "SELECT Items.ID as id, Items.Name as name, Items.ItemNumber as itemnumber, Items.SourceItemNumber, Items.CategoryID, Items.ImageFilename, Items.ThumbnailFilename, Items.UnitCost, Items.UnitPrice, Items.OnSale, Items.SalePrice, Items.SpecialPrice, Items.Featured, Items.Weight, Items.QuantityOnhand, Items.QuantityLow, Organizations.Name AS SourceName, Items.BriefFilename AS BriefFilename, Items.CopyFilename AS CopyFilename, Items.Copy AS Copy, Items.OverView AS OverView, Items.UPC as UPC, Items.Options as Options, Items.altThumbSizeList as altThumbSizeList, Items.altSizeList as altSizeList, Items.HideOrderButton" _
'        & " FROM (OrderDetails LEFT JOIN Items ON OrderDetails.ItemID = Items.ID) LEFT JOIN Organizations ON Items.OrganizationID = Organizations.ID" _
'        & " GROUP BY Items.ID, Items.Name, Items.Active, Items.ItemNumber, Items.SourceItemNumber, Items.CategoryID, Items.ImageFilename, Items.ThumbnailFilename, Items.UnitCost, Items.UnitPrice, Items.OnSale, Items.SalePrice, Items.SpecialPrice, Items.Featured, Items.Weight, Items.QuantityOnhand, Items.QuantityLow, Organizations.Name,Items.BriefFilename,Items.CopyFilename, Items.Copy, Items.OverView, Items.UPC, Items.Options, Items.altThumbSizeList, Items.altSizeList, Items.HideOrderButton" _
'        & " HAVING (Items.Active<>" & KmaEncodeSQLBoolean(SQLFalse) & ")" _
'        & " ORDER BY Sum(OrderDetails.Quantity) DESC;"
'    CSPointer = main.OpenCSSQL("Default", SQL)
'    If main.IsCSOK(CSPointer) Then
'        s = s & ("<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" class=""ccPanel"">")
'        s = s & ("<tr><td>" & scSpanNormalStyle & "Most Popular Items" & "</SPAN></td></tr>")
'        s = s & ("</table >")
'        Do While main.IsCSOK(CSPointer) And ItemCount < 10
'            s = s & (GetCatalog2_ListingForm_GetRow_NextRecord(CSPointer, RowDescriptor))
'            ItemCount = ItemCount + 1
'        Loop
'    End If
'    Call main.CloseCS(CSPointer)
'    '
'    s = s & "</td></tr></table >"
'    GetCatalog2_IndexFormSpecials = s
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("GetCatalog2_IndexFormSpecials"): Resume Next
'    '
'    End Function
'
'===============================================================================
'   Print out multicolumn catalog row
'       Main.GetCSText( CSPointer, ) has the item to display
'
'   ---------------------------------
'   | photo | name (linked)         |
'   |       | description       |
'   ---------------------------------
'   | price and more info           |
'   ---------------------------------
'
'===============================================================================
'
Private Function GetCatalog2_ListingForm_GetRow_NextRecord(CSPointer As Long, RowDescriptor As RowDescriptorType) As String
    On Error GoTo ErrorTrap
    '
                        
    Dim TestFilename As String
    Dim CatalogAllowOrdering As Boolean
    Dim Filename As String
    Dim Copy As String
    Dim itemName As String
    Dim itemId As Long
    Dim ThumbnailFilename As String
    Dim imageFilename As String
    Dim BriefFilename As String
    Dim CopyFilename As String
    Dim ColumnWidth As Long
    Dim ColumnCount As Long
    Dim BriefLength As Long
    Dim DetailsLink As String
    Dim Divider As String
    Dim OrganizationName As String
    Dim QueryString As String
    Dim CatalogAllowMfg As Boolean
    Dim stream As New FastStringClass
    Dim PriceRegular As Double
    Dim PriceRetail As Double
    Dim PriceSpecial As Double
    Dim PriceSale As Double
    Dim PriceDiscount As Double
    Dim PriceYour As Double
    Dim GoodPriceSale As Boolean
    Dim GoodPriceRetail As Boolean
    Dim GoodPriceSpecial As Boolean
    Dim GoodPriceRegular As Boolean
    Dim Pricing As New CmrcPricingClass
    Dim MemberAllowSpecialPricing As Boolean
    Dim MemberDiscount As Double
    Dim NewLineTag As String
    Dim CellImage As String
    Dim CellCopy As String
    Dim CellPricing As String
    Dim DiscountFactor As Double
    Dim UPC As String
    Dim OptionList As String
    Dim categoryID As Long
    Dim CategoryName As String
    Dim OptionCell As String
    Dim CheckScript As String
    
    Dim baseFilename As String
    Dim fileExt As String
    Dim imgWidth As String
    Dim altSizeList As String
    Dim altThumbSizeList As String
    Dim sizArr() As String
    Dim sizPtr As Integer
    Dim objFile As New kmaFileSystem3.FileSystemClass
    Dim optionString As String
    
    Dim HideItemOrdering As Boolean
    
    If Not main.StreamOpen Then
        '
        ' skip the record and exit
        '
        main.NextCSRecord (CSPointer)
        RecordPtr = RecordPtr + 1
    Else
        '
        If RowDescriptor.CatalogListingColumns <= 0 Then
            '
            ' skip the record and exit
            '
            main.NextCSRecord (CSPointer)
            RecordPtr = RecordPtr + 1
        Else
            '
            ' ----- start a row
            '
            ColumnWidth = Int(100 / RowDescriptor.CatalogListingColumns)
            stream.Add ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr><td class=ccCatListRow>")
            stream.Add ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%""><tr>")
            For ColumnCount = 0 To RowDescriptor.CatalogListingColumns - 1
                stream.Add ("<td width=""" & ColumnWidth & "% align=""left"" valign=""top"" class=ccCatListCell>")
                If Not main.IsCSOK(CSPointer) Then
                    stream.Add ("&nbsp;")
                Else
                    CellImage = ""
                    CellCopy = ""
                    CellPricing = ""
                    itemId = (main.GetCSInteger(CSPointer, "ID"))
                    itemName = (main.GetCSText(CSPointer, "Name"))
                    UPC = main.GetCS(CSPointer, "UPC")
                    OptionList = main.GetCS(CSPointer, "Options")
                    ThumbnailFilename = (main.GetCSText(CSPointer, "ThumbnailFilename"))
                    altThumbSizeList = main.GetCS(CSPointer, "altThumbSizeList")
                    imageFilename = (main.GetCSText(CSPointer, "ImageFilename"))
                    altSizeList = main.GetCS(CSPointer, "altSizeList")
                    categoryID = main.GetCSInteger(CSPointer, "CategoryID")
                    CategoryName = main.GetCS(CSPointer, "categoryid")
                    'CategoryName = Main.GetRecordName("Item Categories", CategoryID)
                    QueryString = main.RefreshQueryString
                    QueryString = ModifyQueryString(QueryString, rnSrcShopFormId, shopFormCatalogDetails, True)
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogItemID, CStr(itemId), True)
                    DetailsLink = main.ServerPage & "?" & QueryString
                    '
                    ' Image Column
                    '
Dim altSizeFound As Boolean
Dim listFix As String
                    If ThumbnailFilename <> "" Then
                        '
                        ' item contains a thumbnail image, use the correct size
                        '
                        baseFilename = ThumbnailFilename
                        Filename = baseFilename
                        fileExt = Right(baseFilename, Len(baseFilename) - (InStrRev(baseFilename, ".", , vbTextCompare) - 1))
                        imgWidth = RowDescriptor.CatalogThumbnailWidth
                        altSizeFound = False
                        If (altThumbSizeList <> "") And (InStr(1, altThumbSizeList, imgWidth & "x") <> 0) Then
                            sizArr = Split(altThumbSizeList, ",")
                            For sizPtr = 0 To UBound(sizArr)
                                If (InStr(1, listFix, sizArr(sizPtr)) <= 0) Then
                                    '
                                    ' not repeat, add to fix list and check it
                                    '
                                    listFix = listFix & "," & sizArr(sizPtr)
                                    If InStr(1, sizArr(sizPtr), imgWidth & "x") <> 0 Then
                                        ThumbnailFilename = Replace(baseFilename, fileExt, "-" & sizArr(sizPtr) & fileExt)
                                        altSizeFound = True
                                        'Exit For
                                    End If
                                End If
                            Next
                            If listFix <> "" Then
                                listFix = Mid(listFix, 2)
                            End If
                            If listFix <> altThumbSizeList Then
                                Call main.SetCS(CSPointer, "altThumbSizeList", listFix)
                            End If
'                            'If InStr(1, altThumbSizeList, ",") <> 0 Then
'                                sizArr = Split(altThumbSizeList, ",")
'                                For sizPtr = 0 To UBound(sizArr)
'                                    If InStr(1, sizArr(sizPtr), imgWidth & "x") <> 0 Then
'                                        ThumbnailFilename = Replace(baseFilename, fileExt, "-" & sizArr(sizPtr) & fileExt)
'                                        altSizeFound = True
'                                        Exit For
'                                    End If
'                                Next
'                            'Else
'                            '    ThumbnailFilename = Replace(baseFilename, fileExt, "-" & altThumbSizeList & fileExt)
'                            'End If
                        End If
                        If Not altSizeFound Then
                            ThumbnailFilename = getResizedImage(baseFilename, CStr(RowDescriptor.CatalogThumbnailWidth), CSPointer, "altThumbSizeList")
                            If imageFilename = "" Then
                                imageFilename = baseFilename
                            End If
                        End If
                        TestFilename = main.PhysicalFilePath & ThumbnailFilename
                        TestFilename = Replace(main.PhysicalFilePath & ThumbnailFilename, "/", "\")
                        If Not objFile.CheckFile(TestFilename) Then
                            ThumbnailFilename = baseFilename
                        End If
                        CellImage = "<A href=""" & DetailsLink & """><img class=""scThumbNail"" src=""" & main.serverFilePath & ThumbnailFilename & """ border=""0"" align=""left""  style=""margin: 10px;"" width=""" & RowDescriptor.CatalogThumbnailWidth & """ ></A>"
                    ElseIf imageFilename <> "" Then
                        '
                        ' Item has an image but no thumbnail, use image as thumbnail
                        '
                        baseFilename = imageFilename
                        Filename = baseFilename
                        fileExt = Right(baseFilename, Len(baseFilename) - (InStrRev(baseFilename, ".", , vbTextCompare) - 1))
                        imgWidth = RowDescriptor.CatalogThumbnailWidth
                        altSizeFound = False
                        If (altSizeList <> "") And (InStr(1, altSizeList, imgWidth & "x") <> 0) Then
                            sizArr = Split(altSizeList, ",")
                            For sizPtr = 0 To UBound(sizArr)
                                If (InStr(1, listFix, sizArr(sizPtr)) <= 0) Then
                                    '
                                    ' not repeat, add to fix list and check it
                                    '
                                    listFix = listFix & "," & sizArr(sizPtr)
                                    If (Not altSizeFound) And (InStr(1, sizArr(sizPtr), imgWidth & "x") <> 0) Then
                                        imageFilename = Replace(baseFilename, fileExt, "-" & sizArr(sizPtr) & fileExt)
                                        altSizeFound = True
                                        'WidthKnown = True
                                        'Exit For
                                    End If
                                End If
                            Next
                            If listFix <> "" Then
                                listFix = Mid(listFix, 2)
                            End If
                            If listFix <> altSizeList Then
                                Call main.SetCS(CSPointer, "altSizeList", listFix)
                            End If
'                            'If InStr(1, altSizeList, ",") <> 0 Then
'                                sizArr = Split(altSizeList, ",")
'                                For sizPtr = 0 To UBound(sizArr)
'                                    If InStr(1, sizArr(sizPtr), imgWidth & "x") <> 0 Then
'                                        imageFilename = Replace(baseFilename, fileExt, "-" & sizArr(sizPtr) & fileExt)
'                                        altSizeFound = True
'                                        Exit For
'                                    End If
'                                Next
'                            'Else
'                            '    ImageFilename = Replace(baseFilename, fileExt, "-" & altSizeList & fileExt)
'                            'End If
                        End If
                        If Not altSizeFound Then
                            imageFilename = getResizedImage(baseFilename, CStr(RowDescriptor.CatalogThumbnailWidth), CSPointer, "altSizeList")
                            If imageFilename = "" Then
                                imageFilename = baseFilename
                            End If
                        End If
                        TestFilename = main.PhysicalFilePath & imageFilename
                        TestFilename = Replace(main.PhysicalFilePath & imageFilename, "/", "\")
                        If Not objFile.CheckFile(TestFilename) Then
                            imageFilename = baseFilename
                        End If
                        CellImage = "<A href=""" & DetailsLink & """><img class=""scThumbNail"" src=""" & main.serverFilePath & imageFilename & """ border=""0"" align=""left""  style=""margin: 10px;"" width=""" & RowDescriptor.CatalogThumbnailWidth & """ ></A>"

'                        baseFilename = ImageFilename
'                        Filename = baseFilename
'                        fileExt = Right(baseFilename, Len(baseFilename) - (InStrRev(baseFilename, ".", , vbTextCompare) - 1))
'                        imgWidth = RowDescriptor.CatalogThumbnailWidth
'                        If (altSizeList <> "") And (InStr(1, altSizeList, imgWidth & "x") <> 0) Then
'                            If InStr(1, altSizeList, ",") <> 0 Then
'                                sizArr = Split(altSizeList, ",")
'                                For sizPtr = 0 To UBound(sizArr)
'                                    If InStr(1, sizArr(sizPtr), imgWidth & "x") <> 0 Then
'                                        ImageFilename = Replace(baseFilename, fileExt, "-" & sizArr(sizPtr) & fileExt)
'                                    End If
'                                Next
'                            Else
'                                ImageFilename = Replace(baseFilename, fileExt, "-" & altSizeList & fileExt)
'                            End If
'                        Else
'                            ImageFilename = getResizedImage(baseFilename, CStr(RowDescriptor.CatalogThumbnailWidth), CSPointer, "altThumbSizeList")
'                        End If
'
'                        If Not objFile.CheckFile(Main.PhysicalFilePath & ImageFilename) Then
'                            ImageFilename = baseFilename
'                        End If
'
'                        '
'                        ' ----- Image present, display it narrow
'                        '
'                        CellImage = "<A href=""" & DetailsLink & """><img class=""scThumbNail"" src=""" & Main.ServerFilePath & ImageFilename & """ border=""0"" align=""left"" style=""margin: 10px;"" width=""" & RowDescriptor.CatalogThumbnailWidth & """ ></A>"
                    Else
                        '
                        ' ----- No Image, use stock gif
                        '
                        CellImage = "<A href=""" & DetailsLink & """><img class=""scThumbNail"" src=""/cclib/images/noimagethumbnail.gif"" border=""0"" align=""left"" style=""margin: 10px;"" width=""" & RowDescriptor.CatalogThumbnailWidth & """ ></A>"
                    End If
                    '
                    ' copy column
                    '
                    CellCopy = CellCopy & main.GetRecordEditLink("items", itemId)
                    CellCopy = CellCopy & "<A href=""" & DetailsLink & """><B>" & itemName & "</B></A>"
                    '
                    CatalogAllowMfg = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowMfg", False))
                    If CatalogAllowMfg Then
                        OrganizationName = main.GetCSEncodedField(CSPointer, "OrganizationID")
                        If OrganizationName <> "" Then
                            CellCopy = CellCopy & "<br>made by " & OrganizationName
                            End If
                        End If
                    Copy = ""
                    Copy = main.GetCS(CSPointer, "OverView") & main.GetCS(CSPointer, "BriefFilename")
                    If (Copy = "<P>&nbsp;</p>") Or (Copy = "<P>&nbsp;</p><P>&nbsp;</p>") Then
                        Copy = ""
                    End If
                    If Copy <> "" Then
                        If InStr(1, Copy, "<P", vbTextCompare) = 0 Then
                            Copy = "<P>" & Copy & "</p>"
                        End If
                    End If
                    If Copy = "" Then
                        NewLineTag = "<BR><BR>"
                    Else
                        NewLineTag = ""
                        CellCopy = CellCopy & Copy
                    End If
                    '
                    Copy = (main.GetCSText(CSPointer, "ItemNumber"))
                    If Copy <> "" Then
                        CellCopy = CellCopy & "<BR>Item # " & Copy
                    End If
                    '
'Dim options() As String
'Dim optionCnt As Long
'Dim ptr As Long
'
'                    copy = Trim(main.GetCSText(CSPointer, "Options"))
'                    If copy <> "" Then
'                        options = Split(copy, vbCrLf)
'                        copy = ""
'                        optionCnt = 0
'                        For ptr = 0 To UBound(options)
'                            If options(ptr) <> "" Then
'                                copy = copy & "<li>" & options(ptr) & "</li>"
'                                optionCnt = optionCnt + 1
'                            End If
'                        Next
'                        If optionCnt > 0 Then
'                            If optionCnt = 1 Then
'                                copy = "<div>Option</div>""<ul class=""scOptionList"">" & copy & "</ul>"
'                            Else
'                                copy = "<div>Options</div>""<ul class=""scOptionList"">" & copy & "</ul>"
'                            End If
'                            copy = copy & "<ul>" & copy & "</ul>"
'                        End If
'
'                        CellCopy = CellCopy & "<BR>Option " & copy
'                    End If
'                    '
                    Copy = (main.GetCSText(CSPointer, "UPC"))
                    If Copy <> "" Then
                        CellCopy = CellCopy & "<BR>UPC " & Copy
                    End If
                    '
                    If RowDescriptor.CatalogAllowInventory Then
                        CellCopy = CellCopy & "<BR>Availability: "
                        If main.GetCSInteger(CSPointer, "QuantityOnHand") > 0 Then
                            CellCopy = CellCopy & "In Stock"
                        Else
                            CellCopy = CellCopy & "Back Order"
                        End If
                    End If
                    '
                    'CellCopy = CellCopy & "<div>&nbsp;</div>"
                    If CategoryName <> "" Then
                        QueryString = main.RefreshQueryString
                        QueryString = ModifyQueryString(QueryString, rnSrcShopFormId, shopFormCatalogListing, True)
                        QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, CStr(categoryID), True)
                        CellCopy = CellCopy & "<div class=""scCategories"">Category: " & getCategory(categoryID) & "</div>"
                    End If
                    If itemId <> 0 Then
                        QueryString = main.RefreshQueryString
                        QueryString = ModifyQueryString(QueryString, rnSrcShopFormId, shopFormCatalogDetails, True)
                        QueryString = ModifyQueryString(QueryString, RequestNameCatalogItemID, CStr(itemId), True)
                        CellCopy = CellCopy & "<div class=""scMoreInfo""><A href=""" & main.ServerPage & "?" & QueryString & """>More Information</A></div>"
                    End If
                    '
                    ' ----- pricing column
                    '
                    If RowDescriptor.CatalogAllowPricing Then
                        CellPricing = CellPricing & Pricing.GetPricingCopy(main, CSPointer, RowDescriptor.CatalogAliasYourPrice, RowDescriptor.CatalogAliasRegularPrice, RowDescriptor.CatalogAliasSalePrice, RowDescriptor.CatalogAliasRetailPrice, RowDescriptor.CatalogAliasSpecialPrice, RowDescriptor.CatalogAllowSpecialPrice, CommerceAccountMemberID, AllowGroupPolicies)
                    End If
                    '
                    ' ----- Get Options
                    '
                    Call GetOptionCell(main.GetCSText(CSPointer, "Options"), OptionCell, CheckScript, itemId)
                    '
                    '   now also checking if the group a user in can order
                    '
                    'CatalogAllowOrdering = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowOrdering", True))
                    CatalogAllowOrdering = allowOrdering
                    HideItemOrdering = main.GetCSBoolean(CSPointer, "HideOrderButton")
                    '
                    If (CatalogAllowOrdering) And (Not HideItemOrdering) Then
                        QueryString = main.RefreshQueryString
                        CellPricing = CellPricing _
                            & main.GetFormStart(QueryString) _
                            & main.GetFormInputHidden(RequestNameCatalogOrder, CStr(itemId))
                        If OptionCell <> "" Then
                            CellPricing = CellPricing _
                                & "<div>&nbsp;</div>" _
                                & OptionCell _
                                & ""
                        End If
                        
                        CellPricing = CellPricing _
                            & "<div>&nbsp;</div>" _
                            & main.GetFormButton(ButtonOrder, , , CheckScript) _
                            & main.GetFormEnd _
                            & ""
'                        CellPricing = CellPricing _
'                            & "<div>&nbsp;</div>" _
'                            & Main.GetFormButton(ButtonOrder, , , "CheckOptions_" & itemID & "()") _
'                            & Main.GetFormEnd _
'                            & ""
                        Divider = "<br>"
                    End If
                    '
                    If RowDescriptor.CatalogListingColumns = 1 Then
                        stream.Add ("<table border=0 cellpadding=4 cellspacing=0><tr>")
                        stream.Add ("<td valign=top class=ccCatListImage>" & CellImage & "</td>")
                        stream.Add ("<td valign=top class=ccCatListCopy width=50% >" & CellCopy & "</td>")
                        stream.Add ("<td valign=top class=ccCatListCopy width=50% >" & CellPricing & "</td>")
                        stream.Add ("</tr></table >")
                    Else
                        stream.Add ("<table border=0 cellpadding=4 cellspacing=0>")
                        stream.Add ("<tr><td valign=top class=ccCatListImage>" & CellImage & "</td></tr>")
                        stream.Add ("<tr><td valign=top class=ccCatListCopy>" & CellCopy & "</td></tr>")
                        stream.Add ("<tr><td valign=top class=ccCatListCopy>" & CellPricing & "</td></tr>")
                        stream.Add ("</table >")
                    End If
                    '
                    ' ----- done, go to next item
                    '
                    main.NextCSRecord (CSPointer)
                    RecordPtr = RecordPtr + 1
                End If
                '
                ' ----- next column
                '
                stream.Add ("</td>")
            Next
            stream.Add ("</tr></table ></td></tr></table >")
            '
            GetCatalog2_ListingForm_GetRow_NextRecord = stream.Text
        End If
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCatalog2_ListingForm_GetRow_NextRecord"): Resume Next
    '
    End Function
''
''=============================================================================
''   GetCatalog2_ListingForm( CatalogListingColumns )
''       Prints a catalog page with multiple listings in a 100% table.
''=============================================================================
''
'Private Function GetCatalog2_ListingForm(ItemContentName As String, CategoryContentName As String, SearchCategoryID As Long) As String
'    On Error GoTo ErrorTrap
'    '
'    GetCatalog2_ListingForm = GetCatalog2_ListingForm(ItemContentName, CategoryContentName, SearchCategoryID)
'    Exit Function
'    '
'    ' ----- Error Trap
'    '
'ErrorTrap:
'    Call HandleClassErrorAndResumeNext("GetCatalog2_ListingForm"): Resume Next
'    '
'    End Function
'
'=============================================================================
'   GetCatalog2_ListingForm( CatalogListingColumns )
'       Prints a catalog page with multiple listings in a 100% table.
'=============================================================================
'
Private Function GetCatalog2_ListingForm(ItemContentName As String, CategoryContentName As String, SearchCategoryID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim paginationMsg As String
    Dim RowDescriptor As RowDescriptorType
    Dim cs As Long
    Dim SQL As String
    Dim categoryID As Long
    Dim CategoryIDHeader As Long
    Dim CategoryHeadroom As Long
    Dim RecordNext As Long
    Dim RecordTopPrevious As Long
    Dim RecordTopNext As Long
    Dim StatusOK As Boolean
    Dim RowColor As String
    Dim CategoryNew As Boolean
    Dim FirstSpace As String
    Dim FirstComma As String
    Dim FirstDelimiter As String
    Dim keywords As String
    Dim keyword As String
    Dim SQLLikeWildcard As String
    Dim DataSourceName As String
    Dim criteria As String
    Dim CriteriaDelimiter As String
    Dim CSPointer As Long
    Dim QueryString As String
    Dim CatalogAllowMfg As Boolean
    Dim s As String
    Dim SearchCaption As String
    Dim Copy As String
    Dim CellPtr As Long
    Dim CellWidth As String
    Dim SubID As Long
    Dim SubName As String
    Dim SubCopy As String
    Dim WorkingQS As String
    Dim qs As String
    Dim IsAuthoring As Boolean
    Dim SearchCategoryName As String
    
    Dim Sort As String
    Dim csSort As Long
    Dim recordCnt As Long
    
    Const CellsPerRow = 3
    SQLLikeWildcard = "%"
    '
    If main.StreamOpen Then
        '
        Call main.AddPageTitle("Catalog")
        If catalogForm = shopFormCatalogDetails Then
            '
            ' Detail Page
            '
            s = s & GetCatalog2_DetailsForm(itemId, ItemContentName, CategoryContentName)
        Else
            RowDescriptor = GetRowDescriptor()
            '
            ' Listing Page - Category Landing, then SubCategories, then Items
            '
            IsAuthoring = main.IsAuthoring("Items")
            '
            ' Search Criteria and Page Heading
            '
            DataSourceName = main.GetContentDataSource(ItemContentName)
            '
            criteria = ""
            CriteriaDelimiter = ""
            If main.IsContentFieldSupported("items", "isInCatalog") Then
                criteria = criteria & CriteriaDelimiter & "(isInCatalog=1)"
                CriteriaDelimiter = "AND"
            End If
            SearchCaption = "All Items"
            recordCnt = 0
            If (itemId > 0) Then
                criteria = criteria & CriteriaDelimiter & "(ID=" & KmaEncodeSQLNumber(itemId) & ")"
                CriteriaDelimiter = "AND"
            Else
                If ItemFeatured Then
                    SearchCaption = " that are featured"
                    criteria = criteria & CriteriaDelimiter & "(Featured<>0)"
                    CriteriaDelimiter = "AND"
                End If
                If (SearchManufacturerID > 0) Then
Dim SearchManufacturerName As String
                    SearchManufacturerName = main.GetRecordName("Organizations", SearchManufacturerID)
                    If SearchManufacturerName <> "" Then
                        SearchCaption = SearchCaption & " from " & SearchManufacturerName
                        criteria = criteria & CriteriaDelimiter & "(OrganizationID=" & KmaEncodeSQLNumber(SearchManufacturerID) & ")"
                        CriteriaDelimiter = "AND"
                    End If
                End If
                If (ItemAge > 0) Then
                    SearchCaption = SearchCaption & " added in the past " & ItemAge & " days"
                    criteria = criteria & CriteriaDelimiter & "(DateAdded>=" & KmaEncodeSQLDate(Now() - ItemAge) & ")"
                    CriteriaDelimiter = "AND"
                End If
                If (SearchKeywords <> "") Then
                    keywords = SearchKeywords
                    SearchCaption = SearchCaption & " matching keyword(s) " & kmaEncodeHTML(keywords)
                    Do While Len(keywords) <> 0
                        FirstSpace = InStr(keywords, " ")
                        FirstComma = InStr(keywords, ",")
                        FirstDelimiter = FirstSpace
                        If FirstComma <> 0 And FirstComma < FirstDelimiter Then
                            FirstDelimiter = FirstComma
                        End If
                        If FirstDelimiter = 0 Then
                            keyword = keywords
                        keywords = ""
                        Else
                            keyword = Mid(keywords, 1, FirstDelimiter - 1)
                            keywords = Mid(keywords, FirstDelimiter + 1)
                        End If
                        criteria = criteria & CriteriaDelimiter & "("
                        criteria = criteria & "(Name LIKE " & KmaEncodeSQLText(SQLLikeWildcard & keyword & SQLLikeWildcard) & ")"
                        CatalogAllowMfg = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowMfg", False))
                        If CatalogAllowMfg Then
                            criteria = criteria & "or(Organizations.Name LIKE " & KmaEncodeSQLText(SQLLikeWildcard & keyword & SQLLikeWildcard) & ")"
                        End If
                        criteria = criteria & ")"
                        CriteriaDelimiter = "AND"
                    Loop
                End If
                If (SearchCategoryID > 0) Then
                    '
                    ' a specific category
                    '
                    SearchCategoryName = main.GetRecordName("Item Categories", SearchCategoryID)
                    If SearchCategoryName <> "" Then
                        SearchCaption = SearchCaption & " in " & SearchCategoryName
                        criteria = criteria & CriteriaDelimiter & "((CategoryID=" & SearchCategoryID & ")" & GetItemCategoryCriteria(SearchCategoryID) & ")"
                        CriteriaDelimiter = "AND"
                    End If
                ElseIf (SearchCategoryID = -1) Then
                    '
                    ' all categories
                    '
                Else
                    '
                    ' not all categories or a specific category
                    '
                    If (criteria = "") And (catalogHasFeaturedItems()) Then
                        '
                        ' no search specified, if there are featured items show them
                        '
                        ItemFeatured = True
                        SearchCaption = "Featured Items"
                        criteria = criteria & CriteriaDelimiter & "(Featured<>0)"
                    End If
                End If
                Call main.AddPageTitle(SearchCaption)
                s = s & "<div>&nbsp;</div><div>" & SearchCaption & "</div>"
                s = s & "<div>&nbsp;</div>"
            End If
            '
            Sort = main.GetSiteProperty("CatalogItemOrder")
            If Sort <> "" Then
                csSort = main.OpenCSContent("Sort Methods", "ID=" & main.EncodeSQLNumber(Sort), , , , , "OrderByClause")
                If main.CSOK(csSort) Then
                    Sort = main.GetCSText(csSort, "OrderByClause")
                End If
                Call main.CloseCS(csSort)
            Else
                Sort = "name"
            End If
            categoryID = SearchCategoryID
            RecordPtr = RecordTop
            Call main.TestPoint("criteria=" & criteria)
            CSPointer = main.OpenCSContent_Internal(ItemContentName, criteria, Sort, , , , , RecordSpan, 1 + (RecordTop / RecordSpan))
            If Not main.IsCSOK(CSPointer) Then
                s = s & ("<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">")
                s = s & ("<tr><td>")
                s = s & "&nbsp;<BR>There are no items to display<br>"
                If (SearchCategoryID > 0) Then
                    s = s & main.GetCSRecordAddLink(CSPointer, "CategoryID=" & KmaEncodeSQLNumber(SearchCategoryID))
                Else
                    s = s & main.GetCSRecordAddLink(CSPointer)
                End If
                s = s & ("</td></tr></table >")
            Else
                '
                ' Repeat for all items found
                '
                recordCnt = main.GetCSRowCount(CSPointer)
                CategoryIDHeader = 0
                CategoryHeadroom = 1
                RecordTopNext = Int(RecordTop) + Int(RecordSpan)
                '
                s = s & "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                s = s & "<tr><td class=""ccCatHead""><!-- catHead --></td></tr>"
                s = s & "<tr><td class=""ccCatList"">"
                s = s & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
                Do While (main.IsCSOK(CSPointer) And (RecordPtr < RecordTopNext))
                    categoryID = (main.GetCSInteger(CSPointer, "CategoryID"))
                    s = s & "<tr><td width=""100%"" valign=""top"">" & GetCatalog2_ListingForm_GetRow_NextRecord(CSPointer, RowDescriptor) & "</td></tr>"
                    ' next is included in the getrow call
                    'RecordPtr = RecordPtr + 1
                    'Main.NextCSRecord (CSPointer)
                Loop
                If IsAuthoring Then
                    s = s & "<tr><td>" & main.GetRecordAddLink("items", "CategoryID=" & categoryID) & "Add an Item</td></tr>"
                End If
                s = s & "</table >"
                s = s & "</td>"
                RecordNext = RecordTop
                If main.IsCSOK(CSPointer) Then
                    RecordNext = RecordPtr
                End If
                RecordTopPrevious = RecordTop - RecordSpan
                If RecordTopPrevious < 0 Then
                    RecordTopPrevious = 0
                End If
                's = s & (scSpanNormalStyle)
                '
                ' set the pagination message
                '
                If ItemFeatured Then
                    paginationMsg = "Featured Results"
                Else
                    paginationMsg = "Results"
                End If
                If RecordPtr = 0 Then
                    '
                    '
                    '
                    paginationMsg = "No " & paginationMsg & " Found"
                    s = s & ("[no items were found]")
                ElseIf (RecordTop = 0) And (RecordPtr = recordCnt) Then
                    '
                    '
                    '
                    paginationMsg = "Showing All " & paginationMsg & ""
                Else
                    's = s & "<div>&nbsp;</div>"
                    paginationMsg = "Showing " & RecordTop + 1 & " - " & RecordPtr & " of " & recordCnt & " " & paginationMsg & ""
                    's = s & "<div class=""scPagination"">[Items " & RecordTop + 1 & " to " & RecordPtr & "] of " & main.GetCSRowCount(CSPointer) & "</div>"
                End If
                '
                ' First,Next,Previous
                '
                QueryString = main.RefreshQueryString
                QueryString = ModifyQueryString(QueryString, rnSrcShopFormId, shopFormCatalogListing, True)
                If itemId <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogItemID, CStr(itemId), True)
                End If
                If SearchCategoryID <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, CStr(SearchCategoryID), True)
                End If
                If ItemAge <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogItemAge, CStr(ItemAge), True)
                End If
                If ItemFeatured Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogFeatured, CStr(ItemFeatured), True)
                End If
                If SearchManufacturerID <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogSource, CStr(SearchManufacturerID), True)
                End If
                If RecordTop <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogRecordTop, CStr(0), True)
                    paginationMsg = paginationMsg & ("&nbsp;|&nbsp;<A href=""" & main.ServerPage & "?" & QueryString & """>First</A>")
                    '
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogRecordTop, CStr(RecordTopPrevious), True)
                    paginationMsg = paginationMsg & ("&nbsp;|&nbsp;<A href=""" & main.ServerPage & "?" & QueryString & """>Previous</A>")
                End If
                If RecordNext = RecordPtr Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogRecordTop, CStr(RecordNext), True)
                    paginationMsg = paginationMsg & ("&nbsp;|&nbsp;<A href=""" & main.ServerPage & "?" & QueryString & """>Next</A>")
                End If
                '
                '
                '
                s = s & "<tr><td class=""ccCatHead"">" & paginationMsg & "</td></tr>"
                's = s & "<div class=""scPagination"">" & paginationMsg & "</div>"
                s = Replace(s, "<!-- catHead -->", paginationMsg)
                '
                '
                '
                s = s & "</tr></table >"
                's = s & ("</SPAN>")
            End If
            Call main.CloseCS(CSPointer)
        End If
        Select Case catalogForm
            Case shopFormCatalogDetails
            Case Else
        End Select
    End If
    '
    GetCatalog2_ListingForm = s
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCatalog2_ListingForm"): Resume Next
    '
    End Function
'
'===============================================================================
'   GetCatalog2_DetailsForm
'===============================================================================
'
Private Function GetCatalog2_DetailsForm(itemId As Long, ItemContentName As String, CategoryContentName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CategoryName As String
    Dim cs As Long
    Dim itemDescription As String
    Dim ItemBrief As String
    Dim SQL As String
    Dim Copy As String
    Dim StatusOK As Boolean
    Dim CopyFilename As String
    Dim QuantityOnHand As Long
    Dim QuantityLow As Long
    Dim ItemNumber As String
    Dim BriefFilename As String
    Dim baseFilename As String
    Dim Filename As String
    Dim ItemUnitPrice As Double
    Dim ItemRetailPrice As Double
    Dim ItemSpecialPrice As Double
    Dim ItemSalePrice As Double
    Dim CSPointer As Long
    Dim categoryID As Long
    Dim QueryString As String
    Dim CatalogAllowInventory As Boolean
    Dim CatalogAllowOrdering As Boolean
    Dim stream As New FastStringClass
    Dim Cell As String
    Dim PricingCell As String
    Dim Pricing As New CmrcPricingClass
    Dim ThumbnailWidth As Long
    Dim qs As String
    Dim RowPtr As Long
    Dim OptionArgs() As String
    Dim OptionRow As String
    Dim Pos As Long
    Dim OptionCell As String
    Dim CheckScript As String
    Dim RowDescriptor As RowDescriptorType
    
    Dim fileExt As String
    Dim altSizeList As String
    Dim imgWidth As String
    Dim sizArr() As String
    Dim sizPtr As Integer
    Dim ItemContentID As Long
    
    Dim itemName As String
    Dim objFile As New kmaFileSystem3.FileSystemClass
    Dim HideItemOrdering As Boolean
    Dim QSBase As String
    '
    If main.StreamOpen Then
        '
        ' ----- load the Content
        '
        '
        QSBase = main.RefreshQueryString
        QSBase = ModifyQueryString(QSBase, rnSrcShopFormId, "", False)
        QSBase = ModifyQueryString(QSBase, rnDstShopFormId, "", False)
        QSBase = ModifyQueryString(QSBase, RequestNameCatalogItemID, "", False)
        QSBase = ModifyQueryString(QSBase, "method", "", False)
        '
        ' link to shopping
        '
        qs = QSBase
        stream.Add ("<div class=""scDetailReturn""><a href=""?" & qs & """>Return to List</a></div>")
        RowDescriptor = GetRowDescriptor
        CSPointer = main.OpenCSContent_Internal(ItemContentName, "ID=" & itemId)
        If Not main.IsCSOK(CSPointer) Then
            stream.Add ("<p>No items were found. Please return to the <A href=""" & GetShoppingCartLink(main.ServerPage, main.RefreshQueryString) & """>Catalog</A> and search again. " & main.AdminMessage & "</p>")
        Else
            ItemContentID = main.GetCSInteger(CSPointer, "contentcontrolid")
            categoryID = (main.GetCSInteger(CSPointer, "CategoryID"))
            itemId = (main.GetCSInteger(CSPointer, "ID"))
            itemName = main.GetCS(CSPointer, "Name")
            itemDescription = main.GetCS(CSPointer, "Copy") & main.GetCS(CSPointer, "CopyFilename")
            ItemBrief = main.GetCS(CSPointer, "OverView") & main.GetCS(CSPointer, "BriefFilename")
            '
            ' Add meta content, if meta content missing, add meta content defaults
            '
            Call main.SetMetaContent(ItemContentID, itemId)
            cs = main.OpenCSContent("Meta Content", "(ContentID=" & ItemContentID & ")and(RecordID=" & itemId & ")", , , , , "Name")
            If main.IsCSOK(cs) Then
                Copy = main.GetCSText(cs, "Name")
                If Copy = "" Then
                    If itemName <> "" Then
                        Call main.AddPageTitle(kmaEncodeHTML(itemName))
                    End If
                End If
            End If
            Call main.CloseCS(cs)
            'If ItemName = "" Then
            '    ItemName = "Item " & ItemNumber
            'End If
            'Call Main.AddPageTitle(ItemName)
            Call main.TrackContentSet(CSPointer)
            '
            ' ----- Print catalog page
            '
            stream.Add (main.GetSiteProperty("TextSearchStartTag", TextSearchStartTagDefault))
            stream.Add (main.GetCSRecordEditLink(CSPointer))
            '
            ' ----- Category
            '
            If categoryID <> 0 Then
                CategoryName = getCategory(categoryID)
                If CategoryName <> "" Then
                    Cell = Cell & "<div class=""scCategories"">"
                    Cell = Cell & "Category: "
                    Cell = Cell & getCategory(categoryID)
                    Cell = Cell & "</div>"
                End If
            End If
            '
            ' ----- Item Name
            '
            Cell = Cell & "<h2>" & itemName & "</h2>"
            '
            ' ----- Item name
            '
            If Copy <> "" Then
                Cell = Cell & "<div><B>" & Copy & "</B></div>"
            End If
            '
            ' ----- Left Cell, Thumbnail
            '
            Filename = (main.GetCSText(CSPointer, "Thumbnailfilename"))
            If Filename <> "" Then
                ThumbnailWidth = kmaEncodeInteger(main.GetSiteProperty("CatalogThumbnailWidth", 100))
                Cell = Cell & "<a href=#Image><img border=0 align=left src=""http://" & main.ServerHost & main.serverFilePath & Filename & """ border=""0"" style=""margin: 10px;"" width=" & ThumbnailWidth & "></a>"
            End If
            '
            ' ----- Item CopyFilename
            '
            Copy = itemDescription
            If Copy = "" Then
                Copy = ItemBrief
                If Copy = "" Then
                    Copy = "" _
                        & "" _
                        & "<div>No description is available</div>"
                End If
            End If
            Cell = Cell & Copy
            '
            ' ----- Pricing info
            '
            ItemNumber = (main.GetCSText(CSPointer, "ItemNumber"))
            If ItemNumber <> "" Then
                PricingCell = PricingCell & "<div>Item # " & ItemNumber & "</div>"
            End If
            '
            Copy = (main.GetCSText(CSPointer, "UPC"))
            If Copy <> "" Then
                PricingCell = PricingCell & "<div>UPC " & Copy & "</div>"
            End If
            '
            CatalogAllowInventory = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowInventory", "false"))
            If CatalogAllowInventory Then
                QuantityOnHand = (main.GetCSInteger(CSPointer, "QuantityOnHand"))
                If QuantityOnHand > 0 Then
                    PricingCell = PricingCell & "<div>Availability: In Stock</div>"
                Else
                    PricingCell = PricingCell & "<div>Availability: Back Order</div>"
                End If
            End If
            If PricingCell <> "" Then
                PricingCell = "" _
                    & "<div>&nbsp;</div>" _
                    & PricingCell
            End If
            '
            ' Options
            '
            Call GetOptionCell(main.GetCSText(CSPointer, "Options"), OptionCell, CheckScript, itemId)
            '
            ' ----- Pricing line
            '
            If RowDescriptor.CatalogAllowPricing Then
                PricingCell = PricingCell & Pricing.GetPricingCopy(main, CSPointer, RowDescriptor.CatalogAliasYourPrice, RowDescriptor.CatalogAliasRegularPrice, RowDescriptor.CatalogAliasSalePrice, RowDescriptor.CatalogAliasRetailPrice, RowDescriptor.CatalogAliasSpecialPrice, RowDescriptor.CatalogAllowSpecialPrice, CommerceAccountMemberID, AllowGroupPolicies)
            End If
            '
            '   now also checking if the group a user in can order
            '
            'CatalogAllowOrdering = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowOrdering", True))
            CatalogAllowOrdering = allowOrdering
            HideItemOrdering = main.GetCSBoolean(CSPointer, "HideOrderButton")
            
            If (CatalogAllowOrdering) And (Not HideItemOrdering) Then
                QueryString = main.RefreshQueryString
                PricingCell = PricingCell _
                    & main.GetFormStart(QueryString)
                PricingCell = PricingCell _
                    & "<div>&nbsp;</div>" _
                    & OptionCell _
                    & ""
                PricingCell = PricingCell _
                    & main.GetFormInputHidden(RequestNameCatalogOrder, CStr(itemId)) _
                    & "<div>&nbsp;</div>" _
                    & main.GetFormButton(ButtonOrder, , , CheckScript) _
                    & main.GetFormEnd _
                    & ""
            End If
            Cell = Cell & "<P>" & PricingCell & "</p>"
            Cell = Cell & getAddedMaterials()
            '
            ' Full size image
            '
Dim altSizeFound As Boolean
Dim imageFilename As String
Dim TestFilename As String
Dim WidthKnown As Boolean
Dim listFix As String
            baseFilename = main.GetCSText(CSPointer, "ImageFilename")
            altSizeList = main.GetCS(CSPointer, "altSizeList")
            If baseFilename <> "" Then
                WidthKnown = False
                Filename = baseFilename
                fileExt = Right(baseFilename, Len(baseFilename) - (InStrRev(baseFilename, ".", , vbTextCompare) - 1))
                imgWidth = RowDescriptor.CatalogImageWidth
                altSizeFound = False
                If (altSizeList <> "") And (InStr(1, altSizeList, imgWidth & "x") <> 0) Then
                    sizArr = Split(altSizeList, ",")
                    For sizPtr = 0 To UBound(sizArr)
                        If (InStr(1, listFix, sizArr(sizPtr)) <= 0) Then
                            '
                            ' not repeat, add to fix list and check it
                            '
                            listFix = listFix & "," & sizArr(sizPtr)
                            If (Not altSizeFound) And (InStr(1, sizArr(sizPtr), imgWidth & "x") <> 0) Then
                                imageFilename = Replace(baseFilename, fileExt, "-" & sizArr(sizPtr) & fileExt)
                                altSizeFound = True
                                WidthKnown = True
                                'Exit For
                            End If
                        End If
                    Next
                    If listFix <> "" Then
                        listFix = Mid(listFix, 2)
                    End If
                    If listFix <> altSizeList Then
                        Call main.SetCS(CSPointer, "altSizeList", listFix)
                    End If
                End If
                If Not altSizeFound Then
                    imageFilename = getResizedImage(baseFilename, CStr(imgWidth), CSPointer, "altSizeList")
                    If imageFilename = "" Then
                        '
                        ' resize failed - either bad image or the original was smaller then the resize
                        '
                        imageFilename = baseFilename
                    Else
                        '
                        ' resize worked, we know the width
                        '
                        WidthKnown = True
                    End If
                End If
                TestFilename = main.PhysicalFilePath & imageFilename
                TestFilename = Replace(main.PhysicalFilePath & imageFilename, "/", "\")
                If Not objFile.CheckFile(TestFilename) Then
                    imageFilename = baseFilename
                    WidthKnown = False
                End If
                Cell = Cell & "<div id=""scMainImage""><a href=""" & main.serverFilePath & baseFilename & """ name=Image>"
                If WidthKnown Then
                    Cell = Cell & "<img border=0 src=""" & main.serverFilePath & imageFilename & """ width=""" & imgWidth & """>"
                Else
                    Cell = Cell & "<img border=0 src=""" & main.serverFilePath & imageFilename & """>"
                End If
                Cell = Cell & "</a></div>"
            End If
'            baseFilename = Main.GetCSText(CSPointer, "ImageFilename")
'            If baseFilename <> "" Then
'                Filename = baseFilename
'                fileExt = Right(baseFilename, Len(baseFilename) - (InStrRev(baseFilename, ".", , vbTextCompare) - 1))
'                altSizeList = Main.GetCS(CSPointer, "altSizeList")
'                imgWidth = RowDescriptor.CatalogImageWidth
'                If (altSizeList <> "") And (InStr(1, altSizeList, imgWidth & "x") <> 0) Then
'                    If InStr(1, altSizeList, ",") <> 0 Then
'                        sizArr = Split(altSizeList, ",")
'                        For sizPtr = 0 To UBound(sizArr)
'                            If InStr(1, sizArr(sizPtr), imgWidth & "x") <> 0 Then
'                                Filename = Replace(baseFilename, fileExt, "-" & sizArr(sizPtr) & fileExt)
'                            End If
'                        Next
'                    Else
'                        Filename = Replace(baseFilename, fileExt, "-" & altSizeList & fileExt)
'                    End If
'                Else
'                    Filename = getResizedImage(baseFilename, CStr(imgWidth), CSPointer, "altSizeList")
'                End If
'
'                If Filename <> "" Then
'                    Cell = Cell & "<div id=""scMainImage""><a href=""" & Main.ServerFilePath & baseFilename & """ name=Image><img border=0 src=""" & Main.ServerFilePath & Filename & """ ></a></div>"
'                End If
'            End If
            stream.Add ("<table border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%""><tr>")
            stream.Add ("<td width=""100%"" valign=""top"" align=""left"">" & Cell & "</td></tr></table >")
            stream.Add (main.GetSiteProperty("TextSearchEndTag", TextSearchEndTagDefault))
        End If
        Call main.CloseCS(CSPointer)
    End If
    '
    'Call 'Main.'TestPointExit
    '
    GetCatalog2_DetailsForm = stream.Text
    '
    ' un-block content spider for just the detail page
    ' and root catalog page which has a link list to all detail pages
    '
    GetCatalog2_DetailsForm = "" _
        & cr & "<!-- TextSearchStart -->" _
        & kmaIndent(GetCatalog2_DetailsForm) _
        & cr & "<!-- TextSearchEnd -->"
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCatalog2_DetailsForm"): Resume Next
    '
End Function
'
'===============================================================================
' Print a Listing category header
'===============================================================================
'
Private Function GetCategoryHeader(categoryID As Long, ItemContentName As String, LinkCurrentCategory As Boolean) As String
    On Error GoTo ErrorTrap
    '
    If Not main.StreamOpen Then
        Exit Function
        End If
    '
    GetCategoryHeader = "" _
        & "<div class=ccCatRowHeader>" _
        & GetCategoryHeaderBranch(ItemContentName, categoryID, LinkCurrentCategory) _
        & "</div>"
    
    'GetCategoryHeader = GetCategoryHeader & ("<table width=""100%"" border=""0"" class=""ccCatRowHeader"" style=""padding: 4px;""><tr><td>")
    'GetCategoryHeader = GetCategoryHeader & ("")
    'GetCategoryHeader = GetCategoryHeader & ("</td></tr></table >")
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCategoryHeader"): Resume Next
    '
    End Function
'
'===============================================================================
'   Prints the list of category names up to the index
'===============================================================================
'
Private Function GetCategoryHeaderBranch(ContentName As String, categoryID As Long, LinkCurrentCategory As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Dim SQL As String
    Dim CategoryName As String
    Dim ParentID As Long
    Dim ps As String
    Dim TempCategoryID As Long
    Dim Id As Long
    Dim Name As String
    Dim QueryString As String
    Dim BreadCrumbPrefix As String
    Dim BreadCrumbDelimiter As String
    '
    ' Init defaults
    '
    If (categoryID > 0) Then
        TempCategoryID = 0
    Else
        TempCategoryID = Int(categoryID)
        End If
    BreadCrumbPrefix = main.GetSiteProperty("BreadCrumbPrefix", "Return To")
    BreadCrumbDelimiter = main.GetSiteProperty("BreadCrumbDelimiter", "&nbsp;&gt;&nbsp;")
    '
    ' put index on pagestring
    '
    QueryString = main.RefreshQueryString
    QueryString = ModifyQueryString(QueryString, rnSrcShopFormId, "")
    QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, "")
    'QueryString = ModifyQueryString(QueryString, rnSrcShopFormId, shopFormCatalogListing, True)
    ps = BreadCrumbPrefix & " <A class=cccatrowheader href=""" & GetShoppingCartLink(main.ServerPage, QueryString) & """>Categories</A>"
    '
    ' Put the tree on the pagestring
    '
    If (categoryID > 0) Then
        '
        SQL = "SELECT ItemCategories.ID AS ID0, ItemCategories.Name AS name0, ItemCategories_1.ID AS id1, ItemCategories_1.Name AS name1, ItemCategories_2.ID AS ID2, ItemCategories_2.Name AS name2, ItemCategories_3.ID AS ID3, ItemCategories_3.Name AS Name3, ItemCategories_4.ID AS ID4, ItemCategories_4.Name AS Name4" _
            & " FROM (((ItemCategories LEFT JOIN ItemCategories AS ItemCategories_1 ON ItemCategories.CategoryID = ItemCategories_1.ID) LEFT JOIN ItemCategories AS ItemCategories_2 ON ItemCategories_1.CategoryID = ItemCategories_2.ID) LEFT JOIN ItemCategories AS ItemCategories_3 ON ItemCategories_2.CategoryID = ItemCategories_3.ID) LEFT JOIN ItemCategories AS ItemCategories_4 ON ItemCategories_3.CategoryID = ItemCategories_4.ID" _
            & " WHERE (((ItemCategories.ID)=" & categoryID & "));"
        cs = main.OpenCSSQL("Default", SQL)
        If main.IsCSOK(cs) Then
            If True Then
                Id = (main.GetCSInteger(cs, "ID4"))
                Name = (main.GetCSText(cs, "Name4"))
                QueryString = main.RefreshQueryString
                QueryString = ModifyQueryString(QueryString, RequestNameCatalogItemID, "")
                QueryString = ModifyQueryString(QueryString, rnSrcShopFormId, shopFormCatalogListing, True)
                If Id <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, CStr(Id), True)
                    ps = ps & BreadCrumbDelimiter & "<A class=cccatrowheader href=""" & GetShoppingCartLink(main.ServerPage, QueryString) & """>" & Name & "</A>"
                End If
                Id = (main.GetCSInteger(cs, "ID3"))
                Name = (main.GetCSText(cs, "Name3"))
                If Id <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, CStr(Id), True)
                    ps = ps & BreadCrumbDelimiter & "<A class=cccatrowheader href=""" & GetShoppingCartLink(main.ServerPage, QueryString) & """>" & Name & "</A>"
                End If
                Id = (main.GetCSInteger(cs, "ID2"))
                Name = (main.GetCSText(cs, "Name2"))
                If Id <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, CStr(Id), True)
                    ps = ps & BreadCrumbDelimiter & "<A class=cccatrowheader href=""" & GetShoppingCartLink(main.ServerPage, QueryString) & """>" & Name & "</A>"
                End If
                Id = (main.GetCSInteger(cs, "ID1"))
                Name = (main.GetCSText(cs, "Name1"))
                If Id <> 0 Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, CStr(Id), True)
                    ps = ps & BreadCrumbDelimiter & "<A class=cccatrowheader href=""" & GetShoppingCartLink(main.ServerPage, QueryString) & """>" & Name & "</A>"
                End If
                Id = (main.GetCSInteger(cs, "ID0"))
                Name = (main.GetCSText(cs, "Name0"))
                If LinkCurrentCategory Then
                    QueryString = ModifyQueryString(QueryString, RequestNameCatalogCategoryID, CStr(Id), True)
                    ps = ps & BreadCrumbDelimiter & "<A class=cccatrowheader href=""" & GetShoppingCartLink(main.ServerPage, QueryString) & """>" & Name & "</A>"
                Else
                    ps = ps & BreadCrumbDelimiter & "<B>" & Name & "</B>"
                End If
            End If
        End If
        Call main.CloseCS(cs)
    End If
    GetCategoryHeaderBranch = ps
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCategoryHeaderBranch"): Resume Next
    End Function
'
'===============================================================================
'   Build the sql AND statements to include the categoryID and all children of it
'===============================================================================
'
Private Function GetItemCategoryCriteria(categoryID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim itemId As Long
    '
    GetItemCategoryCriteria = ""
    If categoryID <> 0 Then
    cs = main.OpenCSContent_Internal("Item Categories", "(CategoryID=" & categoryID & ")")
        Do While main.IsCSOK(cs)
            itemId = (main.GetCSInteger(cs, "ID"))
            GetItemCategoryCriteria = GetItemCategoryCriteria & "or(Items.CategoryID=" & kmaEncodeText(itemId) & ")"
            GetItemCategoryCriteria = GetItemCategoryCriteria & GetItemCategoryCriteria(itemId)
            main.NextCSRecord (cs)
        Loop
    End If
    Call main.CloseCS(cs)
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetItemCategoryCriteria"): Resume Next
    '
End Function
'
'=============================================================================
'   Catalog Search Form
'=============================================================================
'
Public Function GetCatalog_SearchForm(ItemContentName As String, CategoryContentName As String) As String
    GetCatalog_SearchForm = GetCatalog2_SearchForm(ItemContentName, CategoryContentName)
End Function
'
Friend Function GetCatalog2_SearchForm(ItemContentName As String, CategoryContentName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim Copy As String
    Dim RecordID As Long
    Dim SelectedFound As Boolean
    Dim qs As String
    Dim CatalogAllowMfg As Boolean
    Dim s As String
    Dim CatalogMfgPhrase As String
    Dim CmrcCopy As New CmrcCopyClass
    Dim selector As String
    '
    If main.StreamOpen Then
        Call main.AddPageTitle("Catalog Search")
        s = s & CmrcCopy.GetCopy(main, shopCopyNameSearchHome, shopCopyDefaultSearchHome)
        '
        qs = main.RefreshQueryString
        qs = ModifyQueryString(qs, rnSrcShopFormId, shopFormCatalogListing, True)
        qs = ModifyQueryString(qs, RequestNameCatalogKeywords, "", False)
        qs = ModifyQueryString(qs, RequestNameCatalogCategoryID, "", False)
        qs = ModifyQueryString(qs, RequestNameCatalogItemID, "", False)
        s = s & (main.GetFormStart(qs))
        s = s & ("<table class=""scSearchTable"" border=""0"" cellpadding=""5"" cellspacing=""0"" width=""100%""><tr>")
        s = s & ("<td valign=""baseline"" width=""150"" align=""right"">Keywords&nbsp;</td>")
        s = s & ("<td valign=""baseline"" width=""90%"" align=left><input type=""text"" name=""" & RequestNameCatalogKeywords & """ size=""20"" value=""" & SearchKeywords & """></td>")
        s = s & "</tr>"
        '
        SQL = "select top 1 id from itemcategories where active<>0"
        cs = main.OpenCSSQL("", SQL)
        If main.IsCSOK(cs) Then
            selector = main.GetFormInputSelect(RequestNameCatalogCategoryID, SearchCategoryID, "Item Categories", , "Any")
            s = s & "<tr>"
            s = s & ("<td valign=""baseline"" width=""150"" align=""right"">Category&nbsp;</td>")
            s = s & ("<td valign=""baseline"" width=""90%"" align=left>" & selector & "</td>")
            s = s & ("</tr>")
        End If
        Call main.CloseCS(cs)
        '
        CatalogAllowMfg = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowMfg", False))
        If CatalogAllowMfg Then
            CatalogMfgPhrase = main.GetSiteProperty("CatalogMfgPhrase", "Mfg By ")
            s = s & ("<tr>")
            s = s & ("<td valign=""baseline"" width=""150"" align=""right"">" & scSpanNormalStyle & CatalogMfgPhrase & "</SPAN></td>")
            s = s & ("<td width=""90%"">")
            SQL = "SELECT Organizations.Name, Organizations.ID" _
                    & " FROM Items LEFT JOIN Organizations ON Items.OrganizationID = Organizations.ID" _
                    & " GROUP BY Organizations.Name, Organizations.ID" _
                    & " Having (((Organizations.Name) Is Not Null))" _
                    & " ORDER BY Organizations.Name;"
            cs = main.OpenCSSQL("Default", SQL)
            If main.IsCSOK(cs) Then
                If True Then
                    s = s & (vbLf & "<SELECT size=""1"" NAME=" & RequestNameCatalogSource & ">")
                    SelectedFound = False
                    Do While main.IsCSOK(cs)
                        RecordID = (main.GetCSInteger(cs, "ID"))
                        Copy = (main.GetCSText(cs, "name"))
                        Copy = main.EncodeContent(Copy, CommerceAccountMemberID)
                        s = s & (vbLf & "<option value=""" & RecordID & """ ")
                        If RecordID = SearchManufacturerID Then
                            s = s & ("selected")
                            SelectedFound = True
                        End If
                        s = s & (">" & Copy & "</option>")
                        main.NextCSRecord (cs)
                    Loop
                    If Not SelectedFound Then
                        s = s & ("vblf & <option value="""" selected>Any</option>")
                    Else
                        s = s & ("vblf & <option value=""-1"">Any</option>")
                    End If
                    s = s & ("vblf & </SELECT>")
                End If
            End If
            Call main.CloseCS(cs)
            s = s & ("</td>")
            s = s & ("</tr>")
        End If
        s = s & "<tr>"
        s = s & "<td colspan=""2"">&nbsp;</td>"
        s = s & "</tr>"
        s = s & "<tr>"
        s = s & "<td colspan=""2"" class=""ccButtonCon""><input type=""submit"" value=""Search"" name=""action""></td>"
        s = s & "</tr></table >"
        s = s & "</FORM>"
        End If
    GetCatalog2_SearchForm = s
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCatalog2_SearchForm"): Resume Next
    '
    End Function
'
'=============================================================================
' Print the catalog Listing form
'=============================================================================
'
Public Sub PrintCatalogNewListing(ItemContentName As String, CategoryContentName As String)
    On Error GoTo ErrorTrap
    '
    Dim RowDescriptor As RowDescriptorType
    Dim SQL As String
    Dim cs As Long
    Dim ItemCount As Long
    Dim itemId As Long
    Dim isInCatalog As Boolean
    '
    If main.StreamOpen Then
        RowDescriptor = GetRowDescriptor()
        itemId = kmaEncodeInteger(main.GetStreamInteger(RequestNameCatalogItemID))
        If (itemId <> 0) Then
            Call GetCatalog2_DetailsForm(itemId, ItemContentName, CategoryContentName)
        Else
            ItemCount = 0
            SQL = "SELECT Items.ID, Items.Name, Items.ItemNumber, Items.SourceItemNumber, Items.CategoryID, Items.ImageFilename, Items.ThumbnailFilename, Items.UnitCost, Items.UnitPrice, Items.OnSale, Items.SalePrice, Items.SpecialPrice, Items.Featured, Items.Weight, Items.QuantityOnhand, Items.QuantityLow, Organizations.Name AS SourceName, Items.BriefFilename, Items.CopyFilename, Items.Copy AS Copy, Items.OverView AS OverView, Items.UPC as UPC, Items.Options as Options, Items.altThumbSizeList as altThumbSizeList, Items.altSizeList as altSizeList, Items.HideOrderButton" _
                & " FROM Items LEFT JOIN Organizations ON Items.OrganizationID = Organizations.ID" _
                & " WHERE" _
                & " (Items.New<>" & KmaEncodeSQLBoolean(SQLFalse) & ")" _
                & " AND(Items.Active<>" & KmaEncodeSQLBoolean(SQLFalse) & ")"
            If main.IsContentFieldSupported("items", "isInCatalog") Then
                SQL = SQL & "and(isInCatalog=1)"
            End If
            cs = main.OpenCSSQL("Default", SQL)
            If main.IsCSOK(cs) Then
                main.WriteStream ("<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" class=""ccPanel"">")
                main.WriteStream ("<tr><td>" & scSpanNormalStyle & "<B>New Items</B></SPAN></td></tr>")
                main.WriteStream ("</table >")
                If Not main.IsCSOK(cs) Then
                    main.WriteStream ("<P>There are currently no new items in this area.</p>")
                Else
                    Do While main.IsCSOK(cs) And ItemCount < 10
                        Call main.WriteStream(GetCatalog2_ListingForm_GetRow_NextRecord(cs, RowDescriptor))
                        ItemCount = ItemCount + 1
                        'If Main.IsCSOK(CS) Then
                        '    Main.WriteStream ("<img src=""/ccLib/images/808080.gif"" width=""100%"" height=""2"" >")
                        '    End If
                        Loop
                    End If
                End If
            Call main.CloseCS(cs)
            End If
        End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("PrintCatalogNewListing"): Resume Next
    '
    End Sub
'
Public Function GetCatalog(ItemContentName As String, CategoryContentName As String) As String
    GetCatalog = GetCatalog2(main.GetStreamInteger(rnSrcShopFormId), main.GetStreamInteger(rnDstShopFormId))
End Function
'
'=============================================================================
'
'=============================================================================
'
Friend Function GetCatalog2(srcShopFormId As Long, dstShopFormId As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim qs As String
    Dim Menu As String
    Dim Body As String
    Dim categoryList As String
    Dim isEditing As Boolean
    Dim Copy As New CmrcCopyClass
    '
    Const ItemContentName = "items"
    Const CategoryContentName = "item categories"
    '
    If main.StreamOpen Then
        '
        isEditing = main.IsEditingAnything()
        '
        ' Help Link
        '
        'Body = Body & main.GetHelpLink(16, "Using the Commerce Catalog", "The commerce catalog contains a complete commerce subsystem with both a catalog and order process system. To use the system, add catalog categories and catalog items, and set the commerce configuration, typically under the Administration menu > Admin Tools.") & s
        '
        ' Body
        '
        Select Case dstShopFormId
            Case shopFormCatalogListing, shopFormCatalogDetails
                Body = Body & GetCatalog2_ListingForm(ItemContentName, CategoryContentName, SearchCategoryID)
            Case shopFormCatalogSearch
                Body = Body & GetCatalog2_SearchForm(ItemContentName, CategoryContentName)
            Case Else
                Body = Body _
                    & Copy.GetCopy(main, shopCopyNameCatalogHome, shopCopyDefaultCatalogHome) _
                    & GetCatalog2_ListingForm(ItemContentName, CategoryContentName, SearchCategoryID)
        End Select
        '
        ' add categories to the bottom
        '
        categoryList = getCategoryList(isEditing)
        If categoryList <> "" Then
            Body = "" _
                & Body _
                & categoryList _
                & ""
        End If
'        '
'        ' Menu (below tabs)
'        '
'        Menu = "<div class=scSubMenu>" _
'            & "&nbsp;" _
'            & GetCategoryMenu(0) _
'            & "</div>"
        '
        GetCatalog2 = "" _
            & Menu _
            & vbCrLf & vbTab & "<div class=""ccCatBody"">" _
            & kmaIndent(Body) _
            & vbCrLf & vbTab & "</div>" _
            & ""
'        '
'        GetCatalog2 = "" _
'            & Menu _
'            & vbCrLf & vbTab & "<table border=0 cellpadding=0 cellspacing=0 width=100%>" _
'            & vbCrLf & vbTab & "<tr>" _
'            & vbCrLf & vbTab & "<td class=ccCatBody>" _
'            & vbCrLf & vbTab & "" & Body & "" _
'            & vbCrLf & vbTab & "</td>" _
'            & vbCrLf & vbTab & "</tr>" _
'            & vbCrLf & vbTab & "</table >" _
'            & vbCrLf & vbTab & ""
        '
        ' block content spider everywhere except the detail page
        ' and root catalog page which has a link list to all detail pages
        '
        GetCatalog2 = "" _
            & cr & "<!-- TextSearchEnd -->" _
            & kmaIndent(GetCatalog2) _
            & cr & "<!-- TextSearchStart -->"
    End If
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCatalog2"): Resume Next
    End Function
'
'
'
Private Sub Class_Initialize()
    ClassInitialized = False
    End Sub



Private Sub Class_Terminate()
    Set main = Nothing
    End Sub
'
'=============================================================================
' ----- Compatibility
' Print the catalog Listing form
'=============================================================================
'
Private Function Compatibility_GetCategoryListing(CategoryName As String, ItemContentName As String, CategoryContentName As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CSPointer As Long
    Dim categoryID As Long
    '
    If main.StreamOpen Then
        Select Case catalogForm
            Case shopFormCatalogDetails
                Compatibility_GetCategoryListing = Compatibility_GetCategoryListing & GetCatalog2_DetailsForm(itemId, ItemContentName, CategoryContentName)
            Case Else
                CSPointer = main.OpenCSContent_Internal(CategoryContentName, "name=" & KmaEncodeSQLText(CategoryName))
                If main.IsCSOK(CSPointer) Then
                    categoryID = (main.GetCSInteger(CSPointer, "id"))
                    Compatibility_GetCategoryListing = GetCatalog2_ListingForm(ItemContentName, CategoryContentName, categoryID)
                    'Compatibility_GetCategoryListing = Compatibility_GetCategoryListing_ByCategoryID(CategoryID, ItemContentName, CategoryContentName)
                    End If
                Call main.CloseCS(CSPointer)
            End Select
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("Compatibility_GetCategoryListing"): Resume Next
    '
    End Function
'
'=============================================================================
' ----- Compatibility
' Print the catalog Listing form
'=============================================================================
'
Public Sub PrintCatalogFeaturedListing(ItemContentName As String, CategoryContentName As String)
    On Error GoTo ErrorTrap
    '
    Dim RowDescriptor As RowDescriptorType
    Dim SQL As String
    Dim cs As Long
    Dim ItemCount As Long
    Dim itemId As Long
    '
    If main.StreamOpen Then
        RowDescriptor = GetRowDescriptor()
        itemId = kmaEncodeInteger(main.GetStreamInteger(RequestNameCatalogItemID))
        If (itemId <> 0) Then
            Call GetCatalog2_DetailsForm(itemId, ItemContentName, CategoryContentName)
        Else
            '
            ItemCount = 0
            SQL = "SELECT Items.ID, Items.Name, Items.ItemNumber, Items.SourceItemNumber, Items.CategoryID, Items.ImageFilename, Items.ThumbnailFilename, Items.UnitCost, Items.UnitPrice, Items.OnSale, Items.SalePrice, Items.SpecialPrice, Items.Featured, Items.Weight, Items.QuantityOnhand, Items.QuantityLow, Organizations.Name AS SourceName, Items.BriefFilename, Items.CopyFilename, Items.Copy AS Copy, Items.OverView AS OverView, Items.UPC as UPC, Items.Options as Options, Items.altThumbSizeList as altThumbSizeList, Items.altSizeList as altSizeList, Items.HideOrderButton" _
                & " FROM Items LEFT JOIN Organizations ON Items.OrganizationID = Organizations.ID" _
                & " WHERE " _
                & " (Items.Featured<>" & KmaEncodeSQLBoolean(SQLFalse) & ")" _
                & " AND(Items.Active<>" & KmaEncodeSQLBoolean(SQLFalse) & ")"
            If main.IsContentFieldSupported("items", "isInCatalog") Then
                SQL = SQL & "and(isInCatalog=1)"
            End If
            
            cs = main.OpenCSSQL("Default", SQL)
            If main.IsCSOK(cs) Then
                If main.IsCSOK(cs) Then
                    main.WriteStream ("<table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""0"" class=""ccPanel"">")
                    main.WriteStream ("<tr><td>" & scSpanNormalStyle & "<B>Featured Items</B></SPAN></td></tr>")
                    main.WriteStream ("</table >")
                    End If
                Do While main.IsCSOK(cs) And ItemCount < 10
                    Call main.WriteStream(GetCatalog2_ListingForm_GetRow_NextRecord(cs, RowDescriptor))
                    ItemCount = ItemCount + 1
                    'If Main.IsCSOK(CS) Then
                    '    Main.WriteStream ("<img src=""/ccLib/images/808080.gif"" width=""100%"" height=""2"" >")
                    '    End If
                    Loop
                End If
            Call main.CloseCS(cs)
            End If
        End If
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("PrintCatalogFeaturedListing"): Resume Next
    '
    End Sub
'
'=============================================================================
' ----- Compatibility
'   Catalog Search Form
'=============================================================================
'
Public Sub PrintSearchForm(ItemContentName As String, CategoryContentName As String)
    main.WriteStream (GetCatalog2_SearchForm(ItemContentName, CategoryContentName))
    End Sub
'
'=============================================================================
' ----- Compatibility
' Print the catalog Listing form
'=============================================================================
'
Public Sub PrintCategoryListing(CategoryName As String, ItemContentName As String, CategoryContentName As String)
    main.WriteStream (Compatibility_GetCategoryListing(CategoryName, ItemContentName, CategoryContentName))
    End Sub
'
'=============================================================================
' ----- Compatibility
' Print the standard catalog
'=============================================================================
'
Public Sub PrintCatalog(ItemContentName As String, CategoryContentName As String)
    main.WriteStream (GetCatalog(ItemContentName, CategoryContentName))
    End Sub
'
'
'
Private Function GetCategoryMenu(ParentCategoryID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Dim SQL As String
    Dim RecordID As Long
    Dim ParentID As Long
    Dim qs As String
    '
    Call main.AddMenuEntry("P0", "", , , , "Shop By Category")
    qs = main.RefreshQueryString
    qs = ModifyQueryString(qs, RequestNameCatalogItemID, "", False)
    qs = ModifyQueryString(qs, RequestNameCatalogCategoryID, "", False)
    qs = ModifyQueryString(qs, rnSrcShopFormId, shopFormCatalogListing, True)
    qs = "?" & qs & "&" & RequestNameCatalogCategoryID & "="
    Call main.AddMenuEntry("C0", "P0", , , qs & "0", "All")
    SQL = "select ID,name,categoryid from ItemCategories where active<>0"
    cs = main.OpenCSSQL("default", SQL)
    Do While main.IsCSOK(cs)
        RecordID = main.GetCSInteger(cs, "ID")
        ParentID = main.GetCSInteger(cs, "CategoryID")
        Call main.AddMenuEntry("C" & RecordID, "P" & ParentID, , , qs & RecordID, main.GetCSText(cs, "name"))
        main.NextCSRecord (cs)
    Loop
    main.CloseCS (cs)
    GetCategoryMenu = main.GetMenu("P0", 1, "ccCatMenu")
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetCategoryMenu"): Resume Next
    '
End Function
'
'
'
Private Function getCategoryList(isEditing As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim cs As Long
    Dim SQL As String
    Dim RecordID As Long
    Dim ParentID As Long
    Dim qs As String
    Dim s As String
    Dim idString As String
    Dim listCaption As String
    '
    s = ""
    qs = main.RefreshQueryString
    qs = ModifyQueryString(qs, RequestNameCatalogItemID, "", False)
    qs = ModifyQueryString(qs, rnSrcShopFormId, "", True)
    qs = ModifyQueryString(qs, rnDstShopFormId, shopFormCatalogListing, True)
    '
    SQL = "select ID,name from ItemCategories where (active<>0)and(categoryId is null) order by name"
    cs = main.OpenCSSQL("default", SQL)
    Do While main.IsCSOK(cs)
        idString = main.GetCSText(cs, "ID")
        qs = ModifyQueryString(qs, RequestNameCatalogCategoryID, idString, True)
        listCaption = Trim(main.GetCSText(cs, "name"))
        If listCaption = "" Then
            listCaption = "Category " & idString
        End If
        listCaption = kmaEncodeHTML(listCaption)
        listCaption = "<a href=""?" & qs & """>" & listCaption & "</a>"
        If isEditing Then
            listCaption = main.GetRecordEditLink("item categories", kmaEncodeInteger(idString)) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & listCaption
        End If
        s = s _
            & vbCrLf & vbTab & "<li>" & listCaption & "</li>"
        main.NextCSRecord (cs)
    Loop
    main.CloseCS (cs)
    '
    ' if editing, include an add category item
    '
    If isEditing Then
        s = s _
            & vbCrLf & vbTab & "<li>" & main.GetRecordAddLink("item categories", "") & "Add a Category</li>"
    End If
    '
    If s <> "" Then
        s = "" _
            & vbCrLf & vbTab & "<h2 class=""scCategoryHead"">Categories</h2>" _
            & vbCrLf & vbTab & "<ul class=""scCategoryList"">" _
            & vbCrLf & vbTab & "<li><a href=""?" & ModifyQueryString(qs, RequestNameCatalogCategoryID, "-1") & """>Any</a></li>" _
            & kmaIndent(s) _
            & vbCrLf & vbTab & "</ul>" _
            & ""
    End If
    getCategoryList = s
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCategoryList"): Resume Next
End Function
'
'
'
Private Function GetFormInputSelectListLocal(Name As String, CurrentValue As String, OptionList As String, NoneCaption As String, Id As String) As String
    On Error GoTo ErrorTrap
    '
    Dim s As String
    Dim Ptr As Long
    Dim Options() As String
    Dim OptionPtr As Long
    '
    Options = Split(OptionList, ",")
    For Ptr = 0 To UBound(Options)
        If Trim(Options(Ptr)) <> "" Then
            s = s & "<option>" & Trim(Options(Ptr)) & "</option>"
        End If
    Next
    If s <> "" Then
        s = "<select size=1 name=""" & Name & """ ID=""" & Id & """><option value=""-1"">" & NoneCaption & "</option>" & s & "</select>"
    End If
    '
    GetFormInputSelectListLocal = s
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetFormInputSelectListLocal")
    End Function
'
'
'
Private Sub GetOptionCell(OptionList As String, ByRef OptionCell As String, ByRef CheckScript As String, itemId As Long)
    On Error GoTo ErrorTrap
    '
    Dim OptionRowCaption As String
    Dim OptionRowList As String
    Dim optionCnt As Long
    Dim OptionSelect As String
    Dim OptionRows() As String
    Dim RowPtr As Long
    Dim OptionRow As String
    Dim Pos As Long
    '
    optionCnt = 0
    If OptionList <> "" Then
        OptionCell = "" _
            & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" >"
        OptionRows = Split(OptionList, vbCrLf)
        For RowPtr = 0 To UBound(OptionRows)
            OptionRow = Trim(OptionRows(RowPtr))
            If OptionRow <> "" Then
                Pos = InStr(1, OptionRow, ":")
                If Pos < 2 Then
                    OptionRowCaption = "Options"
                    OptionRowList = OptionRow
                Else
                    OptionRowCaption = Mid(OptionRow, 1, Pos - 1)
                    'OptionRowCaption = Replace(OptionRowCaption, " ", "&nbsp;")
                    OptionRowList = Mid(OptionRow, Pos + 1)
                End If
                '
                OptionSelect = "" _
                    & main.GetFormInputHidden("OptionCaption" & optionCnt, OptionRowCaption) _
                    & GetFormInputSelectListLocal("Option_" & itemId & "_" & optionCnt, "", OptionRowList, "Select One", "Option_" & itemId & "_" & optionCnt)
                OptionSelect = Replace(OptionSelect, "value=""""", "value=""-1""", , , vbTextCompare)
                'OptionSelect = Replace(OptionSelect, "<select ", "<select style=""height:10px;"" ", , , vbTextCompare)
                CheckScript = CheckScript _
                    & vbCrLf & "var e" & optionCnt & "=document.getElementById('Option_" & itemId & "_" & optionCnt & "');" _
                    & vbCrLf & "if ( e" & optionCnt & ".value=='-1' ) {" _
                    & vbCrLf & "alert('Please select an option for " & OptionRowCaption & ".');" _
                    & vbCrLf & "return false;" _
                    & vbCrLf & "}"
                '
                OptionCell = OptionCell & "<tr>"
                OptionCell = OptionCell & "<td width=1 align=left>" & Replace(OptionRowCaption, " ", "&nbsp;") & "&nbsp;</td>"
                OptionCell = OptionCell & "<td width=99% align=left>" & OptionSelect & "</td>"
                OptionCell = OptionCell & "</tr>"
                '
                optionCnt = optionCnt + 1
            End If
        Next
        OptionCell = OptionCell & main.GetFormInputHidden("OptionCnt", optionCnt)
        If CheckScript <> "" Then
            OptionCell = OptionCell & "" _
                & vbCrLf & "<script>" _
                & "function CheckOptions_" & itemId & "() {" _
                & CheckScript _
                & "}" _
                & vbCrLf & "</script>"
            CheckScript = "return CheckOptions_" & itemId & "();"
        End If
        OptionCell = OptionCell & "</table>"
    End If
    '
    Exit Sub
ErrorTrap:
    Call HandleClassErrorAndResumeNext("GetOptionCell")
End Sub
'
'
'
Private Function GetRowDescriptor() As RowDescriptorType
    With GetRowDescriptor
        .CatalogAllowPricing = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowPricing", "true"))
        .CatalogAllowInventory = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowInventory", "false"))
        .CatalogAllowSpecialPrice = kmaEncodeBoolean(main.GetSiteProperty("CatalogAllowSpecialPrice", "false"))
        '
        .CatalogListingColumns = kmaEncodeInteger(main.GetSiteProperty("CatalogListingColumns", 1))
        If .CatalogListingColumns = 0 Then
            .CatalogListingColumns = 1
            Call main.SetSiteProperty("CatalogListingColumns", .CatalogListingColumns, True)
        End If
        .CatalogThumbnailWidth = kmaEncodeInteger(main.GetSiteProperty("CatalogThumbnailWidth", 100))
        If .CatalogThumbnailWidth = 0 Then
            .CatalogThumbnailWidth = 100
            Call main.SetSiteProperty("CatalogThumbnailWidth", .CatalogThumbnailWidth, True)
        End If
        .CatalogImageWidth = kmaEncodeInteger(main.GetSiteProperty("CatalogImageWidth", 400))
        If .CatalogImageWidth = 0 Then
            .CatalogImageWidth = 100
            Call main.SetSiteProperty("CatalogImageWidth", .CatalogImageWidth, True)
        End If
        '
        .CatalogAliasRegularPrice = main.GetSiteProperty("CatalogAliasRegularPrice", "Price")
        If .CatalogAliasRegularPrice = "" Then
            .CatalogAliasRegularPrice = "Price"
            Call main.SetSiteProperty("CatalogAliasRegularPrice", .CatalogAliasRegularPrice, True)
        End If
        '
        .CatalogAliasSpecialPrice = main.GetSiteProperty("CatalogAliasSpecialPrice", "Member Price")
        If .CatalogAliasSpecialPrice = "" Then
            .CatalogAliasSpecialPrice = "Member Price"
            Call main.SetSiteProperty("CatalogAliasSpecialPrice", .CatalogAliasSpecialPrice, True)
        End If
        '
        .CatalogAliasSalePrice = main.GetSiteProperty("CatalogAliasSalePrice", "Sale")
        If .CatalogAliasSalePrice = "" Then
            .CatalogAliasSalePrice = "Sale"
            Call main.SetSiteProperty("CatalogAliasSalePrice", .CatalogAliasSalePrice, True)
        End If
        '
        .CatalogAliasRetailPrice = main.GetSiteProperty("CatalogAliasRetailPrice", "Retail Price")
        '
        .CatalogAliasYourPrice = main.GetSiteProperty("CatalogAliasYourPrice", "Your Price")
        If .CatalogAliasYourPrice = "" Then
            .CatalogAliasYourPrice = "Your Price"
            Call main.SetSiteProperty("CatalogAliasYourPrice", .CatalogAliasYourPrice, True)
        End If
    End With
End Function

Private Function getCategory(categoryID As Long) As String
    On Error GoTo ErrorTrap
    
    Dim qs As String
    Dim stream As String
    Dim cs As Long
    Dim parCat As Long
    
    qs = main.RefreshQueryString
    qs = ModifyQueryString(qs, "CatalogItemID", "")
    qs = ModifyQueryString(qs, rnSrcShopFormId, "")
    qs = ModifyQueryString(qs, "CatalogFormID", 2, True)
    qs = ModifyQueryString(qs, "CatalogCatID", CStr(categoryID), True)
    
    cs = main.OpenCSContent("Item Categories", "ID=" & categoryID, , , , , "Name,CategoryID")
    If main.CSOK(cs) Then
        parCat = main.GetCSInteger(cs, "CategoryID")
        If parCat <> 0 Then
            stream = stream & getCategory(parCat) & " > "
        End If
        stream = stream & "<a href=?" & qs & ">" & main.GetCSText(cs, "Name") & "</a>"
    End If
    Call main.CloseCS(cs)
    
    getCategory = stream
    
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCategoryList")
End Function

Private Function getAddedMaterials() As String
    On Error GoTo ErrorTrap
    
    Dim stream As String
    Dim cs As Long
    Dim thsCat As String
    Dim lstCat As String
    Dim innerStream As String
    
    cs = main.OpenCSContent("Item Materials", "ItemID=" & itemId, "ItemMaterialsCategoryID")
    If main.CSOK(cs) Then
        stream = "<div class=""scItemMaterials"">"
        stream = stream & "<h3>Additional Information about this product</h3>"
        Do While main.CSOK(cs)
            thsCat = main.GetRecordName("Item Material Categories", main.GetCSInteger(cs, "ItemMaterialsCategoryID"))
            If thsCat <> lstCat Then
                If innerStream <> "" Then
                    innerStream = innerStream & "</ul>"
                End If
                innerStream = innerStream & "<ul>"
                innerStream = innerStream & "<li class=""scMaterialCategory"">" & thsCat & "</li>"
                lstCat = thsCat
            End If
            innerStream = innerStream & "<li class=""scMaterialItem""><a target=""_blank"" href=""" & main.serverFilePath & main.GetCS(cs, "docFileName") & """>" & main.GetCS(cs, "Name") & "</a></li>"
            Call main.NextCSRecord(cs)
        Loop
        innerStream = innerStream & "</ul>"
        stream = stream & innerStream & "</div>"
    End If
    Call main.CloseCS(cs)
    
    stream = stream
    
    getAddedMaterials = stream
    
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getCategoryList")
End Function

Private Function getResizedImage(imgSource As String, imgWidth As String, Pointer As Long, szField As String) As String
    On Error GoTo ErrorTrap
    
    Dim optionString As String
    Dim doc As New MSXML2.DOMDocument
    Dim retImg As String
    Dim imgHeight As String
    Dim altSizeList As String
    Dim imageSrc As String
    Dim imageFilename As String
    
    
    '
    imageFilename = main.PhysicalFilePath & Replace(imgSource, "/", "\")
    optionString = "Image Source=" & main.encodeNvaArgument(imageFilename)
    optionString = optionString & "&Image Width=" & main.encodeNvaArgument(imgWidth)
    '
    '   {59A5FCF4-AC87-4137-8966-B12B66666B56} - Resize Image out of Primitives
    '
    If doc.loadXML(main.ExecuteAddon(0, "{59A5FCF4-AC87-4137-8966-B12B66666B56}", optionString, ContextSimple, "", 0, "", "", -1)) Then
    
        If doc.hasChildNodes Then
            retImg = doc.getElementsByTagName("newImage").Item(0).Text
            If retImg <> "" Then
                imgHeight = doc.getElementsByTagName("imageHeight").Item(0).Text
                imgWidth = doc.getElementsByTagName("imageWidth").Item(0).Text
                If main.CSOK(Pointer) Then
                    altSizeList = main.GetCSText(Pointer, szField)
                    If altSizeList <> "" Then
                        altSizeList = altSizeList & ","
                    End If
                    Call main.SetCS(Pointer, szField, altSizeList & imgWidth & "x" & imgHeight)
                    Call main.SaveCSRecord(Pointer)
                End If
                getResizedImage = Replace(retImg, main.PhysicalFilePath, "")
            End If
        End If
        
    End If
    
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("getResizedImage")
End Function
'
Private Function catalogHasFeaturedItems() As Boolean
    On Error GoTo ErrorTrap
    '
    Dim SQL As String
    Dim cs As Long
    Dim recCount As Integer
    '
    SQL = "select count(ID) as count FROM items where (featured<>0) and (active<>0)"
    cs = main.OpenCSSQL("default", SQL)
    If main.CSOK(cs) Then
        recCount = main.GetCSInteger(cs, "count")
    End If
    Call main.CloseCS(cs)
    '
    catalogHasFeaturedItems = (recCount <> 0)
    '
    Exit Function
ErrorTrap:
    Call HandleClassErrorAndResumeNext("catalogHasFeaturedItems")
End Function
'
'==========================================================================================
' handle error and resume next
'==========================================================================================
'
Private Sub HandleClassErrorAndResumeNext(MethodName As String, Optional Description As String, Optional ErrorNumber As Long)
    Call HandleErrorAndResumeNext("CatalogClass2", MethodName, Description, ErrorNumber)
End Sub

