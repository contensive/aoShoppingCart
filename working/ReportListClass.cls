VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportListClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'
' ----- Global Objects
'
Private main As Object
'
'-----------------------------------------------------------------------
' ----- report Globals - Loaded with LoadReportGlobals
'-----------------------------------------------------------------------
'
' ----- Global Variables
'
Private ReportGlobalsLoaded As Boolean      ' true after LoadReportGlobals has been called
'
Private ReportID As Long            ' The current Report being printed
Private ReportButton As String
Private RQS As String       ' The query string needed for a refresh of the current report
Const ReportMargin = 20
'
' ----- Values that modify reports
'
Private ReportDateFrom As Date
Private ReportDateTo As Date
Private ReportDateAfterTo As Date           ' The day after the DateTo
Private ReportVisitID As Long
Private ReportMemberID As Long
Private ReportVisitorID As Long
Private ReportExcludeBrowsers As Boolean
Private ReportExcludeIP As Boolean
Private ReportIPAddress As String
Private ReportExcludeNewVisitors As Boolean
Private ReportExcludeOldVisitors As Boolean
Private ReportRecordID As Long              ' Used when the report accesses a single record
Private ReportPageSize As Long
Private ReportPageNumber As Long
Private ReportLinkDepth As Long             ' 0 gives you a cancel button, 1+ give you a close button
Private ReportTimeout As Long
Private ReportVisitReferer As String
'
' ----- ReportID Constants, moved to ccCommonModule
'
Private FormResultsRowCount As Long


Const ReportNameRoot = "Reports"
'Const ReportNameDailyVisits = "Daily Visits Report"
'Const ReportNameWeeklyVisits = "Weekly Visits Report"
'Const ReportNameReferers = "Referring Sites Report"
''Const ReportNameBrowserList = "Browser Report"
'Const ReportNameAddressList = "IP Address Report"
'Const ReportNameSitePath = "Site Path Report"
'Const ReportNameSearchKeywords = "Search Keyword Report"
'Const ReportNameContentProperties = "Content Properties Report"
'Const ReportNameSurveyList = "Survey List Report"
'Const ReportNameSurveyDetails = "Survey Details Report"
Const ReportNameOrdersList = "Shopping Cart Orders"
''Const ReportNameOrderDetails = "Order Details Report"
'Const ReportNameMemberDetails = "Member Details Report"
'Const ReportNameVisitorList = "Visitor List Report"
'Const ReportNameVisitorDetails = "Visitor Details Report"
'Const ReportNameVisitList = "Visit List Report"
'Const ReportNameVisitDetails = "Visit Details Report"
'Const ReportNamePageList = "Page List Report"
'Const ReportNameSpiderDocList = "Spider Page List Report"
'Const ReportNameSpiderDocDetails = "Spider Page Detail Report"
'Const ReportNameSpiderErrorList = "Spider Errors Report"
'Const ReportNameEDGDocErrors = "EDG Publish Errors"
'Const ReportNameDownloadLog = "Download Report"
'Const ReportNameEmailDropList = "Email Drop Report"
'Const ReportNamePageTrafficList = "Page Traffic Report"
'Const ReportNamePagePerformance = "Page Performance Report"
'Const ReportNameEmailDropDetails = "Email Drop Details"
'Const ReportNameEmailOpenDetails = "Email Open Details"
'Const ReportNameEmailClickDetails = "Email Click Details"
'Const ReportNameGroupList = "Group Report"
'Const ReportNameGroupMemberList = "Group Member Report"
'Const ReportNameTrapList = "Trap Report"
'
Const DescriptionDailyVisits = "A summary of all visits, grouped by day, that meet the specified Report Criteria. A visit starts at the time of the first page viewing, so visits may be listed that include page viewings that occurred on the next day."
Const DescriptionWeeklyVisits = "A summary of all visits, grouped by week, that meet the specified Report Criteria."
Const DescriptionSearchKeywords = "A List of all keywords searches for visits that meet the specified Report Criteria."
Const DescriptionIPAddressList = "This report lists all the visits created, sorted by the IP address of the visitor. You can use the 'Blocked' column to designate this IP to be blocked from all reports. They will not be blocked from the site however. The Domain name field is available if the site property 'Lookup Visit Domain Names' has been enabled. Creating these domain names can initially be very taxing on your domain name server, and should only enabled if required. Note: Browser and Address Blocking do not apply to this list."
Const DescriptionContentProperties = "A tree navigation of the current content definition structure"
Const DescriptionSurveyList = "A list of all surveys completed during visits that meet the specified Report Criteria."
Const DescriptionSurveyDetails = "Results from a selected survey."
Const DescriptionReferers = "A list of all referring sites for visits that meet the specified Report Criteria."
Const DescriptionSitePath = "A count of all page hits sorted by site path that meet the specified Report Criteria."
Const DescriptionOrdersList = "A list of all orders from visits that meet the specified Report Criteria."
Const DescriptionOrderDetails = "A detailed list of all properties of a specific order."
Const DescriptionMemberDetails = "A detailed list of all properties of the specified member record (identity)."
Const DescriptionPageList = "A list of all page hits that meet the specified Report Criteria. Note the ReadTime is the time between the current page fetch and the next fetch. This fields is not available if the report includes more then one visit, and on the last page of the visit."
Const DescriptionVisitList = "A list of all visits that meet the specified Report Criteria."
Const DescriptionVisitDetails = "A detailed list of all properties of a specific visit."
Const DescriptionVisitorList = "A list of all visitors whose visits meet the specified Report Criteria. A Visitor represents a computer on the Internet that visits the site. The Identity is the person using the computer."
Const DescriptionVisitorDetails = "The specific details about a specified visitor. A Visitor represents a computer on the Internet that visits the site. The Identity is the person using the computer."
Const DescriptionSpiderDocList = "A list of all pages found by the Contensive Content Spider."
Const DescriptionSpiderErrorList = "A list of all errors found by the Contensive Content Spider."
Const DescriptionEDGDocsList = "A list of all link errors collected during the EDG Publish process, and the documents that refered to them."
Const DescriptionDownloadLog = "A list of all Resource Library downloads, with counts for each."
Const DescriptionContentDefinitionMap = "This report maps all content definitions into a parent/child hierarchy."
Const DescriptionEmailDropList = "This report lists all Email drops."
Const DescriptionPageTrafficReport = "This report counts hits to each content page that meet the report criteria. The URL is provided as a sample, but may not be the exact URL used for all hits to that page."
Const DescriptionPagePerformance = "This report lists all unique URLs with their average performance in milliseconds. Page performace is the time measured from when the Contensive Init() is started, to when the Contensive ClosePage() is completed."
Const DescriptionEmailDropDetails = "This report lists all individual email recipients for a specific email drop."
Const DescriptionEmailOpenDetails = "This report lists each occurance of an email being opened for a specific email drop."
Const DescriptionEmailClickDetails = "This report lists each occurance of an email recipient clicking on a link that brings them to the site, for a specific email drop."
Const DescriptionGroupList = "This report lists all groups in the system, with important statistics about the group."
Const DescriptionGroupMemberList = "This report lists all people in a group, with important statistics about the person."
Const DescriptionTrapList = "This report lists all internal Contensive Traps currently stored."
'
Private ContentBranchUsedList As String
'
Private Type ReportRowType
    Caption As String
    Align As String
    Width As String
    End Type
'
Private Type SurveySummaryType
    Name As String
    Count As Long
    End Type
'
' True if wait page is called
'
' handled in main - Private WaitPageOpened As Boolean
'
'
'
Private Const RequestNameIncludeQuestions = "IncludeQuestions"
Private Const RequestNameRequestDownload = "RequestDownload"
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
    Call HandleError("PageExtensionClass", "Execute", Err.number, Err.Source, Err.Description, True, False)
End Function
'
'===========================================================================================================
'
'===========================================================================================================
'
Public Sub Init(mainObject As Object)
    '
    Set main = mainObject
    '
    Exit Sub
ErrorTrap:
    Call HandleError("PageExtensionClass", "Init", Err.number, Err.Source, Err.Description, True, False)
End Sub
'
'===========================================================================================================
'
'===========================================================================================================
'
Public Function GetContent(optionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim Button As String
    Dim copy As String
    Dim stream As New FastStringClass
    'Dim AdminUI As New AdminUIClass
    Dim Obj As Object
    '
    If Not main.IsAdmin() Then
        stream.Add "<p>You must be an Administrator to view this report</p>"
    Else
        Call main.WritePleaseWaitStart
        Call LoadReportGlobals(main)
        '
        RQS = "" _
            & main.RefreshQueryString() _
            & "&ReportTimeout=" & ReportTimeout _
            & "&DateTo=" & main.EncodeRequestVariable(CStr(ReportDateTo)) _
            & "&DateFrom=" & main.EncodeRequestVariable(CStr(ReportDateFrom)) _
            & "&VisitID=" & main.EncodeRequestVariable(CStr(ReportVisitID)) _
            & "&VisitorID=" & main.EncodeRequestVariable(CStr(ReportVisitorID)) _
            & "&MemberID=" & main.EncodeRequestVariable(CStr(ReportMemberID)) _
            & "&ExcludeBrowsers=" & main.EncodeRequestVariable(CStr(ReportExcludeBrowsers)) _
            & "&ExcludeIP=" & main.EncodeRequestVariable(CStr(ReportExcludeIP)) _
            & "&IPAddress=" & main.EncodeRequestVariable(ReportIPAddress) _
            & "&ExcludeNewVisitors=" & main.EncodeRequestVariable(CStr(ReportExcludeNewVisitors)) _
            & "&ExcludeOldVisitors=" & main.EncodeRequestVariable(CStr(ReportExcludeOldVisitors)) _
            & "&VisitReferer=" & main.EncodeRequestVariable(ReportVisitReferer) _
            & "&RecordID=" & ReportRecordID _
            & "&PageNumber=" & ReportPageNumber _
            & "&PageSize=" & ReportPageSize _
            & "&depth=" & ReportLinkDepth
        '
        stream.Add GetReport()
    End If
    GetContent = GetContent & main.GetUploadFormStart(RQS) & stream.Text & main.GetUploadFormEnd
    GetContent = GetContent & main.GetPleaseWaitEnd
    '
    Exit Function
ErrorTrap:
    Call HandleError("OrderListClass", "GetContent", Err.number, Err.Source, Err.Description, True, False)
    End Function
'
'=============================================================================
'   prints a list of orders
'=============================================================================
'
Private Function GetReport() As String
    On Error GoTo ErrorTrap
    '
    Dim BuildVersion As String
    Dim MemberName As String
    Dim memberId As Long
    Dim SQL As String
    Dim cs As Long
    Dim copy As String
    Dim RowPointer As Long
    Dim RowBAse As Long
    Dim stream As New FastStringClass
    '
    Dim Headers() As ReportRowType
    Dim Cells() As String
    Dim PTC As String
    Dim OTC As String
    Dim DataRowCount As Long
    Dim WQS As String
    Dim OrderChargeTotal As Double
    Dim SQLTablePeople As String
    Dim DetailReportAddonID As Long
    Dim DetailReportAddonGuid As String
    '
    Const ColumnCount = 7
    '
    ReDim Headers(ColumnCount)
    ReDim Cells(ReportPageSize, ColumnCount)
    '
    BuildVersion = main.SiteProperty_BuildVersion
    SQLTablePeople = main.GetContentTablename("people")
    DetailReportAddonGuid = "{3B4B7EC2-736E-41A7-8000-210902B79553}"
    'CS = Main.OpenCSContent("Add-ons", "aoguid=" & KmaEncodeSQLText(DetailReportAddonGuid))
    'If Main.IsCSOK(CS) Then
    '    DetailReportAddonID = Main.GetCSInteger(CS, "id")
    'End If
    'Call Main.CloseCS(CS)
    '
    ' Pre and Post table messages
    '
    PTC = Get_Header(ReportNameOrdersList, DescriptionOrdersList)
    'OTC = GetReportFilter(True, True, True, True, True)
    '
    Headers(0).Caption = "<br>Order#"
    Headers(0).Width = "5%"
    Headers(0).Align = "right"
    '
    Headers(1).Caption = "Date<br>Started"
    Headers(1).Width = "20%"
    Headers(1).Align = "right"
    '
    Headers(2).Caption = "Date<br>Finished"
    Headers(2).Width = "20%"
    Headers(2).Align = "right"
    '
    Headers(3).Caption = "Date<br>Shipped"
    Headers(3).Width = "20%"
    Headers(3).Align = "right"
    '
    'Headers(4).Caption = "<br>Visit#"
    'Headers(4).Width = "5%"
    'Headers(4).Align = "right"
    ''
    Headers(4).Caption = "Total<br>Charge"
    Headers(4).Width = "10%"
    Headers(4).Align = "right"
    '
    Headers(5).Caption = "Payment<br>Method"
    Headers(5).Width = "10%"
    Headers(5).Align = "right"
    '
    Headers(6).Caption = "Customer<br>Identity"
    Headers(6).Width = "10%"
    Headers(6).Align = "right"
    '
    '   Get data
    '
    SQL = " SELECT Orders.ID AS OrderNumber, Orders.DateAdded AS DateStarted, Orders.DateCompleted AS DateCompleted, Orders.ShipDate AS DateShipped, Orders.TotalCharge AS OrderCharge, Orders.PayMethod AS PayMethod, " & SQLTablePeople & ".Name AS MemberName, " & SQLTablePeople & ".ID AS MemberID, ccVisits.ID AS VisitID, ccVisitors.ID AS VisitorID, ccVisitors.Name AS VisitorName" _
        & " FROM ((((ccVisitors RIGHT JOIN ccVisits ON ccVisitors.ID = ccVisits.VisitorID) LEFT JOIN ccReportRules AS BrowserRules ON ccVisits.Browser = BrowserRules.Browser) LEFT JOIN ccReportRules AS AddressRules ON ccVisits.REMOTE_ADDR = AddressRules.Address) RIGHT JOIN Orders ON ccVisits.ID = Orders.VisitID) LEFT JOIN " & SQLTablePeople & " ON Orders.MemberID = " & SQLTablePeople & ".ID" _
        & " WHERE ((dateCompleted is not null)" _
            & Get_WhereCriteria(ReportDateFrom, ReportDateTo) _
            & ")" _
        & " Order By Orders.ID;"
    cs = main.OpenCSSQL(main.GetContentDataSource("Visits"), SQL, ReportPageSize, ReportPageNumber)
    RowPointer = 0
    RowBAse = (ReportPageSize * (ReportPageNumber - 1))
    DataRowCount = main.GetCSRowCount(cs)
    Do While main.IsCSOK(cs) And (RowPointer < ReportPageSize)
        '
        ' OrderNumber
        '
        WQS = RQS
        copy = kmaEncodeText(main.GetCSText(cs, "OrderNumber"))
        WQS = ModifyQueryString(WQS, "orderid", copy)
        WQS = ModifyQueryString(WQS, "addonid", "", False)
        WQS = ModifyQueryString(WQS, "addonguid", CStr(DetailReportAddonGuid), True)
        Cells(RowPointer, 0) = GetLink(main, copy, WQS)
        '
        ' Date Started
        '
        Cells(RowPointer, 1) = kmaEncodeText(main.GetCSText(cs, "DateStarted"))
        '
        ' Date Completed
        '
        Cells(RowPointer, 2) = main.GetCSText(cs, "DateCompleted")
        '
        ' Date Shipped
        '
        Cells(RowPointer, 3) = kmaEncodeText(main.GetCSText(cs, "DateShipped"))
        '
        ' VisitID
        '
        'WQS = ModifyQueryString(WQS, RequestNameReportForm, ReportFormVisitDetails)
        'copy = main.GetCSText(cs, "VisitID")
        'WQS = RQS
        'WQS = ModifyQueryString(WQS, "addonid", "", False)
        'WQS = ModifyQueryString(WQS, "af", "12", True)
        'WQS = ModifyQueryString(WQS, RequestNameReportForm, ReportFormVisitDetails)
        'WQS = ModifyQueryString(WQS, "VisitID", copy)
        'Cells(RowPointer, 4) = GetLink(main, copy, WQS, , , , True)
        '
        ' Order charge
        '
        copy = Format(main.GetCSNumber(cs, "OrderCharge"), "0.00")
        Cells(RowPointer, 4) = copy
        OrderChargeTotal = OrderChargeTotal + CDbl(copy)
        '
        ' Pay Method
        '
        Cells(RowPointer, 5) = main.GetCSText(cs, "PayMethod")
        '
        ' Member
        '
        memberId = main.GetCSInteger(cs, "MemberID")
        MemberName = main.GetCSText(cs, "MemberName")
        MemberName = GetMemberName(MemberName, memberId)
        Cells(RowPointer, 6) = MemberName
        'WQS = RQS
        'WQS = ModifyQueryString(WQS, "addonid", "", False)
        'WQS = ModifyQueryString(WQS, "af", "12", True)
        'WQS = ModifyQueryString(WQS, RequestNameReportForm, ReportFormMemberDetails)
        'WQS = ModifyQueryString(WQS, "MemberID", CStr(memberId))
        'Cells(RowPointer, 6) = GetLink(main, MemberName, WQS, , , , True)
        '
        main.NextCSRecord (cs)
        RowPointer = RowPointer + 1
    Loop
    Call main.CloseCS(cs)
    '
    ' Generate Table
    '
    GetReport = Get_Table(main, RowPointer, Headers, Cells, ReportPageSize, ReportPageNumber, PTC, OTC, DataRowCount, GetCloseOrCancelButton & "," & ButtonRefresh)
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    GetReport = HandleReportTrapError("GetReport", PTC, OTC, "ErrorTrap")
    End Function
'
'
'
Private Function GetCloseOrCancelButton()
    If ReportLinkDepth = 0 Then
        GetCloseOrCancelButton = ButtonCancel
    Else
        GetCloseOrCancelButton = ButtonClose
        End If
    End Function
'
'
'
Private Function GetMemberName(MemberName As String, memberId As Long) As String
    GetMemberName = Trim(MemberName)
    If GetMemberName = "" Then
        GetMemberName = "Person #" & memberId
    End If
End Function
'
'
'
Private Function LoadReportGlobals(main As Object)
    Dim copy As String
    Dim SetPageSize As Long
    Dim ReportOptionsEnabled As Boolean
    Dim SetDateFrom As Date
    Dim SetDateTo As Date
    Dim SetVisitID As Long
    Dim SetVisitorID As Long
    Dim SetMemberID As Long
    Dim SetIPAddress As String
    Dim SetExcludeBrowsers As Boolean
    Dim SetExcludeIP As Boolean
    Dim SetExcludeNewVisitors As Boolean
    Dim SetExcludeOldVisitors As Boolean
    '
    If Not ReportGlobalsLoaded Then
        '
        ' Report Depth
        '
        ReportLinkDepth = main.GetStreamInteger("depth")
        Call main.AddRefreshQueryString("depth", ReportLinkDepth)
        '
        ' Report state
        '
        ReportButton = main.GetStreamText("button")
        ReportRecordID = main.GetStreamInteger("RecordID")
        '
        ' Report Timeout (Get and Set)
        '
        ReportTimeout = main.GetStreamInteger("ReportTimeout")
        If ReportButton <> "" Then
            copy = main.GetStreamText("SetReportTimeout")
            If copy <> "" Then
                ReportTimeout = kmaEncodeInteger(copy)
                End If
            End If
        If ReportTimeout = 0 Then
            ReportTimeout = 180
            'ReportTimeout = Main.SQLCommandTimeout
        Else
            main.SQLCommandTimeout = ReportTimeout
            End If
        Call main.SetStreamTimeout(ReportTimeout)
        Call main.AddRefreshQueryString("ReportTimeout", ReportTimeout)
        '
        ' Determine Page Size / Page Number
        '
        ReportPageNumber = main.GetStreamInteger("PageNumber")
        If ReportPageNumber = 0 Then
            ReportPageNumber = 1
            End If
        ReportPageSize = main.GetStreamInteger("PageSize")
        If ReportButton <> "" Then
            copy = main.GetStreamText("SetPageSize")
            If copy <> "" Then
                SetPageSize = main.GetStreamInteger("SetPageSize")
                If SetPageSize <> ReportPageSize Then
                    ReportPageSize = SetPageSize
                    ReportPageNumber = 1
                    End If
                End If
            End If
        If ReportPageSize = 0 Then
            ReportPageSize = 50
            End If
        If main.GetStreamInteger("GotoPage") <> 0 Then
            ReportPageNumber = main.GetStreamInteger("GotoPage")
            End If
        If ReportButton = ButtonNext Then
            ReportPageNumber = ReportPageNumber + 1
            End If
        If ReportButton = ButtonFirst Then
            ReportPageNumber = 1
            End If
        If ReportButton = ButtonPrevious Then
            ReportPageNumber = ReportPageNumber - 1
            If ReportPageNumber < 1 Then
                ReportPageNumber = 1
                End If
            End If
        ReportOptionsEnabled = main.GetStreamBoolean("ReportOptions")
        '
        ' From Date
        '
        copy = main.GetStreamText("DateFrom")
        If IsDate(copy) Then
            ReportDateFrom = CDate(copy)
        ElseIf IsNumeric(copy) Then
            ReportDateFrom = CDate(copy)
        End If
        If ReportOptionsEnabled Then
            SetDateFrom = main.GetStreamDate("SetDateFrom")
            If SetDateFrom <> ReportDateFrom Then
                ReportDateFrom = Int(SetDateFrom)
                ReportPageNumber = 1
            End If
        End If
        If ReportDateFrom = 0 Then
            ReportDateFrom = Int(Now) - 6
        End If
        '
        ' To Date
        '
        copy = main.GetStreamText("DateTo")
        If IsDate(copy) Then
            ReportDateTo = CDate(copy)
        ElseIf IsNumeric(copy) Then
            ReportDateTo = CDate(copy)
            End If
        If ReportOptionsEnabled Then
            SetDateTo = main.GetStreamDate("SetDateTo")
            If (SetDateTo <> ReportDateTo) Then
                ReportDateTo = Int(SetDateTo)
                ReportPageNumber = 1
                End If
            End If
        If ReportDateTo < ReportDateFrom Then
            ReportDateTo = ReportDateFrom + 6
            End If
        ReportDateAfterTo = ReportDateTo + 1
        '
        ' VisitID
        '
        ReportVisitID = main.GetStreamInteger("VisitID")
        If ReportOptionsEnabled Then
            SetVisitID = main.GetStreamInteger("SetVisitID")
            If (SetVisitID <> ReportVisitID) Then
                ReportVisitID = SetVisitID
                ReportPageNumber = 1
                End If
            End If
        '
        ' VisitorID
        '
        ReportVisitorID = main.GetStreamInteger("VisitorID")
        If ReportOptionsEnabled Then
            SetVisitorID = main.GetStreamInteger("SetVisitorID")
            If (SetVisitorID <> ReportVisitorID) Then
                ReportVisitorID = SetVisitorID
                ReportPageNumber = 1
                End If
            End If
        '
        ' MemberID
        '
        ReportMemberID = main.GetStreamInteger("MemberID")
        If ReportOptionsEnabled Then
            SetMemberID = main.GetStreamInteger("SetMemberID")
            If SetMemberID <> ReportMemberID Then
                ReportMemberID = SetMemberID
                ReportPageNumber = 1
            End If
        End If
        '
        ' IPAddress
        '
        ReportIPAddress = main.GetStreamText("IPAddress")
        If ReportOptionsEnabled Then
            SetIPAddress = main.GetStreamText("SetIPAddress")
            If SetIPAddress <> ReportIPAddress Then
                ReportIPAddress = SetIPAddress
                ReportPageNumber = 1
                End If
            End If
        '
        ' Exclude Blocked Browsers
        '
        If InStr(1, main.ServerQueryString, "excludebrowsers", vbTextCompare) = 0 Then
            ReportExcludeBrowsers = kmaEncodeBoolean(main.GetMemberProperty("ReportExcludeBrowsers", True))
        Else
            ReportExcludeBrowsers = main.GetStreamBoolean("ExcludeBrowsers")
            End If
        If ReportOptionsEnabled Then
            SetExcludeBrowsers = main.GetStreamBoolean("SetExcludeBrowsers")
            If SetExcludeBrowsers <> ReportExcludeBrowsers Then
                ReportExcludeBrowsers = SetExcludeBrowsers
                ReportPageNumber = 1
                End If
            End If
        Call main.SetMemberProperty("ReportExcludeBrowsers", ReportExcludeBrowsers)
        '
        ' Exclude IP Addresses
        '
        If InStr(1, main.ServerQueryString, "ExcludeIP", vbTextCompare) = 0 Then
            ReportExcludeIP = kmaEncodeBoolean(main.GetMemberProperty("ReportExcludeIP", True))
        Else
            ReportExcludeIP = main.GetStreamBoolean("ExcludeIP")
            End If
        If ReportOptionsEnabled Then
            SetExcludeIP = main.GetStreamBoolean("SetExcludeIP")
            If SetExcludeIP <> ReportExcludeIP Then
                ReportExcludeIP = SetExcludeIP
                ReportPageNumber = 1
                End If
            End If
        Call main.SetMemberProperty("ReportExcludeIP", ReportExcludeIP)
        '
        ' Exclude new visitors
        '
        ReportExcludeNewVisitors = main.GetStreamBoolean("ExcludeNewVisitors")
        If ReportOptionsEnabled Then
            SetExcludeNewVisitors = main.GetStreamBoolean("SetExcludeNewVisitors")
            If SetExcludeNewVisitors <> ReportExcludeNewVisitors Then
                ReportExcludeNewVisitors = SetExcludeNewVisitors
                ReportPageNumber = 1
                End If
            End If
        '
        ' Exclude old visitors
        '
        ReportExcludeOldVisitors = main.GetStreamBoolean("ExcludeOldVisitors")
        If ReportOptionsEnabled Then
            SetExcludeOldVisitors = main.GetStreamBoolean("SetExcludeOldVisitors")
            If SetExcludeOldVisitors <> ReportExcludeOldVisitors Then
                ReportExcludeOldVisitors = SetExcludeOldVisitors
                ReportPageNumber = 1
                End If
            End If
        '
        ' Visit Referer
        '
        ReportVisitReferer = main.GetStreamText("VisitReferer")
        If ReportOptionsEnabled Then
            ReportVisitReferer = main.GetStreamText("SetVisitReferer")
            End If
        '
        ' Set defaults if no reportid
        '
        If (ReportID = 0) Then
            ReportID = ReportFormRoot
            ReportExcludeBrowsers = True
            ReportExcludeIP = True
            ReportExcludeNewVisitors = False
            ReportExcludeOldVisitors = False
            End If
        If ReportID <> 0 Then
            Call main.AddRefreshQueryString("rid=" & ReportID)
            End If
        If ReportDateTo <> 0 Then
            Call main.AddRefreshQueryString("DateTo=" & ReportDateTo)
            End If
        If ReportDateFrom <> 0 Then
            Call main.AddRefreshQueryString("DateFrom=" & ReportDateFrom)
            End If
        If ReportVisitID <> 0 Then
            Call main.AddRefreshQueryString("VisitID=" & ReportVisitID)
            End If
        If ReportVisitorID <> 0 Then
            Call main.AddRefreshQueryString("VisitorID=" & ReportVisitorID)
            End If
        If ReportMemberID <> 0 Then
            Call main.AddRefreshQueryString("MemberID=" & ReportMemberID)
            End If
        If ReportExcludeBrowsers Then
            Call main.AddRefreshQueryString("ExcludeBrowsers=" & ReportExcludeBrowsers)
            End If
        If ReportExcludeIP Then
            Call main.AddRefreshQueryString("ExcludeIP=" & ReportExcludeIP)
            End If
        If ReportIPAddress <> "" Then
            Call main.AddRefreshQueryString("IPAddress=" & ReportIPAddress)
            End If
        If ReportExcludeNewVisitors Then
            Call main.AddRefreshQueryString("ExcludeNewVisitors=" & ReportExcludeNewVisitors)
            End If
        If ReportExcludeOldVisitors Then
            Call main.AddRefreshQueryString("ExcludeOldVisitors=" & ReportExcludeOldVisitors)
            End If
        If ReportRecordID <> 0 Then
            Call main.AddRefreshQueryString("RecordID=" & ReportRecordID)
            End If
        If ReportVisitReferer <> "" Then
            Call main.AddRefreshQueryString("VisitReferer=" & kmaEncodeRequestVariable(ReportVisitReferer))
            End If
        Call main.AddRefreshQueryString("PageSize=" & ReportPageSize)
        Call main.AddRefreshQueryString("PageNumber=" & ReportPageNumber)
        '
        ReportGlobalsLoaded = True
    End If

End Function
'
'
'
Public Function GetPageSize(main As Object) As Long
    Call LoadReportGlobals(main)
    GetPageSize = ReportPageSize
End Function
'
'
'
Public Function GetPageNumber(main As Object) As Long
    Call LoadReportGlobals(main)
    GetPageNumber = ReportPageNumber
End Function
'
'
'
Public Function GetLink(main As Object, LinkLabel As String, QueryString As String, Optional NewWindow As Variant, Optional ContentName As String, Optional RecordID As Long, Optional GotoPage1 As Boolean) As String
    Dim WorkingQS As String
    '
    WorkingQS = QueryString
    If GotoPage1 Then
        WorkingQS = ModifyQueryString(WorkingQS, "PageNumber", "1", True)
    End If
    If (ContentName <> "") And (RecordID <> 0) Then
        GetLink = main.GetRecordEditLink(ContentName, RecordID)
        If GetLink <> "" Then
            GetLink = GetLink & "<BR>"
        End If
    End If
    If KmaEncodeMissingBoolean(NewWindow, True) Then
        GetLink = GetLink & "<A  href=""" & main.ServerPage & "?" & WorkingQS & """>" & LinkLabel & "</A>"
    Else
        GetLink = GetLink & "<A href=""" & main.ServerPage & "?" & WorkingQS & """>" & LinkLabel & "</A>"
    End If
End Function
'
'=============================================================================
' Get the Report Body Header (Report name and whatever else)
'=============================================================================
'
Private Function Get_Header(ReportName As String, ReportDescription As String) As String
    On Error GoTo ErrorTrap
    '
    Dim WQS As String
    Dim Panel As String
    'Dim AdminUI As New AdminUIClass
    Dim AdminUI As Object
    Dim ReportLink As String
    
    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    
    ReportLink = ReportName
    If ReportLink <> ReportNameRoot Then
        ReportLink = ReportLink
        'ReportLink = "<a href=""?" & Main.RefreshQueryString & "&" & RequestNameReportForm & "=" & ReportFormRoot & """>Reports</a>&nbsp;&gt;&gt;&nbsp;" & ReportLink
    End If
    '
    Get_Header = AdminUI.GetTitleBar(main, ReportLink, ReportDescription)
    'Get_Header = Get_Header & ("<table border=""0"" cellpadding=""" & ReportMargin & """ cellspacing=""0"" width=""100%""><tr><td class=""ccReports"">")
    'Get_Header = Get_Header & (kmaAddSpan("<B>" & ReportName & "</B>", "ccAdminNormal") & "<BR>")
    'If ReportDescription <> "" Then
    '    Get_Header = Get_Header & (kmaAddSpan(ReportDescription, "ccAdminNormal"))
    'End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Get_Header = HandleReportSubTrapError("Get_Header")
    End Function
'
'=============================================================================
' Print the Report Body Header (Report name and whatever else)
'=============================================================================
'
Private Function Get_HeaderCriteria(Optional ReportDescription As String) As String
    End Function
'
'=============================================================================
' Print the Report Body Header (Report name and whatever else)
'=============================================================================
'
Private Function GetReportFilter(AllowDates As Boolean, AllowVisit As Boolean, AllowOptions As Boolean, Optional IncludePageSize As Boolean, Optional IncludeTimeout As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim copy
    '
    If AllowDates Or AllowVisit Or AllowOptions Or IncludePageSize Or IncludeTimeout Then
        'GetReportFilter = GetReportFilter & "<tr><td>"
        GetReportFilter = GetReportFilter & "<h2>Report Criteria</h2>"
        GetReportFilter = GetReportFilter & "<p>Use this section to request specific report criteria.</p>"
        GetReportFilter = GetReportFilter & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"">"
        If IncludeTimeout Then
            GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetReportTimeout", ReportTimeout, 1, 10), "Maximum time for database activity (seconds)")
        End If
        If IncludePageSize Then
            GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetPageSize", ReportPageSize, 1, 10), "Rows per page")
        End If
        If AllowDates Or AllowOptions Or AllowVisit Then
            GetReportFilter = GetReportFilter & main.GetFormInputHidden("ReportOptions", "1")
            If AllowDates Then
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetDateFrom", ReportDateFrom, 1, 10), "Exclude Visits starting before this date")
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetDateTo", ReportDateTo, 1, 10), "Exclude Visits starting after this date")
            End If
            If AllowVisit Then
                If ReportVisitID = 0 Then
                    copy = ""
                Else
                    copy = CStr(ReportVisitID)
                End If
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetVisitID", copy, 1, 10), "Only include this Visit number")
            End If
            If AllowOptions Then
                If ReportVisitorID = 0 Then
                    copy = ""
                Else
                    copy = CStr(ReportVisitorID)
                End If
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetVisitorID", copy, 1, 10), "Only include visits from this Visitor number")
                If ReportMemberID = 0 Then
                    copy = ""
                Else
                    copy = CStr(ReportMemberID)
                End If
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetMemberID", copy, 1, 10), "Only include visits from this Member number")
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetIPAddress", ReportIPAddress, 1, 10), "Only include visits from this IP Address")
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputText("SetVisitReferer", ReportVisitReferer, 1, 10), "Only include visits from this Referer")
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputCheckBox("SetExcludeNewVisitors", ReportExcludeNewVisitors), "Exclude new visitors")
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputCheckBox("SetExcludeOldVisitors", ReportExcludeOldVisitors), "Exclude repeat visitors")
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputCheckBox("SetExcludeBrowsers", ReportExcludeBrowsers), "Exclude blocked browsers")
                GetReportFilter = GetReportFilter & GetReportFilterRow(main.GetFormInputCheckBox("SetExcludeIP", ReportExcludeIP), "Exclude blocked IP addresses")
            End If
        End If
        GetReportFilter = GetReportFilter & ("</table>")
'        GetReportFilter = GetReportFilter & ("<p></p>")
'        GetReportFilter = GetReportFilter & ("</td></tr>")
    End If
    '
    ' Close the table opened
    '
'    GetReportFilter = GetReportFilter & ("</table>")
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    GetReportFilter = HandleReportSubTrapError("GetReportFilter")
    End Function
'
'=============================================================================
' Create the part of the sql where clause that is modified by the user
'=============================================================================
'
Private Function Get_WhereCriteria(ReportDateFrom As Date, ReportDateTo As Date, Optional VisitTableOnly As Boolean) As String
    On Error GoTo ErrorTrap
    '
    Dim copy As String
    '
    If Not VisitTableOnly Then
        If ReportExcludeBrowsers Then
            Get_WhereCriteria = Get_WhereCriteria & "AND(BrowserRules.ID Is Null)"
            Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.Browser<>'')"
            End If
        If ReportExcludeIP Then
            Get_WhereCriteria = Get_WhereCriteria & "AND(AddressRules.ID Is Null)"
            End If
        End If
    If (ReportDateFrom <> 0) Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.StartTime>=" & KmaEncodeSQLDate(ReportDateFrom) & ")"
        End If
    If (ReportDateTo <> 0) Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.StartTime<" & KmaEncodeSQLDate(ReportDateAfterTo) & ")"
        End If
    If (ReportMemberID <> 0) Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.MemberID=" & KmaEncodeSQLNumber(ReportMemberID) & ")"
        End If
    If (ReportVisitID <> 0) Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.ID=" & KmaEncodeSQLNumber(ReportVisitID) & ")"
        End If
    If (ReportVisitorID <> 0) Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.VisitorID=" & KmaEncodeSQLNumber(ReportVisitorID) & ")"
        End If
    If (ReportIPAddress <> "") Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.Remote_ADDR=" & KmaEncodeSQLText(ReportIPAddress) & ")"
        End If
    If ReportExcludeNewVisitors Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(Not(ccVisits.VisitorNew<>0))"
        End If
    If ReportExcludeOldVisitors Then
        Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.VisitorNew<>0)"
        End If
    If ReportVisitReferer <> "" Then
        If ReportVisitReferer = RequestValueNull Then
            Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.HTTP_REFERER is null)"
        Else
            copy = KmaEncodeSQLText(ReportVisitReferer)
            copy = Mid(copy, 2, Len(copy) - 2)
            Get_WhereCriteria = Get_WhereCriteria & "AND(ccVisits.HTTP_REFERER like '%" & copy & "%')"
            End If
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleReportSubTrapError("Get_WhereCriteria")
    End Function

'
'=============================================================================
'
'=============================================================================
'
Private Function Get_Table(main As Object, RowCount As Long, Headers() As ReportRowType, Cells() As String, PageSize As Long, PageNumber As Long, PTC As String, OTC As String, DataRowCount As Long, ButtonList As String) As String
    On Error GoTo ErrorTrap
    '
    Dim ButtonBar As String
    Dim ColumnCount As Long
    Dim ColumnPointer As Long
    Dim ColCaptions() As String
    Dim ColAligns() As String
    Dim ColWidths() As String
    Dim PanelSpacer As String
    Dim OTCPadded As String
    'Dim AdminUI As New AdminUIClass
    Dim AdminUI As Object
    
    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    
    ColumnCount = UBound(Cells, 2)
    ReDim ColCaptions(ColumnCount)
    ReDim ColAligns(ColumnCount)
    ReDim ColWidths(ColumnCount)
    
    If ColumnCount > 0 Then
        For ColumnPointer = 0 To ColumnCount - 1
            ColCaptions(ColumnPointer) = Headers(ColumnPointer).Caption
            ColAligns(ColumnPointer) = Headers(ColumnPointer).Align
            ColWidths(ColumnPointer) = Headers(ColumnPointer).Width
        Next
    End If
    PanelSpacer = ""
    OTCPadded = "<div style=""padding:10px"">" & OTC & "</div>"
    'OTCPadded = "<div style=""padding:10px"">" & OTC & "</div>"
    ButtonBar = Get_Buttons(ButtonList, RequestNameButton, ButtonFirst & "," & ButtonPrevious & "," & ButtonNext)
    Get_Table = AdminUI.GetReport(main, RowCount, ColCaptions, ColAligns, ColWidths, Cells, PageSize, PageNumber, ButtonBar & PTC, OTCPadded & ButtonBar, DataRowCount, "")
    'Get_Table = Main.GetReport(RowCount, ColCaptions, ColAligns, ColWidths, Cells, PageSize, PageNumber, ButtonBar & PTC, OTCPadded & ButtonBar, DataRowCount, "")
    
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Get_Table = HandleReportSubTrapError("Get_Table")
    End Function
'
'========================================================================
' Draw a Graph
'
'   Xlabel
'   YLabel
'   XCount          The number of plots to draw
'   DataSet( XCount )    The number of points ( x dimension )
'========================================================================
'
Private Function Get_Chart(ChartData As ChartDataType) As String
    On Error GoTo ErrorTrap
    '
    Dim RowPointer
    Dim ColumnPointer
    Dim MethodName As String
    Dim stream As New FastStringClass
    '
    MethodName = "Get_Chart"
    '
    stream.Add ("<table border=""0"" cellspacing=""1"" cellpadding=""3"" width=""100%"">")
    stream.Add ("<tr><td class=""ccPanel"">" & SpanClassAdminSmall & ChartData.Title & "</SPAN></td></tr>")
    stream.Add ("<tr><td class=""ccPanelRowEven"">")
    '
    stream.Add (vbCrLf & "<applet codeBase=""/cclib/tchart/"" code=""tchart2.class"" width=""100%"" height=""400"" >")
    'Stream.Add (vbCrLf & "<applet codeBase=""/cclib/tchart/"" archive=""tchart2.jar"" code=""tchart2.class"" width=""100%"" height=""400"">")
    stream.Add (vbCrLf & "<Param name=""file"" value=""z"">")
    stream.Add (vbCrLf & "<Param name=""font"" value=""arial"">")
    stream.Add (vbCrLf & "<Param name=""fontsize"" value=""10"">")
    stream.Add (vbCrLf & "<Param name=""bgcolor"" value=""15134966"">")
    '
    stream.Add (vbCrLf & "<Param name=""z0"" value=""Bar ; " & ChartData.ColumnCount & " ; za ; 0 ; 1 "">")
    stream.Add (vbCrLf & "<Param name=""z1"" value=""Bar Labeled ; " & ChartData.ColumnCount & " ; za ; 0 ; 0 "">")
    stream.Add (vbCrLf & "<Param name=""z2"" value=""Line ; " & ChartData.ColumnCount & " ; za ; 1 ; 1 "">")
    stream.Add (vbCrLf & "<Param name=""z3"" value=""Line Labeled ; " & ChartData.ColumnCount & " ; za ; 1 ; 0 "">")
    stream.Add (vbCrLf & "<Param name=""z4"" value=""Pie ; " & ChartData.ColumnCount & " ; za ; 2 ; 1 "">")
    stream.Add (vbCrLf & "<Param name=""z5"" value=""Pie Labeled ; " & ChartData.ColumnCount & " ; za ; 2 ; 0 "">")
    stream.Add (vbCrLf & "<Param name=""z6"" value=""Bar Summing ; " & ChartData.ColumnCount & " ; za ; 3 ; 1 "">")
    stream.Add (vbCrLf & "<Param name=""z7"" value=""Bar Summing Labeled ; " & ChartData.ColumnCount & " ; za ; 3 ; 0 "">")
    stream.Add (vbCrLf & "<Param name=""z8"" value=""Table ; " & ChartData.ColumnCount & " ; za ; 4 ; 13"">")
    stream.Add (vbCrLf & "<Param name=""z9"" value=""Area ; " & ChartData.ColumnCount & " ; za ; 5 ; 13"">")
    stream.Add (vbCrLf & "<Param name=""z10"" value=""Area Summing ; " & ChartData.ColumnCount & " ; za ; 6 ; 13"">")
    '
    stream.Add (vbCrLf & "<Param name=""za0"" value=""title ; " & ChartData.Title & """>")
    stream.Add (vbCrLf & "<Param name=""za1"" value=""label ; " & ChartData.YLabel & " ; " & ChartData.XLabel & " "">")
    For ColumnPointer = 0 To ChartData.ColumnCount - 1
        stream.Add (vbCrLf & "<Param name=""za" & (2 + ColumnPointer) & """ value=""col ; " & ChartData.Column(ColumnPointer).Name & " ; " & ColumnPointer & "22500 "">")
        Next
    'Stream.Add (vbCrLf & "<Param name=""za2"" value=""col ; " & ChartData.Column(0).Name & " ; 22500 "">")
    'Stream.Add (vbCrLf & "<Param name=""za3"" value=""col ; " & ChartData.Column(1).Name & " ; 122500 "">")
    'Stream.Add (vbCrLf & "<Param name=""za4"" value=""col ; " & ChartData.Column(2).Name & " ; 222500 "">")
    For RowPointer = 0 To ChartData.RowCount - 1
        stream.Add (vbCrLf & "<Param name=""za" & 2 + ChartData.ColumnCount + RowPointer & """ value=""" & ChartData.RowLabel(RowPointer))
        For ColumnPointer = 0 To ChartData.ColumnCount - 1
            stream.Add (" ; " & ChartData.Column(ColumnPointer).row(RowPointer))
            Next
        stream.Add (""">")
        Next
    stream.Add (vbCrLf & "</applet>")
    stream.Add ("</td></tr></table>")
    Get_Chart = stream.Text
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Get_Chart = HandleReportSubTrapError("Get_Chart")
    End Function
'
'==================================================================================================
'   Called in reports - returns an error report
'==================================================================================================
'
Private Function HandleReportTrapError(MethodName As String, PTC As String, OTC As String, ignore0 As String) As String
    Dim ErrNumber As Long
    Dim ButtonBar As String
    '
    ErrNumber = Err.number
    '
    Call HandleError("Reports", MethodName, Err.number, Err.Source, Err.Description, True, True)
    '
    ButtonBar = Get_Buttons(GetCloseOrCancelButton & "," & ButtonCreateReport, RequestNameButton)
    HandleReportTrapError = ButtonBar & main.GetPanel("") & PTC _
        & Get_Error("There was an error creating your report. The server may be busy, or the data may be too large for the time provided. If possible, try increasing the time 'Maximum time for database activity (seconds)' value, or run this report during a time when the server is less busy.") _
        & OTC & ButtonBar
    End Function
'
'==================================================================================================
'   Called in report subroutines - returns a "[error]" string
'==================================================================================================
'
Private Function HandleReportSubTrapError(MethodName As String) As String
    Dim ErrNumber As Long
    '
    ErrNumber = Err.number
    '
    Call HandleError("Reports", MethodName, Err.number, Err.Source, Err.Description, True, True)
    HandleReportSubTrapError = "[error]"
    End Function
'
'
'
Private Sub HandleReportInternalError(MethodName As String, ErrDescription As String)
    Call Err.Raise(KmaErrorInternal, App.EXEName, ErrDescription)
    End Sub
'
'==========================================================================================
'   Detail Form, Print a Row
'==========================================================================================
'
Private Function Get_DetailsRow(CSPointer As Long, FieldName As String) As String
    '
    Get_DetailsRow = Get_DetailsRow & Get_RowStart
    Get_DetailsRow = Get_DetailsRow & Get_Cell(main.GetCSFieldCaption(CSPointer, FieldName), "right")
    Get_DetailsRow = Get_DetailsRow & Get_Cell(main.GetCSEncodedField(CSPointer, FieldName), "left")
    Get_DetailsRow = Get_DetailsRow & Get_RowEnd
    '
    End Function
'
'==========================================================================================
'
'==========================================================================================
'
Private Function Get_Cell(copy As String, Optional Align As String, Optional Columns As Long) As String
    Dim iAlign As String
    Dim Style As String
    '
    iAlign = KmaEncodeMissingText(Align, "left")
    '
    If (FormResultsRowCount Mod 2) > 0 Then
        Style = "ccPanelRowEven"
    Else
        Style = "ccPanelRowOdd"
        End If
    '
    Get_Cell = Get_Cell & "<td valign=""top"" align=""" & iAlign & """ class=""" & Style & """"
    If Columns <> 0 Then
        Get_Cell = Get_Cell & " colspan=""" & Columns & """"
        End If
    Get_Cell = Get_Cell & ">" & SpanClassAdminSmall & copy & "</SPAN></td>"
    End Function
'
'==========================================================================================
'
'==========================================================================================
'
Private Function Get_RowStart()
    Get_RowStart = kmaStartTableRow
    End Function
'
'==========================================================================================
'
'==========================================================================================
'
Private Function Get_RowEnd()
    Get_RowEnd = kmaEndTableRow
    FormResultsRowCount = FormResultsRowCount + 1
    End Function
'
'==========================================================================================
'
'==========================================================================================
'
Private Function Get_Buttons(LeftButtonList As String, ButtonName As String, Optional RightButtonList As String)
    
    Dim AdminUI As Object
    
    Set AdminUI = CreateObject("ccWeb3.AdminUIClass")
    
    Get_Buttons = AdminUI.GetButtonBar(AdminUI.GetButtonsFromList(main, LeftButtonList, True, True, ButtonName), AdminUI.GetButtonsFromList(main, RightButtonList, True, True, ButtonName))

    End Function
'
'
'
Private Function Get_CellHeader(Title As String, Optional Width As String, Optional Align As String) As String
    On Error GoTo ErrorTrap
    '
    Dim CellTitle As String
    Dim CellWidth As String
    Dim ClassStyle As String
    '
    ClassStyle = "ccPanel"
    If Title <> "" Then
        CellTitle = "<P class=""" & ClassStyle & """"
        If Align <> "" Then
            CellTitle = CellTitle & " align=""" & Align & """"
            End If
        If (Width = "") Or (InStr(1, Width, "%") <> 0) Then
            CellWidth = "100%"
        Else
            CellWidth = Width
            End If
        CellTitle = CellTitle & "><nobr>" & Title & "</nobr><BR><img src=""/ccLib/images/black.gif"" height=1 width=""" & CellWidth & """></p>"
        'CellTitle = GetPanel(CellTitle)
        End If
    If Width = "" Then
        Get_CellHeader = "<td class=""" & ClassStyle & """ valign=""bottom"">" & CellTitle & "</td>"
    Else
        Get_CellHeader = "<td class=""" & ClassStyle & """ valign=""bottom"" width=""" & Width & """>" & CellTitle & "</td>"
        End If
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Get_CellHeader = HandleReportSubTrapError("GetReport_OrderDetails")
    End Function
'
'
'
Private Function GetReport_TableStart() As String
    GetReport_TableStart = GetReport_TableStart & "<div class=""ccPanel"">"
    GetReport_TableStart = GetReport_TableStart & kmaStartTable(3, 1, 0) & kmaStartTableRow
    FormResultsRowCount = 0
    End Function
'
'
'
Private Function GetReport_TableEnd() As String
    GetReport_TableEnd = "</tr></table></div>"
    End Function
'
'
'
Private Function GetReportFilterRow(FormInput As String, Caption As String) As String
    GetReportFilterRow = GetReportFilterRow & "<tr><td width=""200"" align=""right"">" & FormInput & "</td>"
    GetReportFilterRow = GetReportFilterRow & "<td width=""100%"" align=""left"">&nbsp;" & Caption & "</td></tr>"
    End Function

Private Function Get_CellTotal(copy As String) As String
    On Error GoTo ErrorTrap
    '
    Get_CellTotal = Get_CellTotal & "<td align=""right"" class=""ccPanelRowOdd"">" & SpanClassAdminSmall
    Get_CellTotal = Get_CellTotal & "<img src=""/ccLib/images/black.gif"" width=""100%"" height=""1""><br>"
    Get_CellTotal = Get_CellTotal & copy & "</SPAN></td>"
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Get_CellTotal = HandleReportSubTrapError("Get_CellTotal")
    End Function
'
'
'
Private Function Get_Link(LinkLabel As String, QueryString As String, Optional NewWindow As Variant, Optional ContentName As String, Optional RecordID As Long, Optional GotoPage1 As Boolean) As String
    Dim WorkingQS As String
    '
    WorkingQS = QueryString
    If GotoPage1 Then
        WorkingQS = ModifyQueryString(WorkingQS, "PageNumber", "1", True)
        End If
    If (ContentName <> "") And (RecordID <> 0) Then
        Get_Link = main.GetRecordEditLink(ContentName, RecordID)
        If Get_Link <> "" Then
            Get_Link = Get_Link & "<BR>"
            End If
        End If
    If KmaEncodeMissingBoolean(NewWindow, True) Then
        Get_Link = Get_Link & "<A  href=""" & main.ServerPage & "?" & WorkingQS & """>" & LinkLabel & "</A>"
    Else
        Get_Link = Get_Link & "<A href=""" & main.ServerPage & "?" & WorkingQS & """>" & LinkLabel & "</A>"
        End If
    End Function

'
'=============================================================================
'   Page Traffic Report
'=============================================================================
'
Private Function Get_Error(Message As String) As String
    On Error GoTo ErrorTrap
    '
    Get_Error = kmaStartTable(20, 1, 0, "ccError") _
        & kmaGetTableCell(Message, "100%", 1, False, "Left") _
        & kmaEndTable
    '
    Exit Function
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Get_Error = HandleReportSubTrapError("Get_Error")
    End Function


