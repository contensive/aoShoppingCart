VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "imgResizeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private CSV As Object
Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, optionString As String, FilterInput As String) As String
    On Error GoTo ErrorTrap
    
    Set CSV = CsvObject
    Set Main = MainObject
    
    Dim stream As String
    Dim criteria As String
    
    Dim thumbWidth As String
    Dim imgWidth As String
    Dim ItemID As Integer
    
    Dim image As String
    Dim thumb As String
    
    Dim cs As Long
    Dim newFile As String
    Dim newWidth As String
    Dim newHeight As String
    Dim altSizeList As String
    Dim img As New SfImageResize.ImageResize
    Dim retString As String
    Dim doc As New DOMDocument60
    
    thumbWidth = Main.GetSiteProperty("Catalog Thumbnail Width", "100")
    imgWidth = Main.GetSiteProperty("Catalog Image Width", "300")
    ItemID = kmaEncodeInteger(Main.GetAggrOption("itemID", optionString))

    criteria = "((ImageFilename is Not Null) OR (ThumbNailFilename is Not Null))"
    criteria = criteria & " and ((ImageFilename not like '%.bmp%') or (ThumbNailFilename not like '%.bmp%'))"
    If ItemID <> 0 Then
        criteria = criteria & " and (ItemID=" & ItemID & ")"
    End If

    cs = Main.OpenCSContent("Items", criteria, , , , , "ImageFilename,ThumbNailFilename,AltSizeList")
    Do While Main.CSOK(cs)

        altSizeList = Main.GetCS(cs, "AltSizeList")

        optionString = "Image Source=" & encodeNvaArgument(Main.PhysicalFilePath & Main.GetCSText(cs, "ImageFileName"))
        optionString = optionString & "&Image Width=" & encodeNvaArgument(imgWidth)
        'retString = Main.ExecuteAddon(0, "{59A5FCF4-AC87-4137-8966-B12B66666B56}", OptionString, ContextAdmin, "", 0, "", "", -1)

        If retString <> "" Then
            Call doc.loadXML(retString)
            If doc.hasChildNodes Then
                newFile = doc.getElementsByTagName("newImage").Item(0).Text
                newWidth = doc.getElementsByTagName("imageWidth").Item(0).Text
                newHeight = doc.getElementsByTagName("imageHeight").Item(0).Text

                If InStr(1, altSizeList, newWidth & "x" & newHeight) = 0 Then
                    If altSizeList <> "" Then
                        altSizeList = altSizeList & "," & newWidth & "x" & newHeight
                    Else
                        altSizeList = newWidth & "x" & newHeight
                    End If
                    Call Main.SetCS(cs, "altSizeList", altSizeList)
                End If

            End If
        End If

        optionString = "Image Source=" & encodeNvaArgument(Main.PhysicalFilePath & Main.GetCSText(cs, "ImageFileName"))
        optionString = optionString & "&Image Width=" & encodeNvaArgument(thumbWidth)
        retString = Main.ExecuteAddon(0, "{59A5FCF4-AC87-4137-8966-B12B66666B56}", optionString, ContextAdmin, "", 0, "", "", -1)

        If retString <> "" Then
            Call doc.loadXML(retString)
            If doc.hasChildNodes Then
                newFile = doc.getElementsByTagName("newImage").Item(0).Text
                newWidth = doc.getElementsByTagName("imageWidth").Item(0).Text
                newHeight = doc.getElementsByTagName("imageHeight").Item(0).Text

                If InStr(1, altSizeList, newWidth & "x" & newHeight) = 0 Then
                    If altSizeList <> "" Then
                        altSizeList = altSizeList & "," & newWidth & "x" & newHeight
                    Else
                        altSizeList = newWidth & "x" & newHeight
                    End If
                    Call Main.SetCS(cs, "altSizeList", altSizeList)
                End If

            End If
        End If

        Call Main.NextCSRecord(cs)
    Loop
    Call Main.CloseCS(cs)
    
    Execute = stream
    
    Exit Function
ErrorTrap:
    Call HandleError("imgResizeClass", "Execute", Err.Number, Err.Source, Err.Description, True, False)
    End Function



