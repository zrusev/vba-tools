Attribute VB_Name = "Module_OCR"
'Option Explicit
  
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SRCCOPY = &HCC0020
 
Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
 
'API

Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "user32.dll" () As Long
Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function CloseClipboard Lib "user32.dll" () As Long
Declare Function CountClipboardFormats Lib "user32" () As Long
Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long

Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

Declare Function CreateIC Lib "GDI32" Alias "CreateICA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
 
Public Function OCR(x_coordinate As Long, y_coordinate As Long, height As Long, widht As Long)
Dim wbDest As Workbook, wsDest As Worksheet
Dim FromType As String, PicHigh As Single
Dim PicWide As Single, PicWideInch As Single
Dim PicHighInch As Single, DPI As Long
Dim PixelsWide As Integer, PixelsHigh As Integer

Dim cht As Chart
Dim obj As Object
Dim shp As Shape
Dim strOCR As String
 
Call TOGGLEEVENTS(False)
Call CaptureScreen(x_coordinate, y_coordinate, height, widht)
 
If CountClipboardFormats = 0 Then
    strOCR = "Error: Clipboard is currently empty."
    GoTo EndOfSub
End If

'Determine the format of the current clipboard contents.  There may be multiple
'formats available but the Paste methods below will always (?) give priority
'to enhanced metafile (picture) if available so look for that first.

If IsClipboardFormatAvailable(14) <> 0 Then
    FromType = "pic"
ElseIf IsClipboardFormatAvailable(2) <> 0 Then
    FromType = "bmp"
Else
    strOCR = "Error: Clipboard does not contain a picture or bitmap to paste."
    GoTo EndOfSub
End If
  
Set wbDest = Application.Workbooks.Add
Set wsDest = wbDest.Worksheets.Add
wbDest.Activate
wsDest.Activate
wsDest.Range("B3").Activate
 
'Paste a picture/bitmap from the clipboard (if possible) and select it.
'The clipboard may contain both text and picture/bitmap format items.  If so,
'using just ActiveSheet.Paste will paste the text.  Using Pictures.Paste will
'paste a picture if a picture/bitmap format is available, and the Typename
'will return "Picture" (or perhaps "OLEObject").  If *only* text is available,
'Pictures.Paste will create a new TextBox (not a picture) on the sheet and
'the Typename will return "TextBox".  (This condition now checked above.)

'On Error GoTo Other_Error 'just in case

wsDest.Pictures.Paste.Select
 
'If the pasted item is an "OLEObject" then must convert to a bitmap
'to get the correct size, including the added border and matting.
'Do this via a CopyPicture-Bitmap and then a second Pictures.Paste.

If TypeName(Selection) = "OLEObject" Then
    With Selection
        .CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        .Delete
        ActiveSheet.Pictures.Paste.Select
         'Modify the FromType (used below in the suggested file name)
         'to signal that the original clipboard image is not being used.
        FromType = "ole object"
    End With
End If
 
'Make sure that what was pasted and selected is as expected.
'Note this is the Excel TypeName, not the clipboard format.
If TypeName(Selection) = "Picture" Then
    With Selection
        PicWide = .Width
        PicHigh = .height
        .Delete
    End With
Else
     'Can get to here if a chart is selected and "Copy"ed instead of "Copy Picture"ed.
     'Otherwise, ???.
    If TypeName(Selection) = "ChartObject" Then
        strOCR = "Error: Use Shift > Edit > Copy Picture on charts, not just Copy."
    Else
        strOCR = "Error: Excel pasted a '" & TypeName(Selection) & "' instead of a Picture."
    End If
     'Clean up and quit.
    ActiveWorkbook.Close SaveChanges:=False
    GoTo EndOfSub
End If
 
'Add an empty embedded chart, sized as above, and activate it.
'Positioned at cell B3 just for convenient debugging and final viewing.
'Tip from Jon Peltier:  Just add the embedded chart directly, don't use the
'macro recorder method of adding a new separate chart sheet and then relocating
'the chart back to a worksheet.

With wsDest
    .ChartObjects.Add(.Range("B3").Left, .Range("B3").Top, PicWide, PicHigh).Activate
End With
 
'Paste the [resized] bitmap into the ChartArea, which creates ActiveChart.Shapes(1).

'On Error Resume Next
ActiveChart.Pictures.Paste.Select

'On Error GoTo 0
If TypeName(Selection) = "Picture" Then
    With ActiveChart
         'Adjust the position of the pasted picture, aka ActiveChart.Shapes(1).
         'Adjustment is slightly greater than the .ChartArea.Left/Top offset, why ???
         ''''         .Shapes(1).IncrementLeft -1
         ''''         .Shapes(1).IncrementTop -4
         'Remove chart border.  This must be done *after* all positioning and sizing.
         '         .ChartArea.Border.LineStyle = 0
    End With
     
    'Show pixel size info above the picture-in-chart-soon-to-be-GIF/JPEG/PNG.
    PicWideInch = PicWide / 72 'points to inches ("logical", not necessarily physical)
    PicHighInch = PicHigh / 72
    DPI = PixelsPerInch() 'typically 96 or 120 dpi for displays
    PixelsWide = PicWideInch * DPI
    PixelsHigh = PicHighInch * DPI
Else
    'Something other than a Picture was pasted into the chart.
    'This is very unlikely.
    strOCR = "Error: Clipboard corrupted, possibly by another task."
End If
          
Set cht = wbDest.ActiveChart

'cht.Select
Set shp = cht.Shapes(1)

'Exporting the Chart to Image file.

cht.Export ThisWorkbook.Path & "\temp.bmp"

'Calling the Second method to read text
strOCR = GetText(ThisWorkbook.Path & "\temp.bmp")

'Deleting the image file
Kill ThisWorkbook.Path & "\temp.bmp"
  
'Deleting the sheet as well
Application.DisplayAlerts = False
wbDest.Close False
Application.DisplayAlerts = True

OCR = strOCR

Exit Function

EndOfSub:
    Call TOGGLEEVENTS(True)
    OCR = strOCR
Exit Function

Other_Error:
    strOCR = "Error: Other error (on error resume next)"
    GoTo EndOfSub
         
End Function
 
Private Function CaptureScreen(Left As Long, Top As Long, Width As Long, height As Long)
Dim srcDC As Long, trgDC As Long, BMPHandle As Long, dm As DEVMODE

srcDC = CreateDC("DISPLAY", "", "", dm)
trgDC = CreateCompatibleDC(srcDC)
    
BMPHandle = CreateCompatibleBitmap(srcDC, Width, height)
SelectObject trgDC, BMPHandle
BitBlt trgDC, 0, 0, Width, height, srcDC, Left, Top, SRCCOPY

OpenClipboard 0&
EmptyClipboard
SetClipboardData 2, BMPHandle
CloseClipboard

DeleteDC trgDC
ReleaseDC BMPHandle, srcDC

End Function
 
Private Function TOGGLEEVENTS(blnState As Boolean)

With Application
    .DisplayAlerts = blnState
    .EnableEvents = blnState
    .ScreenUpdating = blnState
    If blnState Then .CutCopyMode = False
    If blnState Then .StatusBar = False
End With

End Function
 
Private Function PixelsPerInch() As Long
Dim hdc As Long

hdc = CreateIC("DISPLAY", vbNullString, vbNullString, 0)
PixelsPerInch = GetDeviceCaps(hdc, 88)
DeleteDC (hdc)

End Function
  
Private Function GetFileBytes(strPath As String) As Byte()

With CreateObject("ADODB.Stream")
    .Open
    .Type = 1  ' adTypeBinary
    .LoadFromFile strPath
    GetFileBytes = .Read
    .Close
End With

End Function

Private Function EncodeBase64(arrData() As Byte) As String
Dim objXML As MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement

Set objXML = New MSXML2.DOMDocument60
Set objNode = objXML.createElement("b64")

objNode.DataType = "bin.base64"
objNode.nodeTypedValue = arrData
EncodeBase64 = objNode.Text

Set objNode = Nothing
Set objXML = Nothing

End Function

Private Function GetText(strfile As String)
    
Dim bytFile() As Byte
Dim base64String, imageXmlStr
Dim nCounter As Integer
Dim strOCR As String

Dim oneNote As OneNote14.Application
Set oneNote = New OneNote14.Application
 
'Get all of the Notebook nodes.
Dim nodes As MSXML2.IXMLDOMNodeList
Set nodes = GetFirstOneNoteNotebookNodes(oneNote)

Set bitmap = LoadPicture(strfile)
bytFile = GetFileBytes(strfile)
base64String = EncodeBase64(bytFile)

hh = Round(bitmap.Width / 30)
ww = Round(bitmap.height / 30)

If Not nodes Is Nothing Then
    ' Get the first OneNote Notebook in the XML document.
    Dim node As MSXML2.IXMLDOMNode
    Set node = nodes(0)
    Dim noteBookName As String
    noteBookName = node.Attributes.getNamedItem("name").Text
     
    ' Get the ID for the Notebook so the code can retrieve
    ' the list of sections.
    Dim notebookID As String
    notebookID = node.Attributes.getNamedItem("ID").Text
     
    ' Load the XML for the Sections for the Notebook requested.
    Dim sectionsXml As String
    oneNote.GetHierarchy notebookID, hsSections, sectionsXml, xs2010
     
    Dim secDoc As MSXML2.DOMDocument60
    Set secDoc = New MSXML2.DOMDocument60
 
    If secDoc.LoadXML(sectionsXml) Then
        ' select the Section nodes
        Dim secNodes As MSXML2.IXMLDOMNodeList
        Set secNodes = secDoc.DocumentElement.getElementsByTagName("one:Section")
         
        If Not secNodes Is Nothing Then
            ' Get the first section.
            Dim secNode As MSXML2.IXMLDOMNode
            Set secNode = secNodes(0)
             
            Dim sectionName As String
            sectionName = secNode.Attributes.getNamedItem("name").Text
            Dim sectionID As String
            sectionID = secNode.Attributes.getNamedItem("ID").Text
             
            ' Create a new blank Page in the first Section
            ' using the default format.
            Dim newPageID As String
            oneNote.CreateNewPage sectionID, newPageID, npsDefault
             
            ' Get the contents of the page.
            Dim outXML As String
            
            oneNote.GetPageContent newPageID, outXML, piAll, xs2010
             
             
            Dim doc As MSXML2.DOMDocument60
            Set doc = New MSXML2.DOMDocument60
            ' Load Page's XML into a MSXML2.DOMDocument60 object.
            If doc.LoadXML(outXML) Then
                ' Get Page Node.
                Dim pageNode As MSXML2.IXMLDOMNode
                Set pageNode = doc.getElementsByTagName("one:Page")(0)

                Dim newElement As MSXML2.IXMLDOMElement
                Dim newNode As MSXML2.IXMLDOMNode
                 
                ' Create Outline node.
                Set newElement = doc.createElement("one:Outline")
                newElement.setAttribute "lang", "en-US"
                Set newNode = pageNode.appendChild(newElement)
                ' Create OEChildren.
                Set newElement = doc.createElement("one:OEChildren")
                Set newNode = newNode.appendChild(newElement)
                ' Create OE.
                Set newElement = doc.createElement("one:OE")
                newElement.setAttribute "lang", "en-US"
                Set newNode = newNode.appendChild(newElement)
                
                ' Create Image.
                Set newElement = doc.createElement("one:Image")
                'newElement.setAttribute "format", "bmp"
                Set newNode = newNode.appendChild(newElement)
                
                ' Create Size.
                Set newElement = doc.createElement("one:Size")
                newElement.setAttribute "width", ww
                newElement.setAttribute "height", hh
                newElement.setAttribute "isSetByUser", "true"
                newNode.appendChild newElement
                
                'Push the image bnary data
                Set newElement = doc.createElement("one:Data")
                newElement.Text = base64String
                newNode.appendChild newElement
              
                ' Update OneNote with the new content.
               oneNote.UpdatePageContent doc.XML, , , True
               Dim strxml As String
                
               'Get the contnt back from OneNote Page
               oneNote.GetPageContent newPageID, strxml
               'Debug.Print strxml
                doc.LoadXML strxml
                Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
               
               nCounter = 1
               Do While nodes.Length = 0
                    oneNote.GetPageContent newPageID, strxml
                    doc.LoadXML strxml
                    Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
                    nCounter = nCounter + 1
                    If nCounter = 100 Then
                        strOCR = "Error: Image is not readable"
                        GoTo Before_end
                    End If
                Loop
                
                 
                Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
                strOCR = nodes(0).Text
                
            End If
        Else
            strOCR = "Error: OneNote 2010 Section nodes not found."
        End If
    Else
        strOCR = "Error: OneNote 2010 Section XML Data failed to load."
    End If
Else
    strOCR = "Error: OneNote 2010 XML Data failed to load."
End If
     
Before_end:
GetText = strOCR
End Function
 
Private Function GetAttributeValueFromNode(node As MSXML2.IXMLDOMNode, attributeName As String) As String

If node.Attributes.getNamedItem(attributeName) Is Nothing Then
    GetAttributeValueFromNode = "Not found."
Else
    GetAttributeValueFromNode = node.Attributes.getNamedItem(attributeName).Text
End If

End Function
 
Private Function GetFirstOneNoteNotebookNodes(oneNote As OneNote14.Application) As MSXML2.IXMLDOMNodeList

'Get the XML that represents the OneNote notebooks available.
Dim notebookXml As String
    
'OneNote fills notebookXml with an XML document providing information
'about what OneNote notebooks are available.
'You want all the data and thus are providing an empty string
'for the bstrStartNodeID parameter.

oneNote.GetHierarchy "", hsNotebooks, notebookXml, xs2010
     
'Use the MSXML Library to parse the XML.
Dim doc As MSXML2.DOMDocument60
Set doc = New MSXML2.DOMDocument60
     
If doc.LoadXML(notebookXml) Then
    Set GetFirstOneNoteNotebookNodes = doc.DocumentElement.getElementsByTagName("one:Notebook")
Else
    Set GetFirstOneNoteNotebookNodes = Nothing
End If

End Function
