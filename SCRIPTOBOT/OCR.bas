Attribute VB_Name = "OCR"
Option Explicit
Option Private Module

Public Function GetText(strFile As String, Optional bArray As Variant, Optional imgFile As Object)
'Get all of the Notebook nodes.
Dim oneNote As OneNote14.Application: Set oneNote = New OneNote14.Application
Dim nodes As MSXML2.IXMLDOMNodeList: Set nodes = GetFirstOneNoteNotebookNodes(oneNote)

Dim bitmap As Object, bytFile() As Byte, base64String

If Not imgFile Is Nothing Then
    Set bitmap = imgFile
    bytFile = bArray
    base64String = EncodeBase64(bytFile)
Else
    Set bitmap = LoadPicture(strFile)
    bytFile = GetFileBytes(strFile)
    base64String = EncodeBase64(bytFile)
End If

Dim hh As Long: hh = Round(bitmap.Width / 30)
Dim ww As Long: ww = Round(bitmap.height / 30)
  
If Not nodes Is Nothing Then
    'Get the first OneNote Notebook in the XML document.
    Dim node As MSXML2.IXMLDOMNode: Set node = nodes(0)
    Dim noteBookName As String: noteBookName = node.Attributes.getNamedItem("name").text

    'Get the ID for the Notebook so the code can retrieve the list of sections.
    Dim notebookID As String: notebookID = node.Attributes.getNamedItem("ID").text

    'Load the XML for the Sections for the Notebook requested.
    Dim sectionsXml As String: oneNote.GetHierarchy notebookID, hsSections, sectionsXml, xs2010

    Dim secDoc As MSXML2.DOMDocument60: Set secDoc = New MSXML2.DOMDocument60

    If secDoc.LoadXML(sectionsXml) Then
        'Select the Section nodes
        Dim secNodes As MSXML2.IXMLDOMNodeList: Set secNodes = secDoc.DocumentElement.getElementsByTagName("one:Section")

        If Not secNodes Is Nothing Then
            'Get the first section.
            Dim secNode As MSXML2.IXMLDOMNode: Set secNode = secNodes(0)

            Dim sectionName As String: sectionName = secNode.Attributes.getNamedItem("name").text
            Dim sectionID As String: sectionID = secNode.Attributes.getNamedItem("ID").text
            
            'Create a new blank Page in the first Section using the default format.
            Dim newPageID As String: oneNote.CreateNewPage sectionID, newPageID, npsDefault
                
            'Get the contents of the page.
            Dim outXML As String: oneNote.GetPageContent newPageID, outXML, piAll, xs2010

            Dim doc As MSXML2.DOMDocument60: Set doc = New MSXML2.DOMDocument60

            'Load Page's XML into a MSXML2.DOMDocument60 object.
            If doc.LoadXML(outXML) Then
                'Get Page Node.
                Dim pageNode As MSXML2.IXMLDOMNode: Set pageNode = doc.getElementsByTagName("one:Page")(0)
                
                'Create Outline node.
                Dim newElement As MSXML2.IXMLDOMElement: Set newElement = doc.createElement("one:Outline")
                newElement.setAttribute "lang", "en-US"
                Dim newNode As MSXML2.IXMLDOMNode: Set newNode = pageNode.appendChild(newElement)
                
                'Create OEChildren.
                Set newElement = doc.createElement("one:OEChildren")
                Set newNode = newNode.appendChild(newElement)

                'Create OE.
                Set newElement = doc.createElement("one:OE")
                newElement.setAttribute "lang", "en-US"
                Set newNode = newNode.appendChild(newElement)
   
                'Create Image.
                Set newElement = doc.createElement("one:Image")
                
                'newElement.setAttribute "format", "bmp"
                Set newNode = newNode.appendChild(newElement)

                'Create Size.
                Set newElement = doc.createElement("one:Size")
                
                With newElement
                    .setAttribute "width", ww
                    .setAttribute "height", hh
                    .setAttribute "isSetByUser", "true"
                End With
                
                newNode.appendChild newElement

                'Push the image bnary data
                Set newElement = doc.createElement("one:Data")
                newElement.text = base64String
                newNode.appendChild newElement

                'Update OneNote with the new content.
                oneNote.UpdatePageContent doc.XML, , , True
                
                'Get the contnt back from OneNote Page
                Dim strxml As String: oneNote.GetPageContent newPageID, strxml
                'Debug.Print strxml
                doc.LoadXML strxml
                Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")

                Dim nCounter As Integer: nCounter = 1
                Do While nodes.Length = 0
                    oneNote.GetPageContent newPageID, strxml
                    doc.LoadXML strxml
                    Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
                    
                    nCounter = nCounter + 1
    
                    If nCounter = 100 Then
                        Dim strOCR As String: strOCR = "Error: Image is not readable"
                        GoTo Before_end
                    End If
                Loop

                Set nodes = doc.DocumentElement.getElementsByTagName("one:OCRText")
                strOCR = nodes(0).text
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

Before_end: GetText = strOCR
End Function

Private Function GetFirstOneNoteNotebookNodes(oneNote As OneNote14.Application) As MSXML2.IXMLDOMNodeList
'Get the XML that represents the OneNote notebooks available.
'OneNote fills notebookXml with an XML document providing information
'about what OneNote notebooks are available.
'You want all the data and thus are providing an empty string
'for the bstrStartNodeID parameter.
Dim notebookXml As String: oneNote.GetHierarchy "", hsNotebooks, notebookXml, xs2010

'Use the MSXML Library to parse the XML.
Dim doc As MSXML2.DOMDocument60: Set doc = New MSXML2.DOMDocument60

If doc.LoadXML(notebookXml) Then
    Set GetFirstOneNoteNotebookNodes = doc.DocumentElement.getElementsByTagName("one:Notebook")
Else
    Set GetFirstOneNoteNotebookNodes = Nothing
End If
End Function

Private Function GetFileBytes(strPath As String) As Byte()
With CreateObject("ADODB.Stream")
    .Open
    .Type = 1 'adTypeBinary
    .LoadFromFile strPath
    GetFileBytes = .Read
    .Close
End With
End Function

Private Function EncodeBase64(arrData() As Byte) As String
Dim objXML As MSXML2.DOMDocument60: Set objXML = New MSXML2.DOMDocument60
Dim objNode As MSXML2.IXMLDOMElement: Set objNode = objXML.createElement("b64")
  
objNode.DataType = "bin.base64"
objNode.nodeTypedValue = arrData
EncodeBase64 = objNode.text

Set objNode = Nothing
Set objXML = Nothing
End Function
