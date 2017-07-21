' ========================================================================================
' Demonstrates the use of the insertBefore method.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pPri AS IXMLDOMProcessingInstruction
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pItem AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   IF ISTRUE pXmlDoc.load("books.xml") THEN
		pPri = pXmlDoc.createProcessingInstruction("xml", "version=""1.0""")
      pNodeList = pXmlDoc.childNodes
      pItem = pNodeList.item(0)
      pXmlDoc.insertBefore pPri, pItem
      AfxShowMsg pXmlDoc.xml
   END IF

END FUNCTION
' ========================================================================================
