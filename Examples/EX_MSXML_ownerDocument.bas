' ========================================================================================
' Demonstrates the use of the ownerDocument property.
' The following example uses the ownerDocument property to return the parent DOMDocument
' object, and then displays that object's root element tag name.
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
   LOCAL pDocElem1 AS IXMLDOMElement
   LOCAL pDocElem2 AS IXMLDOMElement
   LOCAL pChildNodes1 AS IXMLDOMNodeList
   LOCAL pChildNodes2 AS IXMLDOMNodeList
   LOCAL pItem1 AS IXMLDOMNode
   LOCAL pItem2 AS IXMLDOMNode
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pOwner AS IXMLDOMDocument

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pDocElem1 = pXmlDoc.documentElement
      pChildNodes1 = pDocElem1.childNodes
      pItem1 = pChildNodes1.item(0)
      pChildNodes2 = pItem1.childNodes
      pItem2 = pChildNodes2.item(1)
      pOwner = pItem2.ownerDocument
      pDocElem2 = pOwner.documentElement
      AfxShowMsg pDocElem2.tagName
   END IF

END FUNCTION
' ========================================================================================
