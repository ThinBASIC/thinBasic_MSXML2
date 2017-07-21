' ========================================================================================
' Demonstrates the use of the length property (IXMLDOMNodeList).
' The following example creates an IXMLDOMNodeList object and then uses its length
' property to support iteration.
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
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pItem AS IXMLDOMNode
   LOCAL bstrOut AS STRING
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmLDoc.async = %FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pNodeList = pXmlDoc.getElementsByTagName("author")
      FOR i = 0 TO pNodeList.length - 1
         bstrOut += pNodeList.item(i).text & $CRLF
      NEXT
      AfxShowMsg bstrOut
   END IF

END FUNCTION
' ========================================================================================
