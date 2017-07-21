' ========================================================================================
' Demonstrates the use of the item method.
' The following example creates an IXMLDOMNodeList object with the document's
' getElementsByTagName method. It then iterates through the collection, displaying the
' text value of each item in the list.
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
   LOCAL bstrOut AS WSTRING
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pNodeList = pXmlDoc.getElementsByTagName("author")
      FOR i = 0 TO pNodeList.length - 1
         bstrOut += pNodeList.item(i).text & $CRLF
      NEXT
      AfxShowMsg bstrOut
   END IF

END FUNCTION
' ========================================================================================
