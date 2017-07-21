' =========================================================================================
' Demonstrates the use of the length property (IXMLDOMNamedNodeMap).
' Returns the number of items in the collection.
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument2
   LOCAL pNodeBook AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.setProperty "SelectionLanguage", "XPath"
   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pNodeBook = pXmlDoc.selectSingleNode("//book")
      pNamedNodeMap = pNodeBook.attributes
      AfxShowMsg STR$(pNamedNodeMap.length)
   END IF

END FUNCTION
' =========================================================================================
