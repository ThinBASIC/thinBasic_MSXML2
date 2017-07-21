' =========================================================================================
' Demonstrates the use of the item method (IXMLDOMNamedNodeMap).
' The following example creates an IXMLDOMNamedNodeMap object to retrieve the attributes
' for an element node selected using the SelectSingleNode method. It then iterates through
' the attributes, before displaying the name and value of each attribute in the collection.
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pBookNode AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pItem AS IXMLDOMNode
   LOCAL iLen AS LONG
   LOCAL i AS LONG
   LOCAL vValue AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pBookNode = pXmlDoc.selectSingleNode("//book")
      pNamedNodeMap = pBookNode.attributes
      iLen = pNamedNodeMap.length
      FOR i = 0 TO iLen - 1
         pItem = pNamedNodeMap.item(i)
         AfxShowMsg "Attribute name: " & pItem.nodeName
         vValue = pItem.nodeValue
         AfxShowMsg "Attribute value: " & VARIANT$$(vValue)
         pItem = NOTHING
      NEXT
   END IF

END FUNCTION
' =========================================================================================
