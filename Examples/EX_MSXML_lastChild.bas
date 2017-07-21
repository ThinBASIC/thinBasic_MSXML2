' =========================================================================================
' Demonstrates the use of the lastChild property.
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
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pNewNode AS IXMLDOMNode
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pLastChild AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pRootNode = pXmlDoc.documentElement
      AfxShowMsg pRootNode.xml
      pNewNode = pXmlDoc.createNode(%NODE_ELEMENT, "VIDEOS", "")
      pLastChild = pRootNode.lastChild
      pCurrNode = pRootNode.insertBefore(pNewNode, pLastChild)
      AfxShowMsg pRootNode.xml
   END IF

END FUNCTION
' =========================================================================================
