' =========================================================================================
' Demonstrates the use of the parentNode property.
' The following example sets a variable ('pParentNode') to reference the parent node of
' another IXMLDOMNode object ('pChildNode'). It then uses the reference to the new node to
' display the XML contents of its parent node.
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
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pParentNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pRootNode = pXmlDoc.documentElement
      pCurrNode = pRootNode.childNodes.item(1)
      pParentNode = pCurrNode.childNodes.item(0).parentNode
      AfxShowMsg pParentNode.xml
   END IF

END FUNCTION
' =========================================================================================
