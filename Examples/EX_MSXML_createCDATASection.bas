' ========================================================================================
' Demonstrates the use of createCDATASection.
' The following example creates a new CDATA section node.
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
   LOCAL pNodeCDATA AS IXMLDOMCDATASection

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pNodeCDATA = pXmlDoc.createCDATASection("Hello")
   AfxShowMsg pNodeCDATA.Xml

END FUNCTION
' ========================================================================================
