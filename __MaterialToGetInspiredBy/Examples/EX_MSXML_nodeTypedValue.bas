' =======================================================================================
' Demonstrates the use of the nodeTypedValue property.
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
   LOCAL pXmlDocTest AS IXMLDOMDocument
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pChildNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pXmlDocTest = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(pXmlDocTest) THEN EXIT FUNCTION

   pRootNode = pXmlDoc.createElement("Test")
   pXmlDoc.putref_documentElement = pRootNode
   pChildNode = pXmLDoc.createNode(%NODE_TEXT, "", "")
   pRootNode.appendChild pChildNode
   pRootNode.dataType = "bin.hex"
   pChildNode.nodeTypedValue = "ffab123d"
   pXmlDocTest.async = %FALSE
   IF ISTRUE pXmlDocTest.load(pXmlDoc) THEN
      AfxShowMsg pXmLDocTest.xml
   END IF

END FUNCTION
' ========================================================================================
