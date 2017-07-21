' ========================================================================================
' Demonstrates the use of the validateNode method.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument3
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pSchemaCache AS IXMLDOMSchemaCollection
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pNode AS IXMLDOMNode
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pSchemaCache = NEWCOM "Msxml2.XMLSchemaCache.6.0"
   IF ISNOTHING(pSchemaCache) THEN EXIT FUNCTION

   pSchemaCache.add "urn:books", "validateNode.xsd"
   pXmlDoc.putref_schemas = pSchemaCache
   pXmlDoc.validateOnParse = %VARIANT_FALSE
   pXmlDoc.async = %VARIANT_FALSE

   pXmlDoc.load "validateNode.xml"
   IF pXmlDoc.parseError.errorCode THEN
      AfxShowMsg "You have error " & pXmlDoc.parseError.reason
   ELSE
      pNodeList = pXmlDoc.selectNodes("//book")
      FOR i = 0 TO pNodeList.length - 1
         pNode = pNodeList.item(i)
         pParseError = pXmlDoc.validateNode(pNode)
         IF OBJRESULT = %S_FALSE THEN
            AfxShowMsg "Invalid node - Error &H" & HEX$(pParseError.errorCode, 8)
            AfxShowMsg pParseError.reason
         ELSEIF OBJRESULT <> %S_OK THEN
            AfxShowMsg "Error &H" & HEX$(OBJRESULT)
         ELSE
            AfxShowMsg pNode.xml
         END IF
         pParseError = NOTHING
         pNode = NOTHING
      NEXT
   END IF

END FUNCTION
' ========================================================================================
