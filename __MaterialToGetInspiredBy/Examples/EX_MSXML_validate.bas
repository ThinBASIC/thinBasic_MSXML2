' ========================================================================================
' Demonstrates the use of the validate method.
' When used with MSXML 6.0, you should get the following output:
' Invalid Dom: Element 'review' is unexpected according to content model of parent
' element 'book'
' Expecting: price.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL i AS LONG
   LOCAL pXmlDoc AS IXMLDOMDocument2
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pSchemaCache AS IXMLDOMSchemaCollection
   LOCAL pNodeList AS IXMLDOMNodeList

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pSchemaCache = NEWCOM "Msxml2.XMLSchemaCache.6.0"
   IF ISNOTHING(pSchemaCache) THEN EXIT FUNCTION

   pSchemaCache.add "urn:books", "validateNode.xsd"
   pXmlDoc.putref_schemas = pSchemaCache
   pXmlDoc.validateOnParse = %VARIANT_FALSE
   pXmlDoc.async = %VARIANT_FALSE
   pXmlDoc.load "validateNode.xml"
   pParseError = pXmlDoc.validate
   IF OBJRESULT = %S_FALSE THEN
      AfxShowMsg "Invalid DOM - Error &H: " & HEX$(pParseError.errorCode, 8)
      AfxShowMsg pParseError.reason
   ELSEIF OBJRESULT = %E_PENDING THEN
      AfxShowMsg "The document is not completely loaded."
   ELSEIF OBJRESULT <> %S_OK THEN
      AfxShowMsg "Error &H" & HEX$(OBJRESULT)
   ELSE
      AfxShowMsg "DOM is valid: " & pXmlDoc.xml
      pNodeList = pXmlDoc.selectNodes("//book")
      FOR i = 0 TO pNodeList.length - 1
         AfxShowMsg pNodeList.item(i).xml
      NEXT
   END IF

END FUNCTION
' ========================================================================================
