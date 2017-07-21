' ========================================================================================
' Demonstrates the use of the addCollection method.
' Note  MSXML 6.0 has removed support for XDR schemas, whereas XDR is supported in
' MSXML 3.0 AND MSXML 4.0. If this method is called with an XDR schema, the call will fail.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument2
   LOCAL pSchemaCache AS IXMLDOMSchemaCollection
   LOCAL pSchemaCache2 AS IXMLDOMSchemaCollection

   ' XDR schemas are only supported in MSXML 3.0 AND MSXML 4.0.
   pXmlDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.4.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pSchemaCache = NEWCOM "Msxml2.XMLSchemaCache.4.0"
   IF ISNOTHING(pSchemaCache) THEN EXIT FUNCTION
   pSchemaCache2 = NEWCOM "Msxml2.XMLSchemaCache.4.0"
   IF ISNOTHING(pSchemaCache2) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.validateOnParse = %VARIANT_TRUE
      pSchemaCache.add "x-schema:books", "collection.xsd"
      pSchemaCache2.addCollection pSchemaCache
      pSchemaCache2.add "x-schema:books", "NewBooks.xsd"
      pXmlDoc.putref_schemas = pSchemaCache2
      ' The document will load only if a valid schema is attached to the xml file.
      pXmlDoc.load "collection.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         AfxShowMsg pXmlDoc.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
