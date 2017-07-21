' ========================================================================================
' Demonstrates the use of the add method.
' The following example attaches a schema to an XML document.
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

   pXmlDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.4.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pSchemaCache = NEWCOM "Msxml2.XMLSchemaCache.4.0"
   IF ISNOTHING(pSchemaCache) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.validateOnParse = %VARIANT_FALSE
      pSchemaCache.add "x-schema:books", "collection.xdr"
      pXmlDoc.putref_schemas = pSchemaCache
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
