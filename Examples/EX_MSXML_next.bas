' ========================================================================================
' Demonstrates the use of the next property.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXMLDoc     AS IXMLDOMDocument3
   LOCAL pSCache     AS IXMLDOMSchemaCollection
   LOCAL pEItem      AS IXMLDOMParseError2
   LOCAL pError      AS IXMLDOMParseError2
   LOCAL pErrors     AS IXMLDOMParseErrorCollection
   LOCAL bstrMsg     AS WSTRING
   LOCAL bstrErrors  AS WSTRING
   LOCAL i           AS LONG

   ' Create an instance of XML DOM
   pXMLDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXMLDoc) THEN
      AfxShowMsg "Failed to create an instance on an XML DOM"
      EXIT FUNCTION
   END IF

   ' Create an instance of schema caché
   pSCache = NEWCOM "Msxml2.XMLSchemaCache.6.0"
   IF ISNOTHING(pSCache) THEN
      AfxShowMsg "Cannot instantiate XMLSchemaCache60"
      EXIT FUNCTION
   END IF

   ' Add "urn:books" from "books.xsd" to schema caché
   pSCache.add "urn:books", "books.xsd"
   IF OBJRESULT THEN
      AfxShowMsg "Cannot add 'urn:books' to schema caché. Error &H" & HEX$(OBJRESULT)
      EXIT FUNCTION
   END IF
   ' Set the reference
   pXMLDoc.putref_schemas = pSCache

   ' Set the MultipleErrorMessages property
   pXMLDoc.async = %VARIANT_FALSE
   pXMLDoc.validateOnParse = %VARIANT_FALSE
   pXMLdoc.setProperty "MultipleErrorMessages", %VARIANT_TRUE
   IF OBJRESULT THEN
      AfxShowMsg "Failed to enable mulitple validation errors"
      EXIT FUNCTION
   END IF

   ' Load books.xml
   pXMLDoc.load "books.xml"
   IF pXmlDoc.parseError.errorCode THEN
      AfxShowMsg "Failed to load DOM from books.xml" & $CRLF & pXmlDoc.parseError.reason
      EXIT FUNCTION
   END IF

   ' Validate the entire DOM object
   pError = pXMLDoc.validate
   IF pError.errorCode <> 0 THEN
      pErrors = pError.allErrors
      IF ISOBJECT(pErrors) THEN
         DO
            pEItem = pErrors.next
            IF ISNOTHING(pEItem) THEN EXIT DO
            bstrErrors += "errorItem[" & FORMAT$(i) & "]: " & pEItem.reason & $CRLF
            pEItem = NOTHING
            i = i + 1
         LOOP
         pErrors = NOTHING
      END IF
      AfxShowMsg bstrMsg & bstrErrors
   ELSE
      AfxShowMsg "DOM is valid"
   END IF

END FUNCTION
' ========================================================================================
