' ========================================================================================
' Demonstrates the use of the errorXPath property.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc     AS IXMLDOMDocument3
   LOCAL pSCache     AS IXMLDOMSchemaCollection
   LOCAL pError      AS IXMLDOMParseError2
   LOCAL bstrMsg     AS WSTRING

   ' Create an instance of XML DOM
   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN
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
   pXmlDoc.putref_schemas = pSCache

   ' Set the MultipleErrorMessages property
   pXmlDoc.async = %VARIANT_FALSE
   pXmlDoc.validateOnParse = %VARIANT_FALSE

   ' Load books.xml
   pXmlDoc.load("books.xml")
   IF pXmlDoc.parseError.errorCode THEN
      AfxShowMsg "Failed to load DOM from books.xml" & $CRLF & pXmlDoc.parseError.reason
      EXIT FUNCTION
   END IF

   ' Validate the entire DOM object
   pError = pXmlDoc.validate
   IF pError.errorCode <> 0 THEN
      bstrMsg = "Error as returned from validate():" & $CRLF & _
              "Error code: " & FORMAT$(pError.errorCode) & $CRLF & _
              "Error reason: " & pError.reason & $CRLF & _
              "Error location: " & pError.errorXPath & $CRLF
      AfxShowMsg bstrMsg
   ELSE
      AfxShowMsg "DOM is valid."
   END IF

END FUNCTION
' ========================================================================================
