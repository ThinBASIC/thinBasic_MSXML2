' ========================================================================================
' Demonstrates the use of the length (IXMLDOMParseErrorCollection) property.
' The following code performs an XSD validation on an XML document that has two invalid
' <book> nodes. The code then outputs the number of errors in the resultant error
' collection.
' ========================================================================================

#DIM ALL
#DEBUG ERROR ON
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXMLDoc     AS IXMLDOMDocument3
   LOCAL pSCache     AS IXMLDOMSchemaCollection
   LOCAL pError      AS IXMLDOMParseError2
   LOCAL pErrors     AS IXMLDOMParseErrorCollection
   LOCAL bstrMsg     AS WSTRING
   LOCAL errsCount   AS LONG

   ' Create an instance of XML DOM
   pXMLDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXMLDoc) THEN EXIT FUNCTION

   ' Create an instance of schema caché
   pSCache = NEWCOM "Msxml2.XMLSchemaCache.6.0"
   IF ISNOTHING(pSCache) THEN EXIT FUNCTION

   ' Add "urn:books" from "books.xsd" to schema caché
   pSCache.add "urn:books", "books.xsd"

   ' Set the reference
   pXMLDoc.putref_schemas = pSCache

   ' Set the MultipleErrorMessages property
   pXMLDoc.async = %VARIANT_FALSE
   pXMLDoc.validateOnParse = %VARIANT_FALSE
   pXMLdoc.setProperty "MultipleErrorMessages", %VARIANT_TRUE

   ' Load books.xml
   IF pXMLDoc.load("books.xml") <> %VARIANT_TRUE THEN
      pError = pXmlDoc.parseError
      AfxShowMsg "Failed to load DOM from books.xml" & $CRLF & pError.reason
      pError = NOTHING
      EXIT FUNCTION
   END IF

   ' Validate the entire DOM object
   pError = pXMLDoc.validate
   IF pError.errorCode <> 0 THEN
      pErrors = pError.allErrors
      IF ISOBJECT(pErrors) THEN
         errsCount = pErrors.length
         AfxShowMsg "There are " & STR$(errsCount) & " errors in the collection"
         pErrors = NOTHING
      END IF
   ELSE
      AfxShowMsg "DOM is valid " & $CRLF & pXmlDoc.xml
   END IF

END FUNCTION
' ========================================================================================
