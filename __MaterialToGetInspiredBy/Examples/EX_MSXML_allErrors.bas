' ========================================================================================
' Demonstrates the use of the allErrors property.
' The example uses two resource files, books.xml and books.xsd. The first is an XML data
' file, and the second is the XML Schema for the XML data. The XML document has two
' invalid <book> elements: <book id="bk002"> and <book id="bk003">. The sample application
' uses several methods and properties of the IXMLDOMParseError2 interface to examine the
' resulting parse errors.
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
   IF pXMLDoc.load("books.xml") <> %VARIANT_TRUE THEN
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "Failed to load DOM from books.xml" & $CRLF & pXmlDoc.parseError.reason
         EXIT FUNCTION
      END IF
   END IF

   ' Validate the entire DOM object
   pError = pXMLDoc.validate
   IF pError.errorCode <> 0 THEN
      bstrMsg = "Error as returned from validate():" & $CRLF & _
               "Error code: " & FORMAT$(pError.errorCode) & $CRLF & _
               "Error reason: " & pError.reason & $CRLF & _
               "Error location: " & pError.errorXPath & $CRLF
      pErrors = pError.allErrors
      IF ISOBJECT(pErrors) THEN
         bstrErrors = "Errors count: " & FORMAT$(pErrors.length) & $CRLF & _
                     "Error items from the allErrors collection: " & $CRLF
         FOR i = 0 TO pErrors.length - 1
            pEitem = pErrors.item(i)
            IF ISOBJECT(pEitem) THEN
               bstrErrors = bstrErrors & "Error item: " & FORMAT$(i) & $CRLF & _
                           "reason: " & pEitem.reason & $CRLF & _
                           "location: " & pEitem.errorXPath
               pEItem  = NOTHING
            END IF
         NEXT
         pErrors = NOTHING
      END IF
      AfxShowMsg bstrMsg & bstrErrors
   END IF

END FUNCTION
' ========================================================================================
