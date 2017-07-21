' ========================================================================================
' Demonstrates the use of the errorParameters method.
' This sample code uses features that were first implemented in MSXML 5.0 for Microsoft
' Office Applications.
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
   LOCAL pNodes      AS IXMLDOMNodeList
   LOCAL pNode       AS IXMLDOMNode
   LOCAL bstrMsg     AS WSTRING
   LOCAL i           AS LONG
   LOCAL x           AS LONG

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
   pXmlDoc.setProperty "MultipleErrorMessages", %VARIANT_TRUE
   IF OBJRESULT THEN
      AfxShowMsg "Failed to enable mulitple validation errors"
      EXIT FUNCTION
   END IF

   ' Load books.xml
   IF pXmlDoc.load("books.xml") <> %VARIANT_TRUE THEN
      pError = pXmlDoc.parseError
      AfxShowMsg "Failed to load DOM from books.xml" & $CRLF & pError.reason
      pError = NOTHING
      EXIT FUNCTION
   END IF

   ' Validate the entire DOM object
   pError = pXmlDoc.validate
   bstrMsg = "Validating DOM..." & $CRLF
   IF pError.errorCode <> 0 THEN
      bstrMsg = bstrMsg & "invalid DOM:" & $CRLF & _
                        "   code: " & FORMAT$(pError.errorCode) & $CRLF & _
                        "   reason: " & pError.reason & $CRLF & _
                        "   errorXPath: " & pError.errorXPath & $CRLF & _
                        "Parameters count: " & FORMAT$(pError.errorParametersCount) & $CRLF
      FOR i = 0 TO pError.errorParametersCount - 1
         bstrMsg = bstrMsg & "   errorParameters(" & FORMAT$(i) & "):" & pError.errorParameters(i) & $CRLF
      NEXT
      AfxShowMsg bstrMsg
   ELSE
      AfxShowMsg "DOM is valid" & $CRLF & pXmlDoc.xml
   END IF
   pError  = NOTHING

   bstrMsg = bstrMsg & $CRLF
   bstrMsg = bstrMsg & "Validating nodes..." & $CRLF
   pNodes = pXmlDoc.selectNodes("//book")
   FOR i = 0 TO pNodes.length - 1
      pNode = pNodes.Item(i)
      pError = pXmlDoc.validateNode(pNode)
      IF pError.errorCode <> 0 THEN
         bstrMsg = bstrMsg & $CRLF
         bstrMsg = bstrMsg & "Node is invalid: " & $CRLF
         bstrMsg = bstrMsg & "   reason: " & pError.reason & $CRLF
         bstrMsg = bstrMsg & "   errorXPath: " & pError.errorXPath & $CRLF
         bstrMsg = bstrMsg & "Parameters count: " & FORMAT$(pError.errorParametersCount) & $CRLF
         FOR x = 0 TO pError.errorParametersCount - 1
            bstrMsg = bstrMsg & "   errorParameter(" & FORMAT$(x) & "):" & pError.errorParameters(x) & $CRLF
         NEXT
         AfxShowMsg bstrMsg
      ELSE
         bstrMsg = bstrMsg & "Node is valid:"
         bstrMsg = bstrMsg & pNode.xml
      END IF
   NEXT
   pError = NOTHING

   pSCache = NOTHING
   pXmlDoc = NOTHING

END FUNCTION
' ========================================================================================
