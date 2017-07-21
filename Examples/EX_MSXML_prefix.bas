' ========================================================================================
' Demonstrates the use of prefix property.
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
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL bstrNamespaceURI AS WSTRING
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books2.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         ' Get the namespace of the root element.
         pRootNode = pXmlDoc.documentElement
         bstrNamespaceURI = pRootNode.prefix
         AfxShowMsg bstrNamespaceURI
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
