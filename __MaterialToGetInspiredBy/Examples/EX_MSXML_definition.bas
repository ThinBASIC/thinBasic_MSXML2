' ========================================================================================
' The following example shows the retrieval of the definition property from an
' IXMLDOMElement.
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
   LOCAL pIXMLDOMElement AS IXMLDOMElement
   LOCAL pIXMLDOMNode AS IXMLDOMNode

   ' XDR schemas are not supported in MSXML 6.0.
   ' Therefore, we must use version 3.0 or 4.0.
   pXmlDoc = NEWCOM "Msxml2.DOMDocument.4.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "sample.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pIXMLDOMElement = pXMLDoc.documentElement
         pIXMLDOMNode = pIXMLDOMElement.definition
         IF ISOBJECT(pIXMLDOMNode) THEN AfxShowMsg pIXMLDOMNode.xml
         ' Note: The first two lines above can be replaced by:
         ' pIXMLDOMNode = pXmlDoc.documentElement.definition
         ' And the three lines above can be replaced by:
         ' AfxShowMsg pXmlDoc.documentElement.definition.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
