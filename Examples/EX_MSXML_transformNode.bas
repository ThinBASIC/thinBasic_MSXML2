' ========================================================================================
' Demonstrates the use of the transformNode method.
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
   LOCAL pStyleSheet AS IXMLDOMDocument

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pStyleSheet = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pStyleSheet) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pStyleSheet.async = %VARIANT_FALSE
         pStyleSheet.load "sample2.xsl"
         IF pStyleSheet.parseError.errorCode THEN
            AfxShowMsg "You have error " & pStyleSheet.parseError.reason
         ELSE
            AfxShowMsg pXmlDoc.transformnode(pStyleSheet)
         END IF
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
