' ========================================================================================
' Demonstrates the use of the transformNodeToObject method.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pSource AS IXMLDOMDocument
   LOCAL pStylesheet AS IXMLDOMDocument
   LOCAL pStylesheet2 AS IXMLDOMDocument
   LOCAL pResult AS IXMLDOMDocument
   LOCAL pResult2 AS IXMLDOMDocument

   pSource = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pSource) THEN EXIT FUNCTION
   pStylesheet = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pStylesheet) THEN EXIT FUNCTION
   pStylesheet2 = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pStylesheet2) THEN EXIT FUNCTION
   pResult = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pResult) THEN EXIT FUNCTION
   pResult2 = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pResult2) THEN EXIT FUNCTION

   TRY
      pSource.async = %VARIANT_FALSE
      pSource.load "sample.xml"
      IF pSource.parseError.errorCode THEN
         AfxShowMsg "You have error " & pSource.parseError.reason
      ELSE
         pStyleSheet.async = %VARIANT_FALSE
         pStyleSheet.load "stylesheet1.xsl"
         IF pStyleSheet.parseError.errorCode THEN
            AfxShowMsg "You have error " & pStyleSheet.parseError.reason
         ELSE
            pResult.async = %VARIANT_FALSE
            pResult.validateOnParse = %VARIANT_TRUE
            pResult2.async = %VARIANT_FALSE
            pResult2.validateOnParse = %VARIANT_TRUE
            pSource.transformNodeToObject pStylesheet, pResult
            pStyleSheet2.async = %VARIANT_FALSE
            pStyleSheet2.load "stylesheet2.xsl"
            IF pStyleSheet2.parseError.errorCode THEN
               AfxShowMsg "You have error " & pStyleSheet2.parseError.reason
            ELSE
               pResult.transformNodeToObject pStylesheet2, pResult2
               AfxShowMsg pResult2.xml
            END IF
         END IF
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
