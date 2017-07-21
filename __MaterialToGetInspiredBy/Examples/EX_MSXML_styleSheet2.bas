' ========================================================================================
' Demonstrates the use of the styleSheet property (IXSLTemplate).
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXslt AS IXSLTemplate
   LOCAL pXslDoc AS IXMLDOMDocument

   pXslt = NEWCOM "Msxml2.XSLTemplate.6.0"
   IF ISNOTHING(pXslt) THEN EXIT FUNCTION
   pXslDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(pXslDoc) THEN EXIT FUNCTION

   TRY
      pXslDoc.async = %VARIANT_FALSE
      pXslDoc.load "sample2.xsl"
      IF pXslDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXslDoc.parseError.reason
      ELSE
         pXslt.putref_stylesheet = pXslDoc
         AfxShowMsg pXslt.stylesheet.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
