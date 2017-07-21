' ========================================================================================
' Demonstrates the use of the startMode property
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
   LOCAL pXslt AS IXSLTemplate
   LOCAL pXslDoc AS IXMLDOMDocument
   LOCAL pXslProc AS IXSLProcessor

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
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
         pXmlDoc.async = %VARIANT_FALSE
         pXmlDoc.load "books.xml"
         IF pXmlDoc.parseError.errorCode THEN
            AfxShowMsg "You have error " & pXmlDoc.parseError.reason
         ELSE
            pXslProc = pXslt.createProcessor
            pXslProc.input = pXmlDoc
            pXslProc.setStartMode "view"
            AfxShowMsg pXslProc.startMode
         END IF
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
