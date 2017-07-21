' ========================================================================================
' Demonstrates the use of the output property.
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
   LOCAL vRes AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pXslt = NEWCOM "Msxml2.XSLTemplate.6.0"
   IF ISNOTHING(pXslt) THEN EXIT FUNCTION
   pXslDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(pXslDoc) THEN EXIT FUNCTION

   TRY
      pXslDoc.async = %VARIANT_FALSE
      IF ISTRUE pXslDoc.load("Sample2.xsl") THEN
         pXslt.putref_stylesheet = pXslDoc
         pXmlDoc.async = %FALSE
         IF ISTRUE pXmlDoc.load("books.xml") THEN
            pXslProc = pXslt.createProcessor
            pXslProc.input = pXmlDoc
            pXslProc.transform
            vRes = pXslProc.output
            AfxShowMsg VARIANT$$(vRes)
         END IF
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
