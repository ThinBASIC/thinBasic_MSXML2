' ========================================================================================
' Demonstrates the use of the readyState (IXSLProcessor) property.
' Returns the current state of the processor.
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
   LOCAL iState AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pXslt = NEWCOM "Msxml2.XSLTemplate.6.0"
   IF ISNOTHING(pXslt) THEN EXIT FUNCTION
   pXslDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(pXslDoc) THEN EXIT FUNCTION

   pXslDoc.async = %VARIANT_FALSE
   IF ISTRUE pXslDoc.load("sample2.xsl") THEN
      pXslt.putref_stylesheet = pXslDoc
      pXmlDoc.async = %FALSE
      IF ISTRUE pXmlDoc.load("books.xml") THEN
         pXslProc = pXslt.createProcessor
         iState = pXslProc.readyState
         AfxShowMsg "State = " & FORMAT$(iState)
         pXslProc.input = pXmlDoc
         iState = pXslProc.readyState
         AfxShowMsg "State = " & FORMAT$(iState)
      END IF
   END IF

END FUNCTION
' ========================================================================================
