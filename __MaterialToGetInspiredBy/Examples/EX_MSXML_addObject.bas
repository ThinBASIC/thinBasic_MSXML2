' ========================================================================================
' Demonstrates the use of the addObject method.
' The following example passes an object to the style sheet.
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
   LOCAL pXslDoc AS IXMLDOMDocument
   LOCAL pXslTempl AS IXSLTemplate
   LOCAL pXslProc AS IXSLProcessor
   LOCAL pXmlElement AS IXMLDOMElement
   LOCAL vOutput AS VARIANT

   pXslDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(pXslDoc) THEN EXIT FUNCTION
   pXmlDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION
   pXslTempl = NEWCOM "Msxml2.XSLTemplate.6.0"
   IF ISNOTHING(pXslTempl) THEN EXIT FUNCTION

   TRY
      pXslDoc.load("sampleXSLWithObject.xml")
      IF pXslDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pXmlElement = pXslDoc.documentElement
         pXslTempl.putref_stylesheet = pXmlElement
         pXslProc = pXslTempl.createProcessor
         pXmlDoc.loadXML "<level>Twelve</level>"
         pXslProc.input = pXmlDoc
         pXslProc.addObject pXmlDoc, "urn:my-object"
         pXslProc.transform
         vOutput = pXslProc.output
         AfxShowMsg VARIANT$$(vOutput)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
