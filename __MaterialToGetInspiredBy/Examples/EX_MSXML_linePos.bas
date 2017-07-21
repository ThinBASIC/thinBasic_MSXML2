' ========================================================================================
' Demonstrates the use of the linePos property.
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
   LOCAL pError AS IXMLDOMParseError
   LOCAL pElement AS IXMLDOMElement

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      pError = pXMLDoc.parseError
      IF pError.errorCode THEN
         AfxShowMsg "A parse error occurred on line " & STR$(pError.line) & _
                    " at position " & STR$(pError.linePos)
      ELSE
         pElement = pXmlDoc.documentElement
         AfxShowMsg pElement.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
