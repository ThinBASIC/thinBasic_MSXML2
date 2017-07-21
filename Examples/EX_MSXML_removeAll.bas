' ========================================================================================
' Demonstrates the use of the removeAll method.
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
   LOCAL pSelection AS IXMLDOMSelection

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.setProperty "SelectionLanguage", "XPath"
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML "<root><child/><child/><child/></root>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         AfxShowMsg "Before removing <child> nodes: " & $CRLF & pXmlDoc.xml
         pSelection = pXmlDoc.selectNodes("//child")
         pSelection.removeAll
         AfxShowMsg "After removing <child> nodes: " & $CRLF & pXmlDoc.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
