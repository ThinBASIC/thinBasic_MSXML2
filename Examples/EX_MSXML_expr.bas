' ========================================================================================
' Demonstrates the use of the expr property.
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
      pXmlDoc.loadXML("<Customer><Name>Microsoft</Name></Customer>")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pXmlDoc.setProperty "SelectionLanguage", "XPath"
         pSelection = pXmlDoc.selectNodes("Customer/Name")
         AfxShowMsg pSelection.expr & " --- " & pSelection.item(0).xml
         pSelection.expr = "/Customer"
         AfxShowMsg pSelection.expr & " --- " & pselection.item(0).xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
