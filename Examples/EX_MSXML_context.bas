' ========================================================================================
' Demonstrates the use of the context property.
' The following example shows what is contained in the context of a selection after a
' query is executed. It also shows that the selection is reset when this property is
' changed.
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
      pXmlDoc.loadXML "<Customer>Microsoft</Customer>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pXmlDoc.setProperty "SelectionLanguage", "XPath"
         pSelection = pXmlDoc.selectNodes("//Customer")
         AfxShowMsg pSelection.context.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
