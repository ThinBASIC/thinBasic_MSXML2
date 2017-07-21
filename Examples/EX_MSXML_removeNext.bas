' ========================================================================================
' Demonstrates the use of the removeNext method.
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
   LOCAL pDOMNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.setProperty "SelectionLanguage", "XPath"
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pSelection = pXmlDoc.selectNodes("//book")
         DO
            pDOMNode = pSelection.peekNode
            IF ISNOTHING(pDOMNode) THEN EXIT DO
            pDOMNode = pSelection.removeNext
         LOOP
      END IF
      AfxShowMsg pXmlDoc.xml
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
