' =========================================================================================
' Demonstrates the use of the nodeFromID method.
' Note: The following calls are needed to work with version 6.0.
' They aren't needed in versions 2.0 or 4.0.
' pXmlDoc.setProperty "ProhibitDTD", %VARIANT_FALSE
' pXmlDoc.resolveExternals %VARIANT_TRUE
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument2
   LOCAL pDOMNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.setProperty "ProhibitDTD", %VARIANT_FALSE
      pXmlDoc.resolveExternals = %VARIANT_TRUE
      pXmlDoc.load "people2.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDOMNode = pXmlDoc.nodeFromID("p3")
         AfxShowMsg pDOMNode.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
