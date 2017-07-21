' =========================================================================================
' Demonstrates the use of the createTextNode method.
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pTextNode AS IXMLDOMText
   LOCAL pChildNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML("<root><child/></root>")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pTextNode = pXmlDoc.createTextNode("Hello World")
         pChildNode = pRootNode.childNodes.item(0)
         pRootNode.insertBefore pTextNode, pChildNode
         AfxShowMsg pRootNode.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
