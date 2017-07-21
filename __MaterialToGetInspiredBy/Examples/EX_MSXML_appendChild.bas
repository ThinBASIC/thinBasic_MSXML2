' =========================================================================================
' Demonstrates the use of the appendChild method.
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
   LOCAL pNewNode AS IXMLDOMNode
   LOCAL xmlString AS STRING

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %FALSE
      pXmlDoc.load "appendChild.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         xmlString = pRootNode.xml
         AfxShowMsg "Before appendChild: " & $CRLF & xmlString
         pNewNode = pXmlDoc.createNode(%NODE_ELEMENT, "newChild", "")
         pRootNode.appendChild pNewNode
         xmlString = pRootNode.xml
         AfxShowMsg "After appendChild: " & $CRLF & xmlString
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
