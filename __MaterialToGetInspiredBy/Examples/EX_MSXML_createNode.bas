' =========================================================================================
' Demonstrates the use of the createNode method.
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
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pLastChild AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmLDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         AfxShowMsg pRootNode.xml
         pNewNode = pXmlDoc.createNode(%NODE_ELEMENT, "VIDEOS", "")
         pLastChild = pRootNode.lastChild
         pCurrNode = pRootNode.insertBefore(pNewNode, pLastChild)
         AfxShowMsg pCurrNode.xml
         AfxShowMsg pRootNode.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
