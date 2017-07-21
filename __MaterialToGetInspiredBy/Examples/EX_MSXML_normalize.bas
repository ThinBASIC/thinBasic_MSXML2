' ========================================================================================
' Demonstrates the use of the normalize method.
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
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pNodeList AS IXMLDOMNodeList

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML "<root/>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pRootNode.appendChild pXmlDoc.createTextNode("Hello ")
         pRootNode.appendChild pXmlDoc.createTextNode("World")
         pRootNode.appendChild pXmlDoc.createTextNode("!")
         pNodeList = pRootNode.childNodes
         AfxShowMsg "length = " & STR$(pNodeList.length)
         pRootNode.normalize
         AfxShowMsg "length = " & STR$(pNodeList.length)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
