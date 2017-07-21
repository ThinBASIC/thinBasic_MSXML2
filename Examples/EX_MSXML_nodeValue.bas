' ========================================================================================
' Demonstrates the use of the nodeValue property.
' The following example creates an IXMLDOMNode object and tests if it is a comment node.
' If it is, it displays its value.
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
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL strNodeType AS STRING
   LOCAL vValue AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML "<root><!-- Hello --></root>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pDOMNode = pRootNode.childNodes.item(0)
         strNodeType = pDOMNode.nodeTypeString
         IF strNodeType = "comment" THEN
            vValue = pDOMNode.nodeValue
            AfxShowMsg VARIANT$$(vValue)
         END IF
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
