' ========================================================================================
' Demonstrates the use of spliText method
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
   LOCAL pNodeText AS IXMLDOMText
   LOCAL pNewNodeText AS IXMLDOMText
   LOCAL pNodeList AS IXMLDOMNodeList

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML "<root>Hello World!</root>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         ' Get a reference to the root node
         pRootNode = pXmlDoc.documentElement
         ' Get a reference to the list of child nodes
         pNodeList = pRootNode.childNodes
         ' Show how many nodes are in the list
         AfxShowMsg STR$(pNodeList.length)
         ' Get a reference to the first child node
         pDOMNode = pRootNode.firstChild
         ' Split the text
         pNodeText = pDOMNode
         pNewNodeText = pNodeText.splitText(6)
         ' Show how many nodes are in the list
         AfxShowMsg STR$(pNodeList.length)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
