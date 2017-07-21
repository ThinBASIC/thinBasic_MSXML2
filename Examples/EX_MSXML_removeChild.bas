' ========================================================================================
' Demonstrates the use of the removeChild method.
' The following example creates an IXMLDOMNode object (currNode), removes a child node
' from it, and displays the text of the removed node.
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
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pChildNode AS IXMLDOMNode
   LOCAL pOldChild AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      IF ISTRUE pXmlDoc.load("books.xml") THEN
         pRootNode = pXmlDoc.documentElement
         pCurrNode = pRootNode.childNodes.item(1)
         pChildNode = pCurrNode.childNodes.item(1)
         pOldChild = pCurrNode.removeChild(pChildNode)
         AfxShowMsg pOldChild.text
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
