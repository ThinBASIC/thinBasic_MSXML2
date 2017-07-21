' ========================================================================================
' Demonstrates the use of the target property.
' The following example iterates through the document's child nodes. If it finds a node of
' type NODE_PROCESSING_INSTRUCTION (7), it displays the node's target.
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
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL pPri AS IXMLDOMProcessingInstruction
   LOCAL nodeType AS LONG
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeList = pXmlDoc.childNodes
         FOR i = 0 TO pNodeList.length - 1
            pDOMNode = pNodeList.item(i)
            nodeType = pDOMNode.nodeType
            IF nodeType = %NODE_PROCESSING_INSTRUCTION THEN
               pPri = pDOMNode
               AfxShowMsg pPri.target
            END IF
         NEXT
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
