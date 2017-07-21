' ========================================================================================
' Demonstrates the use of the data (IXMLDOMProcessingInstruction) property.
' The following example displays the node's data of the xml processing instruction.
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
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL pri AS IXMLDOMProcessingInstruction

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDOMNode = pXMLDoc.firstChild
         IF pDOMNode.nodeType = %NODE_PROCESSING_INSTRUCTION THEN
            pri = pDOMNode
            AfxShowMsg pri.Data
            pri = NOTHING
         END IF
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
