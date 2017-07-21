' ========================================================================================
' Demonstrates the use of the documentElement property.
' The following example creates an IXMLDOMElement object and sets it to the root element
' of the document with the documentElement property. It then walks the document tree.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXMLDoc AS IXMLDOMDocument
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL i AS LONG

   pXMLDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXMLDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXMLDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXMLDoc.documentElement
         pNodeList = pRootNode.childNodes
         FOR i = 0 TO pNodeList.length - 1
            pCurrNode = pNodeList.item(i)
            AfxShowMsg pCurrNode.text
            pCurrNode = NOTHING
         NEXT
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
