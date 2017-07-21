' =========================================================================================
' Demonstrates the use of the nextNode method.
' The following example creates an IXMLDOMNamedNodeMap object and uses its nextNode method
' to iterate the collection.
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
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pCurrNode = pRootNode.childNodes.item(1)
         pNamedNodeMap = pCurrNode.attributes
         FOR i = 0 TO pNamedNodeMap.length - 1
            pDOMNode = pNamedNodeMap.nextNode
            AfxShowMsg pDOMNode.text
            pDOMNode = NOTHING
         NEXT
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
