' =========================================================================================
' Demonstrates the use of the nextNode method.
' The following example creates an IXMLDOMNodeList object and uses its nextNode method to
' iterate the collection.
' =========================================================================================

#DIM ALL
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pDOMNodeList AS IXMLDOMNodeList
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL bstrOut AS WSTRING
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDOMNodeList = pXmlDoc.getElementsByTagName("author")
         FOR i = 0 TO pDOMNodeList.length - 1
            pDOMNode = pDOMNodeList.nextNode
            bstrOut += pDOMNode.text & $CRLF
            pDOMNode = NOTHING
         NEXT
         AfxShowMsg bstrOut
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
