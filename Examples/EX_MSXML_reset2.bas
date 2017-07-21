' ========================================================================================
' Demonstrates the use of the reset method (IXMLDOMNodeList).
' The following script example creates an IXMLDOMNodeList object and iterates the
' collection using the nextNode method. It then uses the reset method to reset the
' iterator to point before the first node in the list.
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
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pNode AS IXMLDOMNode
   LOCAL bstrOut AS WSTRING
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.setProperty "SelectionLanguage", "XPath"
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDOc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeList = pXmlDoc.getElementsByTagName("author")
         FOR i = 0 TO pNodeList.length - 1
            bstrOut += pNodeList.nextNode.text & $CRLF
         NEXT
         AfxShowMsg bstrOut
         pNodeList.reset
         AfxShowMsg pNodeList.nextNode.text
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
