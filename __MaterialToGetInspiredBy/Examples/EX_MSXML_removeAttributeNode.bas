' ========================================================================================
' Demonstrates the use of the removeAttributeNode method.
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
   LOCAL pNodeBook AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pNodeId AS IXMLDOMAttribute
   LOCAL pDOMElement AS IXMLDOMElement

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.setProperty "SelectionLanguage", "XPath"
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeBook = pXmlDoc.selectSingleNode("//book")
         pNamedNodeMap = pNodeBook.attributes
         AfxShowMsg STR$(pNamedNodeMap.length)
         pDOMElement = pNodeBook
         pNodeId = pDOMElement.getAttributeNode("id")
         pDOMElement.removeAttributeNode pNodeId
         AfxShowMsg STR$(pNamedNodeMap.length)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
