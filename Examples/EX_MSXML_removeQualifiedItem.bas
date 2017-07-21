' ========================================================================================
' Demonstrates the use of the removeQualifiedItem method.
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
   LOCAL pDOMNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pNodeId AS IXMLDOMNode

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
         pDOMNamedNodeMap = pNodeBook.attributes
         AfxShowMsg STR$(pDOMNamedNodeMap.length)
         pNodeId = pDOMNamedNodeMap.removeQualifiedItem("id", "")
         AfxShowMsg STR$(pDOMNamedNodeMap.length)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
