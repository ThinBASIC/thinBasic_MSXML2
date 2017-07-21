' ========================================================================================
' Demonstrates the use of the getElementsByTagName (IXMLDOMElement) method.
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
   LOCAL pNodeBook AS IXMLDOMNode
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pElement AS IXMLDOMElement
   LOCAL pDOMNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISFALSE ISOBJECT(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.Load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeBook = pXmlDoc.selectSingleNode("//book")
         pElEment = pNodeBook
         pNodeList = pElement.getElementsByTagName("author")
         pDOMNode = pNodeList.item(0)
         AfxShowMsg pDOMNode.text
         ' Note: The above 3 lines can be replaced by:
         ' AfxShowMsg pElement.getElementsByTagName("author").item(0).text
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
