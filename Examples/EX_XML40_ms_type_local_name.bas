' ========================================================================================
' Demonstrates the use of the ms:type-local-name XPath Extension function when
' programming the MSXML DOM.
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
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pNode AS IXMLDOMNode
   LOCAL i AS LONG

   ' only works with version 4.0
   pXmlDoc = NEWCOM "Msxml2.DOMDocument.4.0"
   IF ISFALSE ISOBJECT(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.setProperty "SelectionNamespaces", "xmlns:ms='urn:schemas-microsoft-com:xslt'"
      pXmlDoc.load "books3.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNodeList = pXmlDoc.selectNodes("//*[ms:type-local-name(.)='AuthorType']")
         FOR i = 0 TO pNodeList.length - 1
            AfxShowMsg pNodeList.item(i).nodeName
         NEXT
         pNodeList = NOTHING
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
