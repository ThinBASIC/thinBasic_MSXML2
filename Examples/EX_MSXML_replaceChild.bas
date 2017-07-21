' ========================================================================================
' Demonstrates the use of the replaceChild method.
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
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pNewElem AS IXMLDOMElement
   LOCAL pChildNodes AS IXMLDOMNodeList
   LOCAL pChildNodes2 AS IXMLDOMNodeList
   LOCAL pItem1 AS IXMLDOMNode
   LOCAL pItem2 AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pNewElem = pXmlDoc.createElement("PAGES")
         pChildNodes = pRootNode.childNodes
         pItem1 = pChildNodes.item(1)
         pChildNodes2 = pItem1.childNodes
         pItem2 = pChildNodes2.item(0)
         pItem1.replaceChild pNewElem, pItem2
         AfxShowMsg pItem1.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
