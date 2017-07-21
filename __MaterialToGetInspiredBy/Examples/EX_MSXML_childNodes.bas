' ========================================================================================
' Demonstrates the use of the childNodes property.
' The following example uses the childNodes property (collection) to return an
' IXMLDOMNodeList, and then iterates through the collection, displaying the value of each
' item's xml property.
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
   LOCAL pNodeList AS IXMLDOMNodeList
   LOCAL pCurrNode AS IXMLDOMNode
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
         pNodeList = pRootNode.childNodes
         FOR i = 0 TO pNodeList.length - 1
            pCurrNode = pNodeList.item(i)
            AfxShowMsg pCurrNode.xml
            pCurrNode = NOTHING
         NEXT
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
