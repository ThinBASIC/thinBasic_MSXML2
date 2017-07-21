' ========================================================================================
' Demonstrates the use of the value property.
' The following example gets the value of the first attribute of the document root and
' assigns it to the variable vValue.
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
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL pDOMAttr AS IXMLDOMAttribute
   LOCAL vValue AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pCurrNode = pRootNode.firstChild
         pNamedNodeMap = pXmlDoc.documentElement.firstChild.attributes
         pDOMNode = pNamedNodeMap.item(0)
         pDOMAttr = pDOMNode
         vValue = pDOMAttr.value
         AfxShowMsg VARIANT$$(vValue)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
