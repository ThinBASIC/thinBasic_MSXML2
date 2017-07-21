' ========================================================================================
' Demonstrates the use of the specified property.
' The following example creates an IXMLDOMNode from the specified item in an
' IXMLDOMNamedNodeMap. It then displays whether or not the attribute was specified in the
' element, rather than in a DTD or schema.
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
   LOCAL pCurrNode AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pMyNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pCurrNode = pXmlDoc.documentElement.childNodes.item(0)
         pNamedNodeMap = pCurrNode.attributes
         pMyNode = pNamedNodeMap.item(0)
         AfxShowMsg STR$(pMyNode.specified)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
