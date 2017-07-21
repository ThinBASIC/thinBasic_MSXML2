' ========================================================================================
' Demonstrates the use of name property (IXMLDOMAttribute).
' The following example creates an IXMLDOMAttribute object from an attribute of the first
' child of the root, and then displays the value of its name.
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
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pDOMAttr AS IXMLDOMAttribute

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pNamedNodeMap = pRootNode.firstChild.attributes
         pDOMNode = pNamedNodeMap.item(0)
         pDOMAttr = pDOMNode
         AfxShowMsg pDOMAttr.name
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
