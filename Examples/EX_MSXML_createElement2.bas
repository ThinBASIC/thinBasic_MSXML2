' ========================================================================================
' Demonstrates the use of the createElement method.
' The following example creates an element called NEW and appends it to an IXMLDOMNode
' object. It then sets the text value of the element to 123.
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

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML("<root><child/></root>")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pNewElem = pXmlDoc.createElement("NEW")
         pRootNode.childnodes.item(0).appendChild(pNewElem)
         pRootNode.childnodes.item(0).lastChild.text = "123"
         AfxShowMsg pRootNode.childNodes.item(0).xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
