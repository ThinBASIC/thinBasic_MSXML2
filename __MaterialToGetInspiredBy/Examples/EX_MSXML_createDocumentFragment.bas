' ========================================================================================
' Demonstrates the use of createDocumentFragment.
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
   LOCAL pDocFragment AS IXMLDOMDocumentFragment
   LOCAL pRootNode AS IXMLDOMElement

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML("<root/>")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDocFragment = pXmlDoc.createDocumentFragment
         pDocFragment.AppendChild pXmlDoc.createElement("node1")
         pDocFragment.AppendChild pXmlDoc.createElement("node2")
         pDocFragment.AppendChild pXmlDoc.createElement("node3")
         AfxShowMsg pDocFragment.xml
         ' PB doesn't allow...
         ' pXmlDoc.documentElement.appendChild pDocFragment
         ' .. so we have to use an intermediate step...
         pRootNode = pXmlDoc.documentElement
         pRootNode.appendChild pDocFragment
         AfxShowMsg pXmlDoc.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
