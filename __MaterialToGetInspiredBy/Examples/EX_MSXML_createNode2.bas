' =========================================================================================
' Demonstrates the use of the createNode method.
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pNewNode AS IXMLDOMNode
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmLDoc.loadXml "<root><child/></root>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNewNode = pXmlDoc.createNode(%NODE_ATTRIBUTE, "Sci-Fi", "")
         pNamedNodeMap = pXmlDoc.documentElement.childNodes.item(0).Attributes
         pNamedNodeMap.setNamedItem pNewNode
         AfxShowMsg pXmlDoc.documentElement.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
