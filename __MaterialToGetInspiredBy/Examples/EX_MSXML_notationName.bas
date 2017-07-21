' ========================================================================================
' Demonstrates the use of the notationName property.
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
   LOCAL pEntity AS IXMLDOMEntity
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pDocType AS IXMLDOMDocumentType
   LOCAL pDomNode AS IXMLDOMNode

   ' DTDs are disabled by default in version 6.0
   pXmlDoc = NEWCOM "Msxml2.DOMDocument.4.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "doment1.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDocType = pXmlDoc.doctype
         pNamedNodeMap = pDocType.entities
         pDOMNode = pNamedNodeMap.nextNode
         pEntity = pDOMNode
         AfxShowMsg pEntity.notationName
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
