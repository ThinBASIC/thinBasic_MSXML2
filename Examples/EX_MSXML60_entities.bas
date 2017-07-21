' ========================================================================================
' Demonstrates the use of the entities property.
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
   LOCAL pEntity AS IXMLDOMEntity
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pDOMNode AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.setProperty "ProhibitDTD", %VARIANT_FALSE
      pXmlDoc.load "doment1.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pNamedNodeMap = pXmlDoc.doctype.entities
         pDOMNode = pNamedNodeMap.nextNode
         pEntity = pDOMNode
         AfxShowMsg pEntity.notationName
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
