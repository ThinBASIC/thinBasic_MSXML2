' ========================================================================================
' Demonstrates the use of the publicId property.
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
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pEntity AS IXMLDOMEntity
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pDocType AS IXMLDOMDocumentType
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL vPublicId AS VARIANT

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.setProperty "ProhibitDTD", %VARIANT_FALSE
      pXmLDoc.load("doment1.xml")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDocType = pXmlDoc.doctype
         pNamedNodeMap = pDocType.entities
         pDOMNode = pNamedNodeMap.nextNode
         pEntity = pDOMNode
         vPublicID = pEntity.publicID
         AfxShowMsg VARIANT$$(vPublicId)
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
