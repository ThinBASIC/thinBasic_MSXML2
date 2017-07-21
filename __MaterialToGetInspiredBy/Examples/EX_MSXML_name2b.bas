' ========================================================================================
' Demonstrates the use of name property (IXMLDOMDocumentType).
' Sets the ProhibitDTD to false to allow the inclusion of a DTD in the XML DOM document.
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
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pDocType AS IXMLDOMDocumentType

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISFALSE ISOBJECT(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmLDoc.setProperty "ProhibitDTD", %VARIANT_FALSE
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books1.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDocType = pXmlDoc.doctype
         AfxShowMsg pDocType.name
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
