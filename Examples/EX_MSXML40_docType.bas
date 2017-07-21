' ========================================================================================
' Demonstrates the use of the docType property.
' The following example creates an IXMLDOMDocumentType object, and then displays the name
' property of the object.
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
   LOCAL pDocType AS IXMLDOMDocumentType

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.4.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load("books1.xml")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pDocType = pXmlDoc.doctype
         IF ISNOTHING(pDocType) THEN
            AfxShowMsg "There is no document type node"
         ELSE
            AfxShowMsg pDocType.name
         END IF
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY


END FUNCTION
' ========================================================================================
