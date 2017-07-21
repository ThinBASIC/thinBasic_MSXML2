' ========================================================================================
' Demonstrates the use of namespaceURI property (XMLSchemaCache/IXMLSchemaCollection).
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
   LOCAL pSchemaCollection AS IXMLDOMSchemaCollection
   LOCAL strNamespaceURI AS STRING
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.load "books2.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pSchemaCollection = pXmlDoc.namespaces
         FOR i = 0 TO pSchemaCollection.length - 1
            AfxShowMsg pSchemaCollection.namespaceURI(i)
         NEXT
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
