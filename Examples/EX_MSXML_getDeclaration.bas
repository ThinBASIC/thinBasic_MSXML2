' ========================================================================================
' Demonstrates the use of the getDeclaration method.
' Only works with version 4.0.
' The syntax IXMLDOMSchemaCollection2_getDeclaration(pSchemaCollection, pRootNode)
' is no longer supported and will return %E_NOTIMPL.
' When used with the sample XML (doc.xml) and schema file (doc.xsd) files, this example
' returns the namespace URI for the schema declared in doc.xml:
' http://xsdtesting
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
   LOCAL pSchemaCollection AS IXMLDOMSchemaCollection2
   LOCAL pSchemaItem AS ISchemaItem

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.4.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.validateOnParse = %VARIANT_FALSE
      pXmLDoc.setProperty "SelectionLanguage", "XPath"
      pXmlDoc.load "doc.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         ' Retrieve the namespace URI for schema used in the XML document.
         pRootNode = pXmlDoc.documentElement
         pSchemaCollection = pXmlDoc.namespaces
         pSchemaItem = pSchemaCollection.getDeclaration(pRootNode)
         AfxShowMsg pSchemaItem.namespaceURI
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
