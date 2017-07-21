' ========================================================================================
' Demonstrates the use of length property (XMLSchemaCache/IXMLSchemaCollection).
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
   LOCAL pSchemaCollection AS IXMLDOMSchemaCollection
   LOCAL strNamespaceURI AS STRING

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmLDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books2.xml") THEN
      ' Get the namespace of the root element.
      pRootNode = pXmlDoc.documentElement
      strNamespaceURI = pRootNode.namespaceURI
      pSchemaCollection = pXmlDoc.namespaces
      AfxShowMsg STR$(pSchemaCollection.length)
   END IF

END FUNCTION
' ========================================================================================
