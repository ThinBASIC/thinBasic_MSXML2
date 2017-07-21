' ========================================================================================
' Demonstrates the use of the ms:type-namespace-uri XPath Extension function.
' The following example uses an XSLT template rule to select from books.xml all the
' elements and to output the elements data types and the namesapce URI as defined in
' books.xsd.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pxsd AS IXMLDOMSchemaCollection
   LOCAL pxml AS IXMLDOMDocument2
   LOCAL pxsl AS IXMLDOMDocument2

   pxsd = NEWCOM "Msxml2.XMLSchemaCache.6.0"
   IF ISNOTHING(pxsd) THEN EXIT FUNCTION
   pxml = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pxml) THEN EXIT FUNCTION
   pxsl = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pxsl) THEN EXIT FUNCTION

   ' namespace uri ("urn:books") must be declared as one of the namespace
   ' declarations in the "books.xml" that is an instance of "books.xsd"
   pxsd.add "urn:books", "books4.xsd"

   pxml.putref_schemas = pxsd
   pxml.setProperty "SelectionLanguage", "XPath'"
   pxml.setProperty "SelectionNamespaces", "xmlns:ms='urn:schemas-microsoft-com:xslt'"

   pxml.async = %FALSE
   pxml.validateOnParse = %TRUE
   pxml.load "books4.xml"

   pxsl.async = %FALSE
   pxsl.load "books4.xslt"
   AfxShowMsg pxml.transformNode(pxsl)

END FUNCTION
' ========================================================================================
