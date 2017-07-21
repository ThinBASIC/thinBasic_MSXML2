' ========================================================================================
' Demonstrates the use of various methods of the IXMLDOMSchemaCollection interface.
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
   LOCAL pSchemaCache AS IXMLDOMSchemaCollection2
   LOCAL pSchemaNode AS IXMLDOMNode
   LOCAL nsTarget AS WSTRING

   ' Must use version 4.0
   pXmlDoc = NEWCOM "Msxml2.FreeThreadedDOMDocument.4.0"
   IF ISFALSE ISOBJECT(pXmlDoc) THEN EXIT FUNCTION
   pSchemaCache = NEWCOM "Msxml2.XMLSchemaCache.4.0"
   IF ISFALSE ISOBJECT(pSchemaCache) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   pXmlDoc.validateOnParse = %VARIANT_TRUE

   nsTarget = "myURI"
   pSchemaCache.add nsTarget, "books.xsd"
   pSchemaNode = pSchemaCache.get(nsTarget)
   ' Validate the collection
   pSchemaCache.validate
   IF OBJRESULT THEN AfxShowMsg "Validate failed"
   ' Get the namespaceURI
   AfxShowMsg pSchemaCache.namespaceURI(0)
   ' Get the length of the collection
   AfxShowMsg STR$(pSchemaCache.length)
   ' Remove the namesapce
   pSchemaCache.remove nsTarget
   ' Get again the length of the collection
   AfxShowMsg STR$(pSchemaCache.length)

END FUNCTION
' ========================================================================================
