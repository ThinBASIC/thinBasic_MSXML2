' ========================================================================================
' Demonstrates the use of the getSchema method.
' The following example shows the getSchema method being used to return a schema object.
' When used with the sample schema file (doc.xsd) file above, the examples in this topic
' return the namespace URI for the schema:
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

   LOCAL pSchemaCache AS IXMLDOMSchemaCollection2
   LOCAL pSchema AS ISchema
   LOCAL nsTarget AS STRING

   pSchemaCache = NEWCOM "Msxml2.XMLSchemaCache.6.0"
   IF ISFALSE ISOBJECT(pSchemaCache) THEN EXIT FUNCTION

   nsTarget = "http://xsdtesting"
   pSchemaCache.add nsTarget, "doc.xsd"
   pSchema = pSchemaCache.getSchema(nsTarget)
   AfxShowMsg pSchema.targetNamespace

END FUNCTION
' ========================================================================================
