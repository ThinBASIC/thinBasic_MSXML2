' ========================================================================================
' Demonstrates the use of the get method.
' Notes: This method is deprecated in MSXML 6.0, where it throws a Not Implemented
' exception. Instead of using this method in MSXML 6.0, use the getSchema Method.
' MSXML 6.0 has removed support for XDR schemas, whereas XDR is supported in MSXML 3.0
' and MSXML 4.0. If this method is called with an XDR schema, the call will fail.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pSchemaCache AS IXMLDOMSchemaCollection
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL nsTarget AS STRING

   ' Must use version 3.0 or 4.0
   pSchemaCache = NEWCOM "Msxml2.XMLSchemaCache.4.0"
   IF ISNOTHING(pSchemaCache) THEN EXIT FUNCTION

   TRY
      nsTarget = ""
      pSchemaCache.add nsTarget, "rootChild.xdr"
      pDOMNode = pSchemaCache.get(nsTarget)
      AfxShowMsg pDOMNode.xml
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
