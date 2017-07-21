' =========================================================================================
' Demonstrates the use of the length property (IXMLDOMCharacterData).
' The following example creates an IXMLDOMComment object and then assigns the length
' property to a variable.
' =========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' =========================================================================================
' Main
' =========================================================================================
FUNCTION PBMAIN

   LOCAL pXmlDoc AS IXMLDOMDocument
   LOCAL pComment AS IXMLDOMComment
   LOCAL lValue AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pComment = pXmLDoc.createComment("Hello World!")
      lValue = pComment.length
      AfxShowMsg STR$(lValue)
   END IF

END FUNCTION
' =========================================================================================
