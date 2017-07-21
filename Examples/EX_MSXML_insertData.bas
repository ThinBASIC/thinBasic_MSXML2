' ========================================================================================
' Demonstrates the use of the insertData method.
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
   LOCAL pComment AS IXMLDOMComment

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   pXmlDoc.async = %VARIANT_FALSE
   IF ISTRUE pXmlDoc.load("books.xml") THEN
      pComment = pXmlDoc.createComment("Hello...")
      pComment.insertData 6, "Beautiful"
      AfxShowMsg pComment.xml
   END IF

END FUNCTION
' ========================================================================================
