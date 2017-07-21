' ========================================================================================
' Demonstrates the use of the appendData method.
' The following example creates an IXMLDOMComment object and uses the appendData method to
' add text to the string.
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

   TRY
      pXmlDoc.async = %VARIANT_FALSE
      pXmlDoc.loadXML "<root></root>"
      pComment = pXmlDoc.createComment("Hello...")
      pComment.appendData " ... World!"
      AfxShowMsg pComment.data
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
