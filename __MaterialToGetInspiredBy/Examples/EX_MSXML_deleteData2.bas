' ========================================================================================
' Demonstrates the use of the deleteData method.
' The following example creates an IXMLDOMComment object and uses the deleteData method to
' delete three characters of data starting after the third character in data for the
' IXMLDOMComment object.
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
      pXmlDOc.loadXML("<root><child/></root>")
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pComment = pXmlDoc.createComment("123456789")
         pComment.deleteData 3, 3
         AfxShowMsg pComment.xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
