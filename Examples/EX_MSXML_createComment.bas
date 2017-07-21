' ========================================================================================
' Demonstrates the use of createComment and appendChild.
' The following example creates a new comment node and appends it as the last child node
' of the DOMDocument object.
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
   LOCAL pParseError AS IXMLDOMParseError
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pComment AS IXMLDOMComment

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pComment = pXmlDoc.createComment("Hello World!")
         pRootNode.appendChild pComment
         AfxShowMsg pRootNode.Xml
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
