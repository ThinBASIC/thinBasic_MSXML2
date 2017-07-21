' =========================================================================================
' Demonstrates the use of the createAttribute method.
' The following example creates a new attribute called ID and adds it to the attributes of
' the DOMDocument object.
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
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pNamedNodeMap AS IXMLDOMNamedNodeMap
   LOCAL pNewAtt AS IXMLDOMAttribute
   LOCAL i AS LONG

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmlDoc.load "books.xml"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         pRootNode = pXmlDoc.documentElement
         pNewAtt = pXmlDoc.createAttribute("ID")
         pNamedNodeMap = pRootNode.attributes
         pNamedNodeMap.setNamedItem(pNewAtt)
         FOR i = 0 TO pNamedNodeMap.length - 1
            AfxShowMsg pNamedNodeMap.item(i).xml
         NEXT
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' =========================================================================================
