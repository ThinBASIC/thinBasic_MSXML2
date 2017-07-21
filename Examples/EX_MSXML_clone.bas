' ========================================================================================
' Demonstrates the use of the clone method.
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
   LOCAL pXPath AS IXMLDOMSelection
   LOCAL pXPath2 AS IXMLDOMSelection
   LOCAL pTemp1 AS IXMLDOMNode
   LOCAL pTemp2 AS IXMLDOMNode

   pXmlDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXmlDoc) THEN EXIT FUNCTION

   TRY
      pXmLDoc.loadXML "<root><elem1>Hello</elem1><elem2>World!</elem2></root>"
      IF pXmlDoc.parseError.errorCode THEN
         AfxShowMsg "You have error " & pXmlDoc.parseError.reason
      ELSE
         ' Create an XMLDOMSelection object from selected nodes.
         pXPath = pXmlDoc.selectNodes("root/elem1")
         ' Cache the XPath expression and context.
         pXPath.expr = "root/elem1"
         pXPath.putref_context = pXmlDoc
         ' Clone the XMLDOMSelection object.
         pXPath2 = pXPath.clone
         pTemp1 = pXpath.peekNode    ' temp1 == <elem1/>
         AfxShowMsg pTemp1.xml
         pTemp2 = pXPath2.peekNode   ' temp2 == <elem1/>
         AfxShowMsg pTemp2.xml
         ' Note that position and context are maintained.
      END IF
   CATCH
      AfxShowMsg OleGetErrorInfo(OBJRESULT)
   END TRY

END FUNCTION
' ========================================================================================
