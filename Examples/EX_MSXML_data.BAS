' ========================================================================================
' Demonstrates the use of the data (IXMLDOMCharacterData) property.
' The following script example walks the document tree and checks for comment node types.
' If one is found, it displays its contents with the data property.
' ========================================================================================

#DIM ALL
#COMPILE EXE
#INCLUDE ONCE "msxml.inc"
#INCLUDE ONCE "ole2utils.inc"

' ========================================================================================
' Main
' ========================================================================================
FUNCTION PBMAIN

   LOCAL pXMLDoc AS IXMLDOMDocument2
   LOCAL pRootNode AS IXMLDOMElement
   LOCAL pDOMNode AS IXMLDOMNode
   LOCAL pComment AS IXMLDOMComment
   LOCAL i AS LONG

   pXMLDoc = NEWCOM "Msxml2.DOMDocument.6.0"
   IF ISNOTHING(pXMLDoc) THEN EXIT FUNCTION

   pXMLDoc.async = %VARIANT_FALSE
   pXMLDoc.loadXML "<root><!-- Hello --></root>"
   IF pXmlDoc.parseError.errorCode THEN
      AfxShowMsg "You have error " & pXmlDoc.parseError.reason
   ELSE
      pRootNode = pXMLDoc.documentElement
      FOR i = 0 TO pRootNode.childNodes.length - 1
         IF pRootNode.childNodes.item(i).nodeType = %NODE_COMMENT THEN
            ' This doesn't work with PB because pRootNode.childNodes.item(i)
            ' returns a pointer to yhe IXMLDOMNode interface, not a pointer
            ' to the IXMLDOMComment interface.
            ' pComment = pRootNode.childNodes.item(i)
            ' MSGBOX pComment.Data
            ' Therefore, we need assign it to an object variable declared as
            ' IXMLDOMNode and then assign it to pComment, forcing an implicit
            ' call to QueryInterface to retrieve the correct object pointer.
            pDOMNode = pRootNode.childNodes.item(i)
            pComment = pDOMNode
            AfxShowMsg pComment.Data
         END IF
      NEXT
   END IF

END FUNCTION
' ========================================================================================
