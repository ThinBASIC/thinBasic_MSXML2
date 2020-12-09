'  $DOMDOC_LoadXML                 = "LOADXML"
'  $DOMDOC_Load                    = "LOAD"
  $DOMDOCNODE_TEXT                    = "TEXT"
'  $DOMDOC_XML                     = "XML"
'  $DOMDOC_COUNT                   = "COUNT"
'  $DOMDOC_SetProperty             = "SETPROPERTY"
'  $DOMDOC_GetProperty             = "GETPROPERTY"
'
'  $DOMDOC_ParseError              = "PARSEERROR"
'  $DOMDOC_ParseError_ErrorCode    = "ERRORCODE"
'  $DOMDOC_ParseError_Reason       = "REASON"
'  $DOMDOC_ParseError_SrcText      = "SRCTEXT"
'  $DOMDOC_ParseError_Line         = "LINE"
'  $DOMDOC_ParseError_LinePos      = "LINEPOS"
'  $DOMDOC_ParseError_FilePos      = "FILEPOS"
'  $DOMDOC_ParseError_Url          = "URL"
'
'  $DOMDOC_GetElementsByTagName    = "GETELEMENTSBYTAGNAME"
'  $DOMDOC_SelectSingleNode        = "SELECTSINGLENODE"

  '----------------------------------------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  ' Class
  '----------------------------------------------------------------------------
  Class cMsxml2_DOMNode

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    Instance oDOMNode   As IXMLDOMNode


    '----------------------------------------------------------------------------
    '---Define here any Class Method (if needed)
    '---Class methods are visible only from inside the class
    '---so they can be called only from inside other methods.
    '----------------------------------------------------------------------------
    Class Method Create()
      'Set oDOMNode = newcom $IID_MSXML2_IXMLDOMNode'$PROGID_MSXML2_DOMDocument600
      'If Not IsObject(oDOMNode) Then
      '  thinBasic_RunTimeError(%ERR__MODULE_SPECIFIC, "It was not possible to create a DOMDocument_Node object (" & $PROGID_MSXML2_DOMDocument600 & ")")
      'End If
    End Method

    Class Method Destroy()
      oDOMNode = Nothing
    End Method

    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iMsxml2_DOMNode

      Inherit IUnknown

      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method Msxml2_DOMNode_GetClassObject() As Dword
MsgBox FuncName$
        Method = VarPtr(oDOMNode)
      End Method

      '----------------------------------------------------------------------------
      ' Creator Direct
      '----------------------------------------------------------------------------
      Method Msxml2_DOMNode_Create_Direct(ByVal ptrToDirectRecordSet As Dword) As Dword
MsgBox FuncName$ & Str$(ptrToDirectRecordSet)
        '---Here is the magic: we store at the memory address of DOMNode (that is a IXMLDOMNode COM pointer)
        '---The address passed by the script. The address is usually returned by another Class method or property
        '---Instructed to return an OBJPTR to an ADORecordset
        Poke Dword, VarPtr(oDOMNode), ptrToDirectRecordSet
      End Method
      
      Method Text() As String
        Method = oDOMNode.Text
      End Method
      
    End Interface

  End Class

  '--------------------------------------------------------------------------------------
  ' We need to create a constructor wrapper function
  ' Constructor function will be executed by thinBasic when the class will be instantiated
  ' in script using the NEW keyword
  '--------------------------------------------------------------------------------------
  ' IMPORTANT:
  '   ALL wrapper functions MUST HAVE a "BYVAL pObject AS LONG" parameter.
  '   pObject will contain a pointer to the object of our class
  '   This pointer is initialized by thinBasic engine to hold object reference
  '--------------------------------------------------------------------------------------
  Sub Msxml2_DOMNode_Create( ByVal pObject As Long )
    Local pp      As Byte

MsgBox FuncName$

    pp = thinBasic_CheckOpenParens_Optional
    If pp Then thinBasic_CheckCloseParens_Mandatory

    If thinBasic_ErrorFree Then
      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iMsxml2_DOMNode At pObject
      oObject(1&) = Class "cMsxml2_DOMNode"
    End If

  End Sub

  '----------------------------------------------------------------------------
  Sub Msxml2_DOMNode_Create_Direct(ByVal pObject As Dword, ByVal pDirectClass As Dword)
  '----------------------------------------------------------------------------
    Local pp          As Byte

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iMsxml2_DOMNode At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cMsxml2_DOMDocument_Node"
      '---We call sSet in order to assign initial value
      oObject(1&).Msxml2_DOMNode_Create_Direct(pDirectClass)

    'else
      '---emit some error
    End If
  End Sub

  '--------------------------------------------------------------------------------------
   Sub Msxml2_DOMNode_Destroy( ByVal pObject As Long )
  '--------------------------------------------------------------------------------------
    If pObject Then
      ReDim oObject(1& To 1&) As iMsxml2_DOMNode At pObject
      Set oObject(1&) = Nothing
    End If

  End Sub

  '----------------------------------------------------------------------------
  Function Msxml2_DOMNode_GetClassObject(ByVal pObject As Long) As Dword
  '----------------------------------------------------------------------------

    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iMsxml2_DOMNode At pObject
      Function = oObject(1&).Msxml2_DOMNode_GetClassObject()
    End If
  End Function
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------


  '--------------------------------------------------------------------------------------
  Function Msxml2_DOMNode_ClassHandling ( _
                                                        ByVal pObject       As Long       , _
                                                        ByVal lAction       As Long       , _
                                                        ByVal sPreviousName As String     , _
                                              Optional  ByVal RetValNum     As Ext Ptr    , _
                                                        ByVal RetValStr     As String Ptr , _
                                                        ByVal lRecursive    As Long       , _
                                                        ByVal sParam1       As String     , _
                                                        ByVal sParam2       As String     , _
                                                        ByVal eParam1       As Ext        , _
                                                        ByVal eParam2       As Ext          _
                                            ) As Long
  '--------------------------------------------------------------------------------------
  ' IMPORTANT: function must return main return type of called method/property
  '--------------------------------------------------------------------------------------
    Local pp            As Long
    Local sName         As String
    Local sName2        As String
    Local sName3        As String

    Local sXml          As String
    Local sTagName      As String
    Local vResult       As Variant

    'Local pNodeList     As IXMLDOMNodeList
    Local Counter       As Long
    Local sBuffer       As String
    
    Local lIdx          As Long
    
    Local sPropName     As String
    Local sPropValue    As String

    '---As usual we map our instance (whatever it will be) to a pObject pointer
    'ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject

    ReDim oObject(1& To 1&)   As iMsxml2_DOMNode At pObject
MsgBox "OK1"
    ReDim pDomNode(1& To 1&)  As IXMLDOMNode At oObject(1&).Msxml2_DOMNode_GetClassObject
MsgBox "OK2"

    '---We need to parse next token and we need to use special fucntion: thinBasic_GetUnknownToken
    '---This function is special because it instruct parser (thinCore) to not check for token validity
    '---but just parse something and return whetever it will find in terms of token string

    '   ATTENTION: thinBasic_GetUnknownToken always return UCASE tokens in order to simplify SELECT/CASE
    sName = thinBasic_GetUnknownToken

MsgBox "sName:" & sName

    '---Now we have to check parsed token using token content
    Select Case Const$ sName      '<<<---IMPORTANT: use CONST$ to speed up Select Case operations

      '-------------------------------------------------------
      '---Methods
      '---------------
      Case $DOMDOCNODE_TEXT
        Function = %thinBasic_ReturnString
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        If pp Then thinBasic_CheckCloseParens_Mandatory
        
        '---If all went fine ...
        If thinBasic_ErrorFree Then
MsgBox "OK3"
          If IsObject(pDomNode(1&)) Then
            If lRecursive Then
              '@RetValNum = Left$(@RetValStr, nChars)
            Else
MsgBox "OK4"
              '@RetValStr = pDomNode(1&).NodeName'Text
              @RetValStr = oObject(1&).Text
MsgBox "OK5"

            End If
          End If
        End If
MsgBox "OK9"

'      Case $DOMDOC_Load
'        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
'        '---Than assign to @RetValNum or @RetValStr the return value
'        Function = %thinBasic_ReturnNumber
'        '---parse Params
'        pp = thinBasic_CheckOpenParens_Optional
'        thinBasic_ParseString sXML
'        If pp Then thinBasic_CheckCloseParens_Mandatory
'        '---If all went fine ...
'        If thinBasic_ErrorFree Then
'          If lRecursive Then
'            '@RetValNum = Left$(@RetValStr, nChars)
'          Else
'            @RetValNum = pDomDocument(1&).Load(sXml)
'          End If
'        End If
'
'      Case $DOMDOC_TEXT
'        Function = %thinBasic_ReturnString
'        @RetValStr = pDomDocument(1&).Text
'
'      Case $DOMDOC_XML
'        Function = %thinBasic_ReturnString
'        @RetValStr = pDomDocument(1&).XML
'      
'      Case $DOMDOC_ParseError
'        '---Do net iteration
'        
'        Case $DOMDOC_ParseError_ErrorCode
'          Select Case Const$ sPreviousName
'            Case $DOMDOC_ParseError
'              Function = %thinBasic_ReturnNumber            
'              @RetValNum = pDomDocument(1&).ParseError.errorCode
'          End Select
'        Case $DOMDOC_ParseError_Reason
'          Select Case Const$ sPreviousName
'            Case $DOMDOC_ParseError
'              Function = %thinBasic_ReturnString
'              @RetValStr = pDomDocument(1&).ParseError.Reason
'          End Select
'        Case $DOMDOC_ParseError_SrcText
'          Select Case Const$ sPreviousName
'            Case $DOMDOC_ParseError
'              Function = %thinBasic_ReturnString
'              @RetValStr = pDomDocument(1&).ParseError.srcText
'          End Select
'        Case $DOMDOC_ParseError_Line
'          Select Case Const$ sPreviousName
'            Case $DOMDOC_ParseError
'              Function = %thinBasic_ReturnNumber
'              @RetValNum = pDomDocument(1&).ParseError.Line
'          End Select
'        Case $DOMDOC_ParseError_LinePos
'          Select Case Const$ sPreviousName
'            Case $DOMDOC_ParseError
'              Function = %thinBasic_ReturnNumber
'              @RetValNum = pDomDocument(1&).ParseError.LinePos
'          End Select
'        Case $DOMDOC_ParseError_FilePos
'          Select Case Const$ sPreviousName
'            Case $DOMDOC_ParseError
'              Function = %thinBasic_ReturnNumber
'              @RetValNum = pDomDocument(1&).ParseError.FilePos
'          End Select
'        Case $DOMDOC_ParseError_Url
'          Select Case Const$ sPreviousName
'            Case $DOMDOC_ParseError
'              Function = %thinBasic_ReturnString
'              @RetValStr = pDomDocument(1&).ParseError.Url
'          End Select
'
'      Case $DOMDOC_SetProperty
'        Function = %thinBasic_ReturnNumber
'        pp = thinBasic_CheckOpenParens_Optional
'          thinBasic_ParseString sPropName 
'          If thinBasic_CheckComma_Mandatory Then
'            thinBasic_ParseString sPropValue
'          End If
'        If pp Then thinBasic_CheckCloseParens_Mandatory
'        
'        If thinBasic_ErrorFree Then
'          If lRecursive Then
'          Else
'            pDomDocument(1&).SetProperty(sPropName, sPropValue)
'          End If
'        End If
'      Case $DOMDOC_GetProperty
'        Function = %thinBasic_ReturnString
'        pp = thinBasic_CheckOpenParens_Optional
'          thinBasic_ParseString sPropName 
'        If pp Then thinBasic_CheckCloseParens_Mandatory
'        
'        If thinBasic_ErrorFree Then
'          If lRecursive Then
'          Else
'            sBuffer = Variant$$(pDomDocument(1&).GetProperty(sPropName))
'          End If
'          @RetValStr = sBuffer
'        End If
'                
''      Case "LEN"
''        Function = %thinBasic_ReturnNumber
''        @RetValNum = Len(@RetValStr)
'
'      Case $DOMDOC_GetElementsByTagName
'        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
'        '---Than assign to @RetValNum or @RetValStr the return value
'        pp = thinBasic_CheckOpenParens_Optional
'        thinBasic_ParseString sTagName
'        If pp Then thinBasic_CheckCloseParens_Mandatory
'
'        If thinBasic_CheckOpenParens_Optional Then
'          thinBasic_ParseLong lIdx
'          thinBasic_CheckCloseParens_Mandatory
'        End If  
'
'        If thinBasic_CheckPoint_Optional Then
'          sName2 = thinBasic_GetUnknownToken
'        End If
'        
'        '---If all went fine ...
'        If thinBasic_ErrorFree Then
'          If lRecursive Then
'            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
'          Else 
''MsgBox "OK"
'            '@RetValNum = pDomDocument(1&).GetElementsByTagName(sTagName).Length
''MsgBox "OK1:"
'
'              pNodeList = pDomDocument(1&).GetElementsByTagName(sTagName)
'              'For Counter = 0 To pNodeList.length - 1
'              '  sBuffer += pNodeList.Item(Counter).nodeName + ":" & pNodeList.Item(Counter).Text & $CrLf
'              'Next
''MsgBox "OK2:"
'              If IsObject(pNodeList) Then
'                If pNodeList.Length > 0 Then
''MsgBox "OK3:"    
'                  lIdx = lIdx - 1
'                  If lIdx <= 0 Or lIdx > pNodeList.Length + 1 Then lIdx = 0
'                  
'                  Select Case Const$ sName2
'                    Case $DOMDOC_XML 
'                      Function = %thinBasic_ReturnString
'                      sBuffer = pNodeList.Item(lIdx).xml
'                      @RetValStr = sBuffer
'                    Case $DOMDOC_TEXT
'                      Function = %thinBasic_ReturnString
'                      sBuffer = pNodeList.Item(lIdx).Text
'                      @RetValStr = sBuffer
'                    Case $DOMDOC_COUNT
'                      Function = %thinBasic_ReturnNumber
'                      @RetValNum = pNodeList.Length
'                    Case Else 
'                      sBuffer = pNodeList.Item(lIdx).xml
'                  End Select
'                End If
''MsgBox "OK4:"
'              End If
'          End If
'        End If
'
'      Case $DOMDOC_SelectSingleNode
'        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
'        '---Than assign to @RetValNum or @RetValStr the return value
'        Function = %thinBasic_ReturnString
'        pp = thinBasic_CheckOpenParens_Optional
'        thinBasic_ParseString sTagName
'        If pp Then thinBasic_CheckCloseParens_Mandatory
'
'        If thinBasic_CheckPoint_Optional Then
'          sName2 = thinBasic_GetUnknownToken
'        End If
'        
'        '---If all went fine ...
'        If thinBasic_ErrorFree Then
'          If lRecursive Then
'            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
'          Else 
''MsgBox "OK1:" & sTagName
'            '@RetValNum = pDomDocument(1&).GetElementsByTagName(sTagName).Length
'
'            Local pNode     As IXMLDOMNode
'              pNode = pDomDocument(1&).SelectSingleNode(sTagName)
'              'For Counter = 0 To pNodeList.length - 1
'              '  sBuffer += pNodeList.Item(Counter).nodeName + ":" & pNodeList.Item(Counter).Text & $CrLf
'              'Next
'              If IsObject(pNode) Then
''MsgBox "OK2"
'                Select Case Const$ sName2
'                  Case $DOMDOC_XML 
'                    sBuffer = pNode.xml
'                  Case $DOMDOC_TEXT
'                    sBuffer = pNode.Text
'                  Case Else 
'                    sBuffer = pNode.xml
'                End Select
'              End If
'              @RetValStr = sBuffer
'          End If
'        End If
'
'
      Case Else
        thinBasic_RunTimeError(%ERR__CLASS_METHODPROPERTY_NOTFOUND)
    End Select

    '-------------------------------------------------------
    '---Compound methods
    '---------------
    '   In case we have compound methods (defined as point after a method or property
    '   call this function recursively.
    '---------------
    If thinBasic_CheckPoint_Optional Then
      Function = Msxml2_DOMNode_ClassHandling ( _
                                                  pObject   , _
                                                  lAction   , _
                                                  sName     , _
                                                  RetValNum , _
                                                  RetValStr , _
                                                  %TRUE       _
                                                )
    End If

  End Function

 