  '----------------------------------------------------------------------------------------------------------
  '----------------------------------------------------------------------------
  ' Class
  '----------------------------------------------------------------------------
  Class cMsxml2_DOMDocument

    '----------------------------------------------------------------------------
    '---Class Private variables
    '----------------------------------------------------------------------------
    'Instance lString   As String
    Instance oDOMDoc    As Dispatch

    Instance vResult    As Variant
    Instance lResult    As Long

    '----------------------------------------------------------------------------
    '---Define here any Class Method (if needed)
    '---Class methods are visible only from inside the class
    '---so they can be called only from inside other methods.
    '----------------------------------------------------------------------------
    Class Method Create()
      ' Do initialization
      Set oDOMDoc = newcom $PROGID_MSXML2_DOMDocument600'"Msxml2.DOMDocument.6.0"

      If Not IsObject(oDOMDoc) Then
        thinBasic_RunTimeError(%ERR__MODULE_SPECIFIC, "It was not possible to create a DOMDocument object (" & $PROGID_MSXML2_DOMDocument600 & ")")
      End If
 
    End Method

    Class Method Destroy()
      '---Do cleanup
      oDOMDoc = Nothing
    End Method


    '----------------------------------------------------------------------------
    ' We need to define an interface that will contain public methods
    '----------------------------------------------------------------------------
    Interface iMsxml2_DOMDocument


      Inherit IUnknown


      '----------------------------------------------------------------------------
      ' GetClassObject
      '----------------------------------------------------------------------------
      Method Msxml2_DOMDocument_GetClassObject() As Dword
        Method = VarPtr(oDOMDoc)
      End Method

      '----------------------------------------------------------------------------
      '
      '----------------------------------------------------------------------------
      Method LoadXML(ByVal sXml As String ) As Long
        Local vXml As variant

        vXml = sXml

        Object Call oDOMDoc.LoadXml(vXML) To vResult

      End Method

      '----------------------------------------------------------------------------
      '
      '----------------------------------------------------------------------------
      Method GetElementsByTagName(ByVal sTag As String) As String
        Local vTag    As variant
        Local ofields As Dispatch
        Local sResult As String
        Local vResult As Variant
        Local nItems  As Long
        Local nItem   As Long
        Local v1      As variant

        Set vTag = sTag
        Object Call oDOMDoc.GetElementsByTagName(vTag) To vResult

        'Object Call oDOMDoc.getElementsByTagName(v1) To vresult
        Set oFields = vResult
        sResult = Variant$(vResult)
        ''
        Object Get oFields.length To vResult
        nItems = Variant#(vresult)
        '
        For nItem = 0 To nItems -1
            v1 = nItem
            Object Get oFields.Item(v1).Text To vResult
            If VariantVT(vResult) <> 0 Then
              If nItem = 0 Then
                sResult += Variant$(vResult)
              Else
                sResult += $CrLf & Variant$(vResult)
              End If
            End If
        Next
        oFields = Nothing
        Method = sResult

      End Method



    End Interface

  End Class
  '----------------------------------------------------------------------------
  '----------------------------------------------------------------------------




  '--------------------------------------------------------------------------------------
  Function Msxml2_DOMDocument_ClassHandling ( _
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

    Local sXml         As String
    Local sTagName     As String


    '---As usual we map our instance (whatever it will be) to a pObject pointer
    ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject

    '---We need to parse next token and we need to use special fucntion: thinBasic_GetUnknownToken
    '---This function is special because it instruct parser (thinCore) to not check for token validity
    '---but just parse something and return whetever it will find in terms of token string

    '   ATTENTION: thinBasic_GetUnknownToken always return UCASE tokens in order to simplify SELECT/CASE
    sName = thinBasic_GetUnknownToken

    '---Now we have to check parsed token using token content
    Select Case Const$ sName      '<<<---IMPORTANT: use CONST$ to speed up Select Case operations

      '-------------------------------------------------------
      '---Methods
      '---------------
      Case "LOADXML"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnNumber
        '---parse Params
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sXML
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValNum = Left$(@RetValStr, nChars)
          Else
            @RetValNum = oObject(1&).LoadXml(sXml)
          End If
        End If

      Case "LEN"
        Function = %thinBasic_ReturnNumber
        @RetValNum = Len(@RetValStr)

      Case "GETELEMENTSBYTAGNAME"
        '---We need to tell thinCore what type of info we are returning assigning to return function value the type
        '---Than assign to @RetValNum or @RetValStr the return value
        Function = %thinBasic_ReturnString
        pp = thinBasic_CheckOpenParens_Optional
        thinBasic_ParseString sTagName
        If pp Then thinBasic_CheckCloseParens_Mandatory
        '---If all went fine ...
        If thinBasic_ErrorFree Then
          If lRecursive Then
            '@RetValStr = Mid$(@RetValStr, eStart, eLen)
          Else
            @RetValStr = oObject(1&).GetElementsByTagName(sTagName)
          End If
        End If


'      '-------------------------------------------------------
'      '---Properties
'      '---------------
'      Case "VALUE"  '---Support GET and SET. Programmer is free to support whatever event type
'
'        '---In all "value" cases we will return a string, so we can assign return type here
'        Function = %thinBasic_ReturnString
'
'        Select Case Long lAction
'          Case %Class_Action_Get
'            @RetValStr = oObject(1&).sGet
'
'          Case %Class_Action_Set
'            lEqualType = thinBasic_CheckEqualType_Mandatory
'            If thinBasic_ErrorFree Then
'              Select Case Long lEqualType
'
'                '---Assignment
'                Case  %Equal_EQ
'                      oObject(1&).sSet(thinBasic_Parse1String)
'
'                '---Assignment with concatenation
'                Case  %Equal_PLUSEQUAL, _
'                      %Equal_CONCEQUAL, _
'                      %Equal_POINTEQUAL
'
'                      oObject(1&).sSet(oObject(1&).sGet() & thinBasic_Parse1String)
'                Case Else
'                  '---Not supperted for strings
'                  'MsgBox FuncName$ & Str$(lEqualType)
'              End Select
'            End If
'
'        End Select

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
      Function = Msxml2_DOMDocument_ClassHandling ( _
                                                  pObject   , _
                                                  lAction   , _
                                                  sName     , _
                                                  RetValNum , _
                                                  RetValStr , _
                                                  %TRUE       _
                                                )
    End If

  End Function

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
  Sub Msxml2_DOMDocument_Create( ByVal pObject As Long )
  ' We want Syntax of constructor to be something like:
  '   object = New cString[(InitialStringData)]
  ' Example
  '   MyCString = New cString("Initial value")
  '--------------------------------------------------------------------------------------
    'Local sBuffer As String
    Local pp      As Byte

    '---Remember: programmer is responsible of syntax so wrapper method MUST parse syntax

    '---Optionally checks if there is a (
    pp = thinBasic_CheckOpenParens_Optional
    If pp Then thinBasic_CheckCloseParens_Mandatory
'      '---If yes, then parse string expression and )
'      thinBasic_ParseString sBuffer
'      thinBasic_CheckCloseParens_Mandatory
'    End If

    '---If all went fine ...
    If thinBasic_ErrorFree Then

      'Here is the magic !!!!!!!!!!!!
      'Use the magic REDIM ... AT and the trick is done
      'We are defining a new class interface variable at the pointer location allocated by thinBasic and passed to all methods
      ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject

      '---We instantiate a new class
      oObject(1&) = Class "cMsxml2_DOMDocument"
'MsgBox FuncName$ & ".1"
'      oObject(1&).Init
      'Set oHTTP = newcom "Msxml2.XMLHTTP"
'MsgBox FuncName$ & ".2"
    'else
      '---emit some error
    End If

  End Sub


  '--------------------------------------------------------------------------------------
   Sub Msxml2_DOMDocument_Destroy( ByVal pObject As Long )
  '--------------------------------------------------------------------------------------

    '---Well, there is nothing we need to do when script class is released
    '---but just in case you allocated some dynamic memory, you will be responsible to
    '---release it. And the best place is the destructor
    If pObject Then
      '---Use the magic REDIM ... AT and the trick is done
      ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject
      Set oObject(1&) = Nothing
'MsgBox FuncName$ & ".1"
    End If

  End Sub

  '----------------------------------------------------------------------------
  Function Msxml2_DOMDocument_GetClassObject(ByVal pObject As Long) As Dword
  '----------------------------------------------------------------------------
    '---If all went fine ...
    If thinBasic_ErrorFree Then
      ReDim oObject(1& To 1&) As iMsxml2_DOMDocument At pObject
      Function = oObject(1&).Msxml2_DOMDocument_GetClassObject()
    End If
  End Function
