#If 0
  =============================================================================
   Program Name:
   Author      :
   Date        :
   Version     :
   Description :
  =============================================================================
  'COPYRIGHT AND PERMISSION NOTICE
  '============================================================================

  =============================================================================
#EndIf

  #Compile Dll
  #Register None
  #Dim All

  #Resource "thinBasic_MSXML2.PBR"

  #Include Once "WIN32API.INC"
  #Include Once "..\thinCore.INC"
  
  #Include Once ".\MSXML2.INC"
  #Include Once ".\thinBasic_MSXML2_Msxml2_ServerXMLHTTP.Inc"
  #Include Once ".\thinBasic_MSXML2_Msxml2_DOMDocument.Inc"



  '----------------------------------------------------------------------------
  Function LoadLocalSymbols Alias "LoadLocalSymbols" (Optional ByVal sPath As String) Export As Long
  ' This function is automatically called by thinCore whenever this DLL is loaded.
  ' This function MUST be present in every external DLL you want to use with thinBasic
  ' Use this function to initialize every variable you need and for loading the
  ' new symbol (read Keyword) you have created.
  '----------------------------------------------------------------------------

    Local RetCode                 As Long
    Local pClass_Msxml2_XMLHTTP   As Long
    Local pClass_Msxml2_XMLDOM    As Long

    '---------------------------------------------------------------------------
    ' There are two methods to create a thinBasic Module Class
    ' Method 1: each method or property must be declare separately
    ' Method 2: only one function (a a class function) will be used
    '---------------------------------------------------------------------------

'    '---------------------------------------------------------------------------
'    ' Method 1: Declare class WITHOUT any function
'    '           Declare methods And properties And Left All the job To thinCore
'    '--------------------------------------------------------------------------- 
'      '---Declare a class WITHOUT a class function
'      pClass_Msxml2_XMLHTTP = thinBasic_Class_Add("Msxml2_XMLHTTP", 0)
'  
'      '---If class was created, define all methods and properties, each connected to a CODEPTR module function/sub
'      If pClass_Msxml2_XMLHTTP Then
'  
'        ' -- Constructor wrapper function needs to be linked in as _Create
'        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "_Create"       , %thinBasic_ReturnNone    , CodePtr(Msxml2_ServerXMLHTTP_Create          ))
'        ' -- Destructor wrapper function needs to be linked in as _Create
'        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "_Destroy"      , %thinBasic_ReturnNone    , CodePtr(Msxml2_ServerXMLHTTP_Destroy         ))
'  
'        ' -- Common methods can take any name
'        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "Open"            , %thinBasic_ReturnString  , CodePtr(Msxml2_XMLHTTP_Open              ))
'        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "SetRequestHeader", %thinBasic_ReturnString  , CodePtr(Msxml2_XMLHTTP_SetRequestHeader  ))
'        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "Send"            , %thinBasic_ReturnString  , CodePtr(Msxml2_XMLHTTP_Send              ))
'        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "Responsetext"    , %thinBasic_ReturnString  , CodePtr(Msxml2_XMLHTTP_Responsetext      ))
'  
'  
''        RetCode = thinBasic_Class_AddProperty (pClass_Msxml2_XMLHTTP, "Value"         , %thinBasic_ReturnString  , CodePtr(CString_Property_Value  ))
'        
'      End If
'    '---------------------------------------------------------------------------


    '---------------------------------------------------------------------------
    ' Configure Class: ServerXMLHTTP
    '---------------------------------------------------------------------------           
      '---Declare a class WITH a class function
      pClass_Msxml2_XMLHTTP = thinBasic_Class_Add("ServerXMLHTTPRequest", CodePtr(Msxml2_ServerXMLHTTP_ClassHandling))

'MsgBox str$(pClass_Msxml2_XMLHTTP)  
      '---If class was created, we just need to mandatory define constructor and destructor
      If pClass_Msxml2_XMLHTTP Then
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "_Create"       , %thinBasic_ReturnNone    , CodePtr(Msxml2_ServerXMLHTTP_Create          ))
        ' -- Destructor wrapper function needs to be linked in as _Destroy
        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLHTTP, "_Destroy"      , %thinBasic_ReturnNone    , CodePtr(Msxml2_ServerXMLHTTP_Destroy         ))
  
      End If
    '---------------------------------------------------------------------------
    
    'https://msdn.microsoft.com/en-us/library/ms753800(v=vs.85).aspx
    thinBasic_AddEquate  "%ServerXMLHTTP_UNINITIALIZED"                 , "", 0                    
    thinBasic_AddEquate  "%ServerXMLHTTP_LOADING"                       , "", 1                    
    thinBasic_AddEquate  "%ServerXMLHTTP_LOADED"                        , "", 2                    
    thinBasic_AddEquate  "%ServerXMLHTTP_INTERACTIVE"                   , "", 3                    
    thinBasic_AddEquate  "%ServerXMLHTTP_COMPLETED"                     , "", 4                    


    
    '---------------------------------------------------------------------------
    ' Configure Class: DOMDocument
    '---------------------------------------------------------------------------           
      '---Declare a class WITH a class function
      pClass_Msxml2_XMLDOM = thinBasic_Class_Add("DOMDocument", CodePtr(Msxml2_DOMDocument_ClassHandling))
  
      '---If class was created, we just need to mandatory define constructor and destructor
      If pClass_Msxml2_XMLDOM Then
        ' -- Constructor wrapper function needs to be linked in as _Create
        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLDOM, "_Create"       , %thinBasic_ReturnNone    , CodePtr(Msxml2_DOMDocument_Create          ))
        ' -- Destructor wrapper function needs to be linked in as _Destroy
        RetCode = thinBasic_Class_AddMethod   (pClass_Msxml2_XMLDOM, "_Destroy"      , %thinBasic_ReturnNone    , CodePtr(Msxml2_DOMDocument_Destroy         ))
  
      End If
    '---------------------------------------------------------------------------
    
    Function = 0&
  End Function

  '----------------------------------------------------------------------------
  Function UnLoadLocalSymbols Alias "UnLoadLocalSymbols" () Export As Long
  ' This function is automatically called by thinCore whenever this DLL is unloaded.
  ' This function CAN be present but it is not necessary.
  ' Use this function to perform uninitialize process, if needed.
  '----------------------------------------------------------------------------


    Function = 0&
  End Function


  Function LibMain Alias "LibMain" (ByVal hInstance   As Long, _
                                    ByVal fwdReason   As Long, _
                                    ByVal lpvReserved As Long) Export As Long
    Select Case fwdReason
      Case %DLL_PROCESS_ATTACH

        Function = 1
        Exit Function
      Case %DLL_PROCESS_DETACH

        Function = 1
        Exit Function
      Case %DLL_THREAD_ATTACH

        Function = 1
        Exit Function
      Case %DLL_THREAD_DETACH

        Function = 1
        Exit Function
    End Select

  End Function
