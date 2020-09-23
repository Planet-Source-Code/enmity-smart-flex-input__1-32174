VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlexInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**********************************************************************************************************************
'*
'*   Program Name: Smart Flex Input
'*   Description : filter, verify and correct input
'*   Version     : 0.5
'*   Copyright   : Smartsoft 2002
'*   Program Type: ActiveX DLL
'*   Language    : Visual Basic 6
'*   Reference   : (none)
'*   Functions   : 1. Filter input
'*                 2. verify input
'*                 3. (auto)correct input
'*                 4. I am thinking... :)
'*   Future      : mask edit
'*   DLL Name    : SmartFlexInput
'*   Class Name  : cFlexInput
'*
'*   ¡óFunction:FilterInput
'*         work:limit input, filter illegal chars,
'*              support number,text, date, time etc
'*         para:objInputBox               ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               intKeyAscii              ---- user input char
'*               strKeyCodeRange(optional)---- char range
'*               blnSmallNumber(optional) ---- allow char(".")
'*               intMaxLength(optional)   ---- maximum length
'*               blnShowMsg(optional)     ---- show help message
'*       return:[Integer]if input valid, return it, otherwise, return 0
'*
'*   ¡óFunction:ValidateNumber
'*         work:input verify(number)
'*         para:objInputBox              ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               curMinNum               ---- minimum number
'*               curMaxNum               ---- maximum number
'*               intMaxLength(optional)  ---- maximum length
'*               blnCanBeEmpty(optional) ---- allow empty input
'*               blnShowMsg(optional)    ---- show help message
'*       return:[enumErrorType]if valid input, return etDefault(0), otherwise, enumErrorType
'*
'*   ¡óFunction:ValidateDateTime
'*         work:input verify(date/time)
'*         para:objInputBox              ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               lngMinDateTime          ---- minimum number
'*               lngMaxDateTime          ---- maximum number
'*               udtType(optional)       ---- type(date/time)
'*               strFormat(optional)     ---- format
'*               intMaxLength(optional)  ---- maximum length
'*               blnCanBeEmpty(optional) ---- allow empty input
'*               blnShowMsg(optional)    ---- show help message
'*       return:[enumErrorType]if valid input, return etDefault(0), otherwise, enumErrorType
'*
'*   ¡óFunction:ValidateText
'*         work:input verify(text)
'*         para:objInputBox               ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               strForbiddenChars        ---- illegal chars
'*               strSplitChar(optional)   ---- split char
'*               strReplaceChar(optional) ---- char to be replaced
'*               blnAuthReplace(optional) ---- auto replace illegal chars
'*               intMaxLength(optional)   ---- maximum length
'*               blnCanBeEmpty(optional)  ---- allow empty input
'*               blnShowMsg(optional)     ---- show help message
'*       return:[enumErrorType]if valid input, return etDefault(0), otherwise, enumErrorType
'*
'*   limitation:1.could not filter the paste text from CTRL+V/Popup Menu
'*
'*   Idea              :Wilson Chan
'*   Design            :Wilson Chan
'*   Program           :Wilson Chan
'*   Last Modified by  :Wilson Chan
'*   Last Modified Date:2002/1/17
'*
'**********************************************************************************************************************

'Error Type
Public Enum enumErrorType
    etDefault = 0 'Valid
    etInvalid = 1 'Invalid
    etRange = 2 'input range
    etMaxLength = 3 'maximum length
    etEmpty = 4 'empty
    etModified = 5 'auto correct
    etUnknown = 99 'unknow error
End Enum

'input type
Public Enum enumInputType
    itNumeric = 0 'number
    itChar = 1 'char
    itDate = 2 'date
    itTime = 3 'time
    itID = 4 'id
End Enum

Public Enum enumDateTime
    dtDate = 0 'date
    dtTime = 1 'time
End Enum



Public Function FilterInput(ByRef objInputBox As Object, _
                            ByVal intKeyAscii As Integer, _
                            Optional ByVal udtInputType As enumInputType = itNumeric, _
                            Optional ByVal strKeyCodeRange As String = "0-9", _
                            Optional ByVal blnSmallNumber As Boolean = False, _
                            Optional ByVal intMaxLength As Integer = 0, _
                            Optional ByVal blnShowMsg As Boolean = False _
                            ) As Integer
Attribute FilterInput.VB_Description = "¹ýÂËÊäÈë£¬Ö§³ÖÊý×Ö¡¢ÎÄ±¾¡¢ÈÕÆÚ¡¢Ê±¼ä¡¢Éí·ÝÖ¤µÈ"
'************************************************************************
'*
'*   ¡óFunction:FilterInput
'*         work:limit input, filter illegal chars,
'*              support number,text, date, time etc
'*         para:objInputBox               ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               intKeyAscii              ---- user input char
'*               strKeyCodeRange(optional)---- char range
'*               blnSmallNumber(optional) ---- allow char(".")
'*               intMaxLength(optional)   ---- maximum length
'*               blnShowMsg(optional)     ---- show help message
'*       return:[Integer]if input valid, return it, otherwise, return 0
'*
'* notation:range char"-", split char"|"
'* sample  : "a-z|.|0-9|,|A-Z|~"
'*
'* Last Modified by:Unruled Boy @ 1/12/2002
'*
'*************************************************************************

    Dim o_strRet() As String
    Dim o_strText As String
    Dim o_strRange As String
    Dim o_intItems As Integer
    Dim o_intRet As Integer
    Dim o_intKeyAscii As Integer
    'Dim o_blnRet As Boolean
    Dim o_udtErrorType As enumErrorType
    
    o_udtErrorType = etDefault 'initial error:none
        
    With objInputBox
        Select Case TypeName(objInputBox)
                Case "TextBox", "RichTextBox"
                    'if there is a maximum length, use it
                    If .MaxLength > 0 Then '
                        If intMaxLength = 0 Then
                            intMaxLength = .MaxLength
                        Else
                            If intMaxLength > .MaxLength Then
                                intMaxLength = .MaxLength
                            Else
                                
                            End If
                        End If
                    Else
                        
                    End If
                Case Else
                
        End Select
        
        o_strText = .Text
    
    End With
        
    'if the current length exceeds the maximum length,
    'show hint
    If intMaxLength > 0 And Len(o_strText) >= intMaxLength Then
        
        Select Case intKeyAscii
                Case vbKeyDelete, vbKeyBack 'Delete & Backspace
                    o_intKeyAscii = intKeyAscii
                Case Else
                    Select Case TypeName(objInputBox)
                            Case "TextBox", "RichTextBox"
                                'if the user select a range of text
                                If objInputBox.SelLength > 0 Then
                                    o_intKeyAscii = intKeyAscii
                                Else
                                    Beep
                                    o_udtErrorType = etMaxLength
                                    o_intKeyAscii = 0
                                End If
                            Case Else
                                Beep
                                o_udtErrorType = etMaxLength
                                o_intKeyAscii = 0
                    End Select
        End Select
    Else
                
        'first, we process those speical keys
        Select Case intKeyAscii
                Case vbKeyDelete, vbKeyBack 'as above
                    o_intKeyAscii = intKeyAscii
                                        
                Case vbKeyDecimal, 190 'dot
                    If udtInputType = itNumeric And blnSmallNumber Then
                        o_intKeyAscii = intKeyAscii
                    Else
                        o_udtErrorType = etInvalid
                        o_intKeyAscii = 0
                    End If
                
                Case vbKeySubtract, vbKeyDivide, 45 'date time chars:-¡¢/
                    If udtInputType = itDate Then
                        o_intKeyAscii = intKeyAscii
                    ElseIf udtInputType = itNumeric Then
                        If Len(o_strText) = 0 Then
                            o_intKeyAscii = intKeyAscii
                        Else
                            o_udtErrorType = etInvalid
                            o_intKeyAscii = 0
                        End If
                    Else
                        o_udtErrorType = etInvalid
                        o_intKeyAscii = 0
                    End If
                    
                Case Asc(":") 'time char(:)
                    If udtInputType = itTime Then
                        o_intKeyAscii = intKeyAscii
                    Else
                        o_udtErrorType = etInvalid
                        o_intKeyAscii = 0
                    End If
                
                Case Else
                
                    'o_blnRet = False
                    
                    o_strRange = strKeyCodeRange
                    
                    If o_strRange <> vbNullString Then
                        
                        o_intItems = 0
                        
                        'trim the redundant split chars
                        Do While o_intItems < Len(o_strRange)
                            If Left(o_strRange, 1) = "|" Then
                                o_strRange = Right(o_strRange, _
                                                    Len(o_strRange) - 1)
                            ElseIf Right(o_strRange, 1) = "|" Then
                                o_strRange = Left(o_strRange, _
                                                    Len(o_strRange) - 1)
                            Else
                            End If
                            
                            o_intItems = o_intItems + 1
                        Loop
                        
                        'explain the combined para
                        o_strRet() = Split(o_strRange, "|")
                        For o_intItems = LBound(o_strRet) To UBound(o_strRet)
                            o_intRet = InStr(o_strRet(o_intItems), "-")
                            If o_intRet <> 0 Then
                                If intKeyAscii >= Asc(Left(o_strRet(o_intItems), o_intRet - 1)) _
                                        And intKeyAscii <= Asc(Right(o_strRet(o_intItems), Len(o_strRet(o_intItems)) - o_intRet)) Then
                                    o_intKeyAscii = intKeyAscii
                                    Exit For
                                End If
                            Else
                                If intKeyAscii = Val(o_strRet(o_intItems)) Then
                                    'o_blnRet = True
                                    o_intKeyAscii = intKeyAscii
                                    Exit For
                                Else
                                End If
                            End If
                        Next
                        
                        If o_intKeyAscii = 0 Then
                            o_udtErrorType = etRange
                        Else
                        End If
                        
                    Else
                        o_intKeyAscii = intKeyAscii
                    End If

        End Select
                    
    End If
    
    'if there is something wrong, give hint
    If o_udtErrorType <> etDefault Then
    
        If blnShowMsg Then
            Select Case o_udtErrorType
                    Case etRange
                        MsgBox "Your input is not in the [range]." & _
                                "Valid input range is:" & _
                                Replace(o_strRange, "|", ",") & ".", _
                                vbInformation
                    Case etMaxLength
                        MsgBox "Your input exceeds the [maximum length]." & _
                                "The maximum length is: " & _
                                CStr(intMaxLength), vbInformation
                    Case etInvalid
                        MsgBox "Your input is invalid, please check it." & _
                                "Hint:valid input are:" & _
                                Replace(o_strRange, "|", ",") & ".", _
                                vbInformation
                    Case Else
                        
            End Select
            
        Else
        
        End If
    Else
        
    End If
    
    FilterInput = o_intKeyAscii
    
    Exit Function
    
    
handleError:
    FilterInput = 0
    If blnShowMsg Then
        MsgBox "Error occurs when filtering the input:" & vbCrLf _
                 & Err.Description
    Else
    End If
    On Error GoTo 0
        
End Function


Public Function ValidateDateTime(ByRef objInputBox As Object, _
                              Optional ByVal datMinDateTime As Date = #1/1/1992#, _
                              Optional ByVal datMaxDateTime As Date = #1/1/2012#, _
                              Optional ByVal udtType As enumDateTime = dtDate, _
                              Optional ByVal strFormat As String = "mm/dd/yyyy", _
                              Optional ByVal intMaxLength As Integer = 10, _
                              Optional ByVal blnCanBeEmpty As Boolean = False, _
                              Optional ByVal blnShowMsg As Boolean = False _
                              ) As enumErrorType
Attribute ValidateDateTime.VB_Description = "Ð£ÑéÊäÈë£¨ÈÕÆÚÓëÊ±¼ä£©"
'********************************************
'*
'*   ¡óFunction:ValidateDateTime
'*         work:input verify(date/time)
'*         para:objInputBox              ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               lngMinDateTime          ---- minimum number
'*               lngMaxDateTime          ---- maximum number
'*               udtType(optional)       ---- type(date/time)
'*               strFormat(optional)     ---- format
'*               intMaxLength(optional)  ---- maximum length
'*               blnCanBeEmpty(optional) ---- allow empty input
'*               blnShowMsg(optional)    ---- show help message
'*       return:[enumErrorType]if valid input, return etDefault(0), otherwise, enumErrorType
'*
'* Last Modified by:Unruled Boy @ 1/11/2002
'*
'********************************************
    
    Dim o_strText As String
    Dim o_datText As Date
    Dim o_udtErrorType As enumErrorType
    
    o_udtErrorType = etDefault 'initial error:none
    
    With objInputBox
        
        Select Case TypeName(objInputBox)
                Case "TextBox", "RichTextBox"
                    'if there is a maximum length, use it
                    If .MaxLength > 0 Then
                        If intMaxLength = 0 Then
                            intMaxLength = .MaxLength
                        Else
                            If intMaxLength > .MaxLength Then
                                intMaxLength = .MaxLength
                            Else
                                
                            End If
                        End If
                    Else
                        
                    End If
                Case Else
                
        End Select
        
        o_strText = .Text
            
        'if the current length exceeds the maximum length,
        'show hint
        If intMaxLength > 0 And Len(o_strText) > intMaxLength Then
            
            o_udtErrorType = etMaxLength
            
        Else
            If o_strText <> vbNullString Then
                If IsDate(o_strText) Then
                    o_datText = CDate(o_strText)
                    If o_datText >= datMinDateTime And o_datText <= datMaxDateTime Then
                        o_udtErrorType = etDefault
                    Else
                        o_udtErrorType = etRange
                    End If
                Else
                    o_udtErrorType = etInvalid
                End If
            Else
                If Not blnCanBeEmpty Then
                    o_udtErrorType = etEmpty
                Else
                
                End If
            End If
        End If
                
        'if there is something wrong, give hint
        If o_udtErrorType <> etDefault Then
            On Error Resume Next
            .SetFocus
            On Error GoTo 0
        
            If blnShowMsg Then
                Select Case o_udtErrorType
                        Case etRange
                            MsgBox "The " & IIf(udtType = dtDate, "Date", "Time") & _
                                    " you input is invalid. Please check it." & _
                                    "Valid input range is: from" & _
                                    CStr(datMinDateTime) & " to " & datMaxDateTime & " ." _
                                    , vbInformation
                        Case etMaxLength
                            MsgBox "Your input exceeds the [maximum length]." & _
                                    "The maximum length is: " & _
                                    CStr(intMaxLength), vbInformation
                        Case etInvalid
                            MsgBox "The " & IIf(udtType = dtDate, "Date", "Time") & _
                                    " you input is invalid. Please check it." & vbCrLf & _
                                    "Input requires to be all numbers(including " & _
                                    IIf(udtType = dtDate, "/ and -", ":") & ")." & vbCrLf & _
                                    "Please check if your input contains non-number chars." _
                                    , vbInformation
                        Case etEmpty
                            MsgBox "Input could not be empty." & vbCrLf & _
                                    "You should input a valid" & _
                                    IIf(udtType = dtDate, "date", "time") _
                                    , vbInformation
                        Case Else
                            
                End Select
                
            Else
            
            End If
        Else
            
        End If
    
    End With
    
    ValidateDateTime = o_udtErrorType
    
    Exit Function
    
    
handleError:
    ValidateDateTime = etUnknown
    If blnShowMsg Then
        MsgBox "Error occurs when validating " & _
                IIf(udtType = dtDate, "date", "time") _
                & ":" & Err.Description
    Else
    End If
    On Error GoTo 0
    
End Function


Public Function ValidateNumber(ByRef objInputBox As Object, _
                              ByVal curMinNum As Currency, _
                              ByVal curMaxNum As Currency, _
                              Optional ByVal intMaxLength As Integer = 0, _
                              Optional ByVal blnCanBeEmpty As Boolean = False, _
                              Optional ByVal blnShowMsg As Boolean = False _
                              ) As enumErrorType
Attribute ValidateNumber.VB_Description = "Ð£ÑéÊäÈë£¨Êý×Ö£©£¬¿ÉÓÃ£¬²»ÍÆ¼ö"
'********************************************
'*
'*   ¡óFunction:ValidateNumber
'*         work:input verify(number)
'*         para:objInputBox              ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               curMinNum               ---- minimum number
'*               curMaxNum               ---- maximum number
'*               intMaxLength(optional)  ---- maximum length
'*               blnCanBeEmpty(optional) ---- allow empty input
'*               blnShowMsg(optional)    ---- show help message
'*       return:[enumErrorType]if valid input, return etDefault(0), otherwise, enumErrorType
'*
'* Last Modified by:Unruled Boy @ 1/11/2002
'*
'********************************************
    
On Error GoTo handleError
    
    Dim o_strText As String
    Dim o_lngNumber As Currency
    Dim o_udtErrorType As enumErrorType
    
    o_udtErrorType = etDefault 'initial error:none
            
    With objInputBox
        
        Select Case TypeName(objInputBox)
                Case "TextBox", "RichTextBox"
                    'if there is a maximum length, use it
                    If .MaxLength > 0 Then
                        If intMaxLength = 0 Then
                            intMaxLength = .MaxLength
                        Else
                            If intMaxLength > .MaxLength Then
                                intMaxLength = .MaxLength
                            Else
                                
                            End If
                        End If
                    Else
                        
                    End If
                Case Else
                
        End Select
        
        o_strText = .Text
        
        'if the current length exceeds the maximum length,
        'show hint
        If intMaxLength > 0 And Len(o_strText) > intMaxLength Then
            
            o_udtErrorType = etMaxLength
            
        Else
            If o_strText <> vbNullString Then
                If IsNumeric(o_strText) Then
                    o_lngNumber = CCur(o_strText)
                    If o_lngNumber < curMinNum Or o_lngNumber > curMaxNum Then
                        On Error Resume Next
                        .SetFocus
                        On Error GoTo 0
                        o_udtErrorType = etRange
                    Else
                        o_udtErrorType = etDefault
                    End If
                Else
                    o_udtErrorType = etInvalid
                End If
            Else
                If Not blnCanBeEmpty Then
                    o_udtErrorType = etEmpty
                Else
                
                End If
            End If
        End If
    
        'if there is something wrong, give hint
        If o_udtErrorType <> etDefault Then
            On Error Resume Next
            .SetFocus
            On Error GoTo 0
            
            If blnShowMsg Then
                Select Case o_udtErrorType
                        Case etRange
                            MsgBox "Your input is out of the required range." & vbCrLf & _
                                    "Valid range is from: " & curMinNum & _
                                    " to " & curMaxNum, vbInformation
                        Case etMaxLength
                            MsgBox "Your input exceeds the [maximum length]." & _
                                    "The maximum length is: " & _
                                    CStr(intMaxLength), vbInformation
                        Case etInvalid
                            MsgBox "Input requires to be all numbers(including .(dot))." & vbCrLf & _
                                    "Please check if your input contains non-number chars." _
                                    , vbInformation
                        Case etEmpty
                            MsgBox "Input could not be empty." & vbCrLf & _
                                    "You should input a valid number" _
                                    , vbInformation
                        Case Else
                            
                End Select
            Else
                
            End If
        Else
            
        End If
        
    End With
        
    ValidateNumber = o_udtErrorType
    
    Exit Function
    
    
handleError:
    ValidateNumber = etUnknown
    If blnShowMsg Then
        MsgBox "Error occurs when validating the number:" & vbCrLf _
                 & Err.Description
    Else
    End If
    On Error GoTo 0
    
End Function


Public Function ValidateText(ByRef objInputBox As Object, _
                              ByVal strForbiddenChars As String, _
                              Optional ByVal strSplitChar As String = "|", _
                              Optional ByVal strReplaceChar As String = vbNullString, _
                              Optional ByVal blnAuthReplace As Boolean = False, _
                              Optional ByVal intMaxLength As Integer = 0, _
                              Optional ByVal blnCanBeEmpty As Boolean = False, _
                              Optional ByVal blnShowMsg As Boolean = False _
                              ) As enumErrorType
Attribute ValidateText.VB_Description = "Ð£ÑéÊäÈë£¨×Ö·û£©"
'********************************************
'*
'*   ¡óFunction:ValidateText
'*         work:input verify(text)
'*         para:objInputBox               ---- Input Object(support TextBox¡¢ComboBox etc.)
'*               strForbiddenChars        ---- illegal chars
'*               strSplitChar(optional)   ---- split char
'*               strReplaceChar(optional) ---- char to be replaced
'*               blnAuthReplace(optional) ---- auto replace illegal chars
'*               intMaxLength(optional)   ---- maximum length
'*               blnCanBeEmpty(optional)  ---- allow empty input
'*               blnShowMsg(optional)     ---- show help message
'*       return:[enumErrorType]if valid input, return etDefault(0), otherwise, enumErrorType
'*
'* Last Modified by:Unruled Boy @ 1/11/2002
'*
'********************************************
        
On Error GoTo handleError
    
    Dim o_strRet() As String
    Dim o_strText As String
    Dim o_intItems As Integer
    Dim o_intItems2 As Integer
    Dim o_intPos As Integer
    Dim o_intPos2 As Integer
    Dim o_intLen As Integer
    Dim o_udtErrorType As enumErrorType
    
    o_udtErrorType = etDefault 'initial error:none
    
    With objInputBox
        
        Select Case TypeName(objInputBox)
                Case "TextBox", "RichTextBox"
                    'if there is a maximum length, use it
                    If .MaxLength > 0 Then
                        If intMaxLength = 0 Then
                            intMaxLength = .MaxLength
                        Else
                            If intMaxLength > .MaxLength Then
                                intMaxLength = .MaxLength
                            Else
                                
                            End If
                        End If
                    Else
                        
                    End If
                Case Else
                
        End Select
        
        o_strText = .Text
    
        
        'if the current length exceeds the maximum length,
        'show hint
        If intMaxLength > 0 And Len(o_strText) > intMaxLength Then
            
            o_udtErrorType = etMaxLength
            
        Else
            
            If o_strText <> vbNullString And strForbiddenChars <> vbNullString Then
            
                o_intItems2 = 0
                
                o_intPos2 = Len(o_strText)
                o_intLen = 0
                
                o_strRet() = Split(strForbiddenChars, strSplitChar)
                
                For o_intItems = LBound(o_strRet) To UBound(o_strRet)
                    o_intPos = InStr(o_strText, o_strRet(o_intItems))
                    'to find the first place the invalid char occur
                    If o_intPos > 0 And o_intPos < o_intPos2 Then
                        o_intPos2 = o_intPos
                        o_intLen = Len(o_strRet(o_intItems))
                    End If
                    
                    If o_intPos <> 0 Then
                        
                        If blnAuthReplace Then
                           o_strText = Replace(o_strText, _
                                                o_strRet(o_intItems), _
                                                strReplaceChar)
                           o_udtErrorType = etModified
                        Else
                            o_udtErrorType = etInvalid
                        End If
                    Else
                        o_intItems2 = o_intItems2 + 1
                    End If
                Next
                            
                If blnAuthReplace Then
                    .Text = o_strText
                Else
                    .SelStart = o_intPos2 - 1
                    .SelLength = o_intLen
                End If
                            
                If o_intItems2 > UBound(o_strRet) Then
                    o_udtErrorType = etDefault
                Else
                
                End If
                
            Else
                If Not blnCanBeEmpty Then
                    o_udtErrorType = etEmpty
                Else
                
                End If
            End If
        End If
    
        'if there is something wrong, give hint
        If o_udtErrorType <> etDefault Then
            On Error Resume Next
            .SetFocus
            On Error GoTo 0
                    
            If blnShowMsg Then
                Select Case o_udtErrorType
                        Case etInvalid
                            MsgBox "Your input is invalid." & vbCrLf & _
                                    "The following chars are not allowed: " & vbCrLf & _
                                    Join(o_strRet, "£¬") & "." _
                                    , vbInformation
                        Case etModified
                            MsgBox "Your input is invalid." & vbCrLf & _
                                    "The following chars are not allowed: " & _
                                    Join(o_strRet, "£¬") & "." & vbCrLf & _
                                    IIf(blnAuthReplace, vbCrLf & _
                                       "The invalid char(s) is/are replaced.", vbNullString) _
                                    , vbInformation
                        Case etEmpty
                            MsgBox "Input could not be empty." & vbCrLf & _
                                    "You should input a valid text" _
                                    , vbInformation
                End Select
            Else
                
            End If
        Else
        
        End If
    
    End With
    
    ValidateText = o_udtErrorType
    
    Exit Function
    
    
handleError:
    ValidateText = etUnknown
    If blnShowMsg Then
        MsgBox "Error occurs when validating the text:" & vbCrLf _
                 & Err.Description
    Else
    End If
    On Error GoTo 0
    
End Function

