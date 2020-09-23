Attribute VB_Name = "HowToUse"

'**********************************************************************************************************************
'*
'*   Program Name: Smart Flex Input
'*   Description : filter, verify and correct input
'*   Version     : 0.5
'*   Copyright   : Smartsoft 2002
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

''Error Type
'Public Enum enumErrorType
'    etDefault = 0 'Valid
'    etInvalid = 1 'Invalid
'    etRange = 2 'input range
'    etMaxLength = 3 'maximum length
'    etEmpty = 4 'empty
'    etModified = 5 'auto correct
'    etUnknown = 99 'unknow error
'End Enum
'
''input type
'Public Enum enumInputType
'    itNumeric = 0 'number
'    itChar = 1 'char
'    itDate = 2 'date
'    itTime = 3 'time
'    itID = 4 'id
'End Enum
'
'Public Enum enumDateTime
'    dtDate = 0 'date
'    dtTime = 1 'time
'End Enum
