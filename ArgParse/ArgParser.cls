VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArgParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Enums
Private Enum States
    Waiting
    PlainArgument
    OptionKey
    OptionValue
    Number  'Parse to Long or Double
    SingleQuoteString
    DoubleQuoteString
End Enum

Private Enum TokenTypes
    IntegerToken  'Actually "Long" xd
    FloatToken
    StringToken
    OptionKeyToken
    OptionValueToken
    PlainArgumentToken
End Enum

'Enums
Public Enum OptionsStyles
    DOS
    UNIX
End Enum

'Private Types
Private Type Token
    Type As TokenTypes
    Value As Variant
End Type

'Private Variables
Private Position As Long  'Position
Private Cache As cStringBuilder  'Cached val
Private State As States  'Lexer state
Private Tokens() As Token  'Tokens
Private Arguments As Collection  'Actually Collection(Of Argument)
Private ParsingCommand As String  'Parsing command (copy of "Command" passed in)

Private Aliases_ As Dictionary  'Option / flag aliases
Private Options_ As Dictionary  'Options

Private PlainArgs__ As Dictionary, _
        Options__ As Dictionary, _
        Flags__ As Dictionary, _
        PlainArgsAsCollection__ As Collection

Private cachedPosition As Long, cachedCurrentChar As String

Private WalkPosition As Long

'Properties
Private OptionsStyle_ As OptionsStyles

Public Property Let OptionsStyle(setOptionsStyle As OptionsStyles)
    OptionsStyle_ = setOptionsStyle
    If OptionsStyle_ = DOS Then
        Aliases_.CompareMode = TextCompare
        Options_.CompareMode = TextCompare
        PlainArgs__.CompareMode = TextCompare
        Options__.CompareMode = TextCompare
        Flags__.CompareMode = TextCompare
    ElseIf OptionsStyle_ = UNIX Then
        Aliases_.CompareMode = BinaryCompare
        Options_.CompareMode = BinaryCompare
        PlainArgs__.CompareMode = BinaryCompare
        Options__.CompareMode = BinaryCompare
        Flags__.CompareMode = BinaryCompare
    End If
End Property

Public Property Get PlainArgs() As Dictionary
    Set PlainArgs = PlainArgs__
End Property

Public Property Get Options() As Dictionary
    Set Options = Options__
End Property

Public Property Get Flags() As Dictionary
    Set Flags = Flags__
End Property

Public Property Get PlainArgsAsCollection() As Collection
    Set PlainArgsAsCollection = PlainArgsAsCollection__
End Property

Public Property Get FlagEnabled(ByRef Flag As String) As Boolean
    If Flags__.Exists(Flag) Then FlagEnabled = True
End Property

'------------------------------
'Public functions & methods
'------------------------------

Public Sub Parse(ByRef Command As String)
Attribute Parse.VB_Description = "Parse arguments"
    'Parse arguments
    Me.ParseOneByOne Command
    Dim Arg As Argument
    For Each Arg In Arguments
        Select Case Arg.ArgType
            Case PlainArg
                PlainArgs__.Add Arg.Value, Arg.Value
            Case FlagArg
                If Not FlagEnabled(Arg.Key) Then Flags__.Add Arg.Key, True
            Case OptionArg
                If Options__.Exists(Arg.Key) Then
                    Options__.Item(Arg.Key) = Arg.Value
                Else
                    Options__.Add Arg.Key, Arg.Value
                End If
        End Select
    Next
    'Convert plain args dict to collection
    Dim PlainArg_ As Variant
    For Each PlainArg_ In PlainArgs__.Keys
        PlainArgsAsCollection__.Add PlainArg_
    Next
End Sub

Public Function ParseOneByOne(ByRef Command As String) As Collection  'Collection(Of Argument)
Attribute ParseOneByOne.VB_Description = "Parse arguments one by one to a collection for user to process manually"
    'Parse arguments one by one to a collection for user to process manually
    ParsingCommand = Command + " "  'To prevent last char to be escaped
    
    'Parse command string to tokens
    Position = 1
    ReDim Tokens(0 To 0)
    While Position <= Len(ParsingCommand)
        ApplyStateMethods
    Wend
    
    'Parse tokens to arguments
    Set Arguments = New Collection
    Dim i As Long, TokenValue As Variant, ArgumentToAdd As Argument
    i = 1
    While i <= UBound(Tokens)
        TokenValue = Tokens(i).Value
        Set ArgumentToAdd = New Argument
        Select Case Tokens(i).Type
            Case TokenTypes.IntegerToken
                With ArgumentToAdd
                    .ArgType = PlainArg
                    .ValueType = IntegerValue
                    .Value = TokenValue
                End With
                Arguments.Add ArgumentToAdd
            Case TokenTypes.FloatToken
                With ArgumentToAdd
                    .ArgType = PlainArg
                    .ValueType = FloatValue
                    .Value = TokenValue
                End With
                Arguments.Add ArgumentToAdd
            Case TokenTypes.OptionValueToken, TokenTypes.StringToken, TokenTypes.PlainArgumentToken
                With ArgumentToAdd
                    .ArgType = PlainArg
                    .ValueType = StringValue
                    .Value = TokenValue
                End With
                Arguments.Add ArgumentToAdd
            Case TokenTypes.OptionKeyToken
                With ArgumentToAdd
                    'Apply alias
                    If Aliases_.Exists(TokenValue) Then
                        .Key = Aliases_(TokenValue)
                    Else
                        .Key = TokenValue
                    End If
                    'Option or flag?
                    If Options_.Exists(.Key) Then
                        'Is option. Check next token
                        If UBound(Tokens) = i Then
                            'Error, value not found
                            Err.Raise OPTIONS_VALUE_NOT_FOUND, , ("Value not found for option: " & .Key)
                        End If
                        'Jump to next token
                        i = i + 1
                        If Tokens(i).Type = OptionKeyToken Then
                            'Error, value not found
                            Err.Raise OPTIONS_VALUE_NOT_FOUND, , ("Value not found for option: " & .Key)
                        Else
                            Select Case Tokens(i).Type
                                Case TokenTypes.IntegerToken: .ValueType = IntegerValue
                                Case TokenTypes.FloatToken: .ValueType = FloatValue
                                Case Else: .ValueType = StringValue
                            End Select
                            .Value = Tokens(i).Value
                        End If
                        .ArgType = OptionArg
                    Else
                        .ArgType = FlagArg
                        .Value = True
                    End If
                End With
                Arguments.Add ArgumentToAdd
        End Select
        i = i + 1
    Wend
    Set ParseOneByOne = Arguments
End Function

Public Sub SetAlias(ByRef LongOption As String, ByRef ShortAlias As String)
Attribute SetAlias.VB_Description = "Set short alias for options / flags"
    'Set short alias for options / flags
    If Aliases_.Exists(ShortAlias) Then
        Aliases_.Item(ShortAlias) = LongOption
    Else
        Aliases_.Add Key:=ShortAlias, Item:=LongOption
    End If
End Sub

Public Sub MarkAsOption(ParamArray LongOptions())
Attribute MarkAsOption.VB_Description = "Mark a flag as option\r\nOptions will be parsed with their values"
    'Mark a flag as option
    'Options will be parsed with their values
    Dim i As Long
    For i = LBound(LongOptions) To UBound(LongOptions)
        Options_.Add CStr(LongOptions(i)), True
    Next
End Sub

Public Function NextArg()
    NextArg = PlainArgsAsCollection__.Item(WalkPosition)
    WalkPosition = WalkPosition + 1
End Function

Public Function ThisArg()
    ThisArg = PlainArgsAsCollection__.Item(WalkPosition - 1)
End Function



'-------------------------
'Lexer functions
'-------------------------

Private Function CurrentChar() As String
    'Get current character (with cache)
    If cachedPosition <> Position Then
        'Update cached value
        cachedPosition = Position
        cachedCurrentChar = Mid$(ParsingCommand, Position, 1)
    'Else
        'Use cached value
    End If
    CurrentChar = cachedCurrentChar
End Function

Private Sub ApplyStateMethods()
    'Parse one character by current state
    Select Case State
        Case States.Waiting: StateWaiting
        Case States.Number: StateNumber
        Case States.SingleQuoteString: StateSingleQuoteString
        Case States.DoubleQuoteString: StateDoubleQuoteString
        Case States.OptionKey: StateOptionKey
        Case States.OptionValue, States.PlainArgument: StatePlainArgument
    End Select
    DoEvents
End Sub

Private Sub AppendCache()
    'Append current character to cache
    Cache.Append CurrentChar
End Sub

Private Sub SkipChar()
    'Skip current character
    Position = Position + 1
End Sub

Private Sub AppendToken(ByRef TokenType As TokenTypes, ByRef Value As Variant)
    'Append new token to list
    Dim TokenToAppend As Token
    With TokenToAppend
        .Type = TokenType
        .Value = Value
    End With
    ReDim Preserve Tokens(0 To UBound(Tokens) + 1)
    Tokens(UBound(Tokens)) = TokenToAppend
End Sub


'----------------------------
'"Is" functions
'----------------------------

Private Function IsSpace(ByRef Char) As Boolean
    IsSpace = (Char = " " Or Char = Chr(9))
End Function

Private Function IsNumber(ByRef Char) As Boolean
    IsNumber = ((Asc(Char) >= 48 And Asc(Char) <= 57) Or Char = ".")
End Function

Private Function IsOptionKeyIdentifier(ByRef Char) As Boolean
    If OptionsStyle_ = UNIX Then
        IsOptionKeyIdentifier = (Char = "-")
    ElseIf OptionsStyle_ = DOS Then
        IsOptionKeyIdentifier = (Char = "/")
    End If
End Function

Private Function IsOptionValueIdentifier(ByRef Char) As Boolean
    If OptionsStyle_ = UNIX Then
        IsOptionValueIdentifier = (Char = "=")
    ElseIf OptionsStyle_ = DOS Then
        IsOptionValueIdentifier = (Char = ":")
    End If
End Function

Private Function StringNeedTwoPassProcess(ByRef CacheStr) As Boolean
    If OptionsStyle_ = UNIX Then
        StringNeedTwoPassProcess = (Left$(CacheStr, 1) = "-") And (InStr(CacheStr, "=") <> 0)
    ElseIf OptionsStyle_ = DOS Then
        StringNeedTwoPassProcess = (Left$(CacheStr, 1) = "/") And (InStr(CacheStr, ":") <> 0)
    End If
End Function


'-------------------------
'State methods
'------------------------

Private Sub StateWaiting()
    'Waiting and switch state
    Cache.Clear
    
    If IsSpace(CurrentChar) Then
        'Space or Tab, skip this character
        SkipChar
        Exit Sub
    End If
    
    If CurrentChar = """" Then
        State = DoubleQuoteString
        SkipChar  'Skip quote
    ElseIf CurrentChar = "'" Then
        State = SingleQuoteString
        SkipChar  'Skip quote
    ElseIf IsNumber(CurrentChar) Then
        ' 0-9 or dot
        State = Number
    ElseIf IsOptionKeyIdentifier(CurrentChar) Then
        'Option key
        State = OptionKey
    ElseIf IsOptionValueIdentifier(CurrentChar) Then
        'Option value
        State = OptionValue
        SkipChar  'Skip "=" and ":"
    Else
        'Other characters?
        'Maybe plain argument
        State = PlainArgument
    End If
End Sub

Private Sub StateNumber()
    If Not IsNumber(CurrentChar) Then
        'Get cache string
        Dim CacheStr As String
        CacheStr = Cache.toString
        
        If InStr(CacheStr, ".") <> 0 Then
            'Is float
            AppendToken FloatToken, CDbl(CacheStr)
        Else
            AppendToken IntegerToken, CLng(CacheStr)
        End If
        
        State = Waiting
        Exit Sub
    End If
    
    AppendCache
    SkipChar
End Sub

Private Sub StateDoubleQuoteString()
    If CurrentChar = """" Then
        'End of string
        Dim CacheStr As String
        CacheStr = Cache.toString
        If StringNeedTwoPassProcess(CacheStr) Then
            ProcessOptionAsString CacheStr
        Else
            'Plain string
            AppendToken StringToken, CacheStr
        End If
        SkipChar  'Skip quote
        State = Waiting
        Exit Sub
    End If
    
    AppendCache
    SkipChar
End Sub

Private Sub StateSingleQuoteString()  'CMD compatible
    If CurrentChar = "'" Then
        'End of string
        Dim CacheStr As String
        CacheStr = Cache.toString
        If StringNeedTwoPassProcess(CacheStr) Then
            ProcessOptionAsString CacheStr
        Else
            'Plain string
            AppendToken StringToken, CacheStr
        End If
        SkipChar  'Skip quote
        State = Waiting
        Exit Sub
    End If
    
    AppendCache
    SkipChar
End Sub

Private Sub ProcessOptionAsString(OptionAsString As String)
    'Options with space in their values will be quoted by PowerShell
    'So they need to be processed again
    Dim pPosition As Long, pCurrentChar As String, pCache As New cStringBuilder
    pPosition = 1
    Do While pPosition <= Len(OptionAsString)
        pCurrentChar = Mid$(OptionAsString, pPosition, 1)
        If IsOptionKeyIdentifier(pCurrentChar) Then
            'Just skip them
            pPosition = pPosition + 1
        ElseIf IsOptionValueIdentifier(pCurrentChar) Then
            AppendToken OptionKeyToken, pCache.toString
            AppendToken OptionValueToken, Mid$(OptionAsString, pPosition + 1)
            Exit Do
        Else
            'Key
            pCache.Append pCurrentChar
            pPosition = pPosition + 1
        End If
    Loop
    Exit Sub
End Sub

Private Sub StateOptionKey()
    If IsOptionKeyIdentifier(CurrentChar) Then
        SkipChar
        Exit Sub
    End If
    
    If IsSpace(CurrentChar) Or IsOptionValueIdentifier(CurrentChar) Then
        AppendToken OptionKeyToken, Cache.toString
        SkipChar
        State = Waiting
        Exit Sub
    End If

    AppendCache
    SkipChar
End Sub

Private Sub StateOptionValue()
    If IsSpace(CurrentChar) Then
        AppendToken OptionValueToken, Cache.toString
        SkipChar
        State = Waiting
        Exit Sub
    End If
    
    AppendCache
    SkipChar
End Sub

Private Sub StatePlainArgument()

    If IsSpace(CurrentChar) Then
        AppendToken PlainArgumentToken, Cache.toString
        SkipChar
        State = Waiting
        Exit Sub
    End If
    
    AppendCache
    SkipChar
End Sub




Private Sub Class_Initialize()
    Set Aliases_ = New Dictionary
    Set Options_ = New Dictionary
    Set Cache = New cStringBuilder
    Set PlainArgs__ = New Dictionary
    Set Options__ = New Dictionary
    Set Flags__ = New Dictionary
    Set PlainArgsAsCollection__ = New Collection
    OptionsStyle = UNIX
    WalkPosition = 1
End Sub

