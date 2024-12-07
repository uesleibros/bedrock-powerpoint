Attribute VB_Name = "XON"
' **************************************************************************
' Bedrock Team - XON Parser
' **************************************************************************
' This module provides a parser and serializer for own language named XON (eXtensible Object Notation), a custom
' hierarchical key-value format. It facilitates tokenizing, parsing, and
' generating XON data structures, enabling seamless integration with VBA projects
' that require advanced data handling.
'
' Designed to handle complex nested data such as dictionaries, lists, and primitives,
' the XON parser is highly versatile. It supports operations like tokenizing
' source code into structured elements, validating syntax, and converting parsed
' structures back into the XON format with precise serialization.
'
' Key features:
' - Robust tokenization, with line and column tracking for error reporting.
' - Supports nested blocks (`[...]`), lists (`(...)`), and key-value assignments.
' - Custom error handling for detailed syntax error feedback.
' - Stringify function to serialize XON data back to string format.
'
' --------------------------------------------------------------------------
' Team: Bedrock
' Last Update: 12/06/2024
' --------------------------------------------------------------------------
' References:
' - XON Specification and Examples (Custom Design)
' - JSON Serialization and Nested Data Structures
' - Data Serialization Techniques in Programming
' --------------------------------------------------------------------------

Option Explicit

Private index As Long
Private Tokens As List, positions As List
Private commonTokensCache As Dictionary
Private splited_tokens() As String
Private Const SYMBOLS As String = "[ ] ( ) ->"

Public Function Parse(Code As String) As Object
    Tokenize (Code)
    If index >= Tokens.Length Then RaiseError ("Unexpected end of input.")
    
    Dim Token As String
    Token = Tokens(index)
    
    Select Case Token
        Case "["
            Set Parse = DeserializeBlock
        Case "("
            Set Parse = DeserializeList
        Case Else
            RaiseError ("Expected start with '" & splited_tokens(0) & "' or '" & splited_tokens(2) & "'.")
    End Select
End Function

' Stringify credits to JSON from PPTGames, reference used to this.
Public Function Stringify(Value As Object) As String
    If TypeName(Value) = "List" Then
        Stringify = TSArray(Value, 1)
    ElseIf TypeName(Value) = "Dictionary" Then
        Stringify = TSDict(Value, 1)
    Else
        RaiseError "Invalid object type '" & TypeName(Value) & "'. Expected List or Dictionary."
    End If
End Function

Private Sub InitializeCommonTokens()
    Set commonTokensCache = New Dictionary
    commonTokensCache.Add "true", True
    commonTokensCache.Add "false", False
    commonTokensCache.Add "null", "null"
End Sub

Private Sub Tokenize(Code As String)
    Set Tokens = New List
    Set positions = New List
    index = 0
    splited_tokens = Split(SYMBOLS, " ")
    InitializeCommonTokens

    Dim i As Long, Line As Long, Col As Long
    Dim in_string As Boolean
    Dim string_delim As String
    Dim Char As String, buffer As String
    Line = 1
    Col = 1
    in_string = False
    string_delim = ""

    For i = 1 To Len(Code)
        Char = Mid(Code, i, 1)

        If Char = vbLf Then
            Line = Line + 1
            Col = 1
        ElseIf Char = """" Or Char = "'" Then
            If in_string And string_delim = Char Then
                ' Fecha string
                buffer = buffer & Char
                ProcessBuffer buffer, Line, Col
                buffer = ""
                in_string = False
                string_delim = ""
            ElseIf Not in_string Then
                ' Inicia string
                in_string = True
                string_delim = Char
                buffer = Char
            Else
                buffer = buffer & Char
            End If
        ElseIf in_string Then
            buffer = buffer & Char
        ElseIf IsSymbol(Char) Then
            ProcessBuffer buffer, Line, Col
            If Char = "-" And Mid(Code, i + 1, 1) = ">" Then
                AddToken "->", Line, Col
                i = i + 1
                Col = Col + 1
            Else
                AddToken Char, Line, Col
            End If
        ElseIf IsNumeric(Char) Or Char = "." Then
            If Len(buffer) > 0 And Not IsNumeric(Left(buffer, 1)) Then
                AddToken buffer, Line, Col - Len(buffer)
                buffer = ""
            End If
            buffer = buffer & Char
        ElseIf IsSpace(Char) Then
            ProcessBuffer buffer, Line, Col
        Else
            If Len(buffer) > 0 And IsNumeric(Left(buffer, 1)) And Not IsNumeric(Char) Then
                AddToken buffer, Line, Col - Len(buffer)
                buffer = ""
            End If
            buffer = buffer & Char
        End If

        Col = Col + 1
    Next i

    If Len(Trim(buffer)) > 0 Then ProcessBuffer buffer, Line, Col
    If in_string Then Err.Raise vbObjectError, "XON", "Error at line " & Line & ", column " & Col & ": Unterminated string"
End Sub

Private Sub AddToken(Token As String, Line As Long, Col As Long)
    If Len(Trim(Token)) > 0 Then
        Tokens.Add Token
        positions.Add Array(Line, Col)
    End If
End Sub

Private Sub ProcessBuffer(buffer As String, Line As Long, Col As Long)
    If Len(Trim(buffer)) > 0 Then
        AddToken buffer, Line, Col - Len(buffer)
        buffer = ""
    End If
End Sub

Private Function IsSymbol(Char As String) As Boolean
    IsSymbol = InStr(SYMBOLS, Char) > 0
End Function

Private Function IsSpace(Char As String) As Boolean
    IsSpace = Char = " " Or Char = vbTab Or Char = vbCr Or Char = vbLf
End Function

Private Function Assignment() As Dictionary
    Set Assignment = New Dictionary
    If index >= Tokens.Length Then RaiseError ("Unexpected end of input.")
    
    Dim Token As String, Key As String
    Token = Tokens(index)
    
    If Token Like "[\" & splited_tokens(0) & "\" & splited_tokens(2) & "]" Then RaiseError ("Unexpected '" & Token & "' at the start of input. Expecting key or assignment.")
    
    Key = GetKey
    If index >= Tokens.Length Or Tokens(index) <> "->" Then RaiseError ("Expected '->' after key '" & Key & "'")
    index = index + 1
    
    Assignment.Add Key, GetValue
End Function

Private Function GetKey() As String
    If index >= positions.Length Then RaiseError ("Unexpected end of input while parsing value")
    
    Dim Token As String
    Token = Tokens(index)
    index = index + 1

    If (Left(Token, 1) = """") And (Right(Token, 1) = """") Then
        GetKey = Mid(Token, 2, Len(Token) - 2)
    ElseIf (Left(Token, 1) = "'") And (Right(Token, 1) = "'") Then
        GetKey = Mid(Token, 2, Len(Token) - 2)
    Else
        GetKey = Token
    End If
End Function

Private Function GetValue() As Variant
    If index >= positions.Length Then RaiseError ("Unexpected end of input while parsing key")
    
    Dim Token As String
    Token = Tokens(index)
    
    If commonTokensCache.Exists(Token) Then
        GetValue = commonTokensCache(Token)
        index = index + 1
        Exit Function
    End If
    
    Select Case Token
        Case splited_tokens(0)
            Set GetValue = DeserializeBlock
            Exit Function
        Case splited_tokens(2)
            Set GetValue = DeserializeList
            Exit Function
    End Select
    
    If (Left(Token, 1) = """") And (Right(Token, 1) = """") Then
        index = index + 1
        GetValue = Mid(Token, 2, Len(Token) - 2)
        Exit Function
    ElseIf (Left(Token, 1) = "'") And (Right(Token, 1) = "'") Then
        index = index + 1
        GetValue = Mid(Token, 2, Len(Token) - 2)
        Exit Function
    End If
    
    If IsNumeric(Token) Then
        GetValue = CDbl(Token)
        index = index + 1
    End If
End Function


Private Function DeserializeBlock() As Dictionary
    If index >= Tokens.Length Or Tokens(index) <> splited_tokens(0) Then RaiseError ("Expected '" & splited_tokens(0) & "' to start a block")
    index = index + 1
    
    Set DeserializeBlock = New Dictionary
    
    Dim Token As String, Key As String
    While index < Tokens.Length
        Token = Tokens(index)
        If Token = splited_tokens(1) Then
            index = index + 1
            Exit Function
        End If
        
        If index + 1 >= Tokens.Length Or Tokens(index + 1) <> "->" Then RaiseError ("Expected '->' after key '" & Token & "'")
        Key = GetKey
        IExpect ("->")
        If index >= Tokens.Length Or Tokens(index) Like "[\]\)]" Then RaiseError ("Missing value for key '" & Key & "'")
        DeserializeBlock.Add Key, GetValue
    Wend
    
    RaiseError ("Unterminated block, missing '" & splited_tokens(1) & "'")
End Function

Private Function DeserializeList() As List
    If index >= Tokens.Length Or Tokens(index) <> "(" Then RaiseError ("Expected '" & splited_tokens(2) & "' to start a list")
    index = index + 1
    
    Set DeserializeList = New List
    
    Dim Token As String
    While index < Tokens.Length
        Token = Tokens(index)
        
        If Token = splited_tokens(3) Then
            index = index + 1
            Exit Function
        End If
        
        DeserializeList.Add GetValue
    Wend
    
    RaiseError ("Unterminated block, missing '" & splited_tokens(3) & "'")
End Function

Private Function TSArray(ByVal obj_list As List, level As Integer) As String
    If obj_list.Length = 0 Then
        TSArray = splited_tokens(2) & splited_tokens(3)
    Else
        Dim i As Long
        Dim buffer As String
        buffer = splited_tokens(2) & vbNewLine
        
        For i = 0 To obj_list.Length - 1
            buffer = buffer & RepeatString(String(2, " "), level) & TSExpression(obj_list(i), level)
            If i < obj_list.Length - 1 Then buffer = buffer & vbNewLine
        Next i
        buffer = buffer & vbNewLine & RepeatString(String(2, " "), level - 1) & splited_tokens(3)
        TSArray = buffer
    End If
End Function

Private Function TSDict(ByVal obj_dict As Dictionary, level As Integer) As String
    If obj_dict.count = 0 Then
        TSDict = splited_tokens(0) & splited_tokens(1)
    Else
        Dim i As Long
        Dim buffer As String
        buffer = splited_tokens(0) & vbNewLine
        
        For i = 0 To obj_dict.count - 1
            buffer = buffer & RepeatString(String(2, " "), level) & """" & obj_dict.Keys(i) & """ -> " & TSExpression(obj_dict.items(i), level)
            If i < obj_dict.count - 1 Then buffer = buffer & vbNewLine
        Next i
        buffer = buffer & vbNewLine & RepeatString(String(2, " "), level - 1) & splited_tokens(1)
        TSDict = buffer
    End If
End Function

Private Function TSExpression(Expression As Variant, Ind As Integer) As String
    Select Case VarType(Expression)
        Case vbBoolean
            TSExpression = IIf(Expression, "true", "false")
        Case vbString
            TSExpression = """" & Expression & """"
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbDecimal, vbCurrency
            TSExpression = CStr(Expression)
        Case Else
            Select Case True
                Case TypeName(Expression) = "List"
                    TSExpression = TSArray(Expression, Ind + 1)
                Case TypeName(Expression) = "Dictionary"
                    TSExpression = TSDict(Expression, Ind + 1)
                Case IsArray(Expression)
                    TSExpression = TypeName(Expression) & "(" & JoinArray(Expression, Ind) & ")"
                Case Else
                    TSExpression = CStr(Expression)
            End Select
    End Select
End Function

Private Function JoinArray(Arr As Variant, Ind As Integer) As String
    Dim i As Long
    Dim elements() As String
    
    ReDim elements(LBound(Arr) To UBound(Arr))
    
    For i = LBound(Arr) To UBound(Arr)
        elements(i) = EvaluateExpression(Arr(i), Ind)
    Next i
    
    JoinArray = Join(elements, ", ")
End Function

Private Function TSString(ByVal Expression As String) As String
    Dim result As String
    Dim i As Long
    Dim currentChar As String

    result = """"
    For i = 1 To Len(Expression)
        currentChar = Mid(Expression, i, 1)
        
        Select Case currentChar
            Case """"
                result = result & "\"""
            Case vbLf
                result = result & "\n"
            Case vbCr
                result = result & "\r"
            Case vbTab
                result = result & "\t"
            Case vbCrLf
                result = result & "\n"
            Case Else
                result = result & currentChar
        End Select
    Next i

    result = result & """"
    TSString = result
End Function

Private Function RepeatString(s As String, count As Integer) As String
    RepeatString = String(count * 2, s)
End Function

Private Sub IExpect(expected_token As String)
    If index >= Tokens.Length Then RaiseError ("Unexpected end of input, expected '" & expected_token & "'")
    If Tokens(index) <> expected_token Then RaiseError ("Expected '" & expected_token & "', got '" & Tokens(index) & "'")
    index = index + 1
End Sub

Private Sub RaiseError(message As String)
    Dim Line As Long, Col As Long
    
    If positions.Length > 0 And index <= positions.Length Then
        Line = positions(index)(0)
        Col = positions(index)(1)
    End If
    
    Err.Raise vbObjectError, "XON", "Error at line " & Line & ", column " & Col & ": " & message
End Sub
