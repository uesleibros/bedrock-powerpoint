Attribute VB_Name = "JSONF"
' **************************************************************************
' Bedrock Team - JSONF
' **************************************************************************
' This module provides an implementation of JSON parsing and stringifying
' functionality. Inspired by various open-source projects, this module aims
' to offer a simple and lightweight approach to handle JSON data in VBA.
' With support for parsing both objects and arrays, it efficiently converts
' between JSON strings and VBA objects. This tool is especially useful for
' applications in data exchange, game development, and more complex computational
' tasks that require interoperability with JSON.
' --------------------------------------------------------------------------
' Team: Bedrock
' Last Update: 12/01/2024
' **************************************************************************
' References:
' - C-Simple-JSON-Parser (GitHub)
'   https://github.com/forkachild/C-Simple-JSON-Parser
' - JSONPPTGames
'   https://pptgamespt.wixsite.com/pptg-coding/json
' **************************************************************************

Option Explicit

Private idx As Long

Public Function parse(json_string As String) As Variant
    Dim i As Long, i2 As Long
    Dim curw As String, curc As String, ncurc As String, state_exp As String
    Dim args As New List
    Dim json_string_size As Long

    json_string = Replace(json_string, vbCrLf, "")
    json_string = Replace(json_string, vbLf, "")
    json_string = Replace(json_string, vbTab, "")
    
    If Len(json_string) = 0 Then
        Err.Raise 1, "JSON2.parse()", "Empty json."
        Exit Function
    End If
    
    json_string_size = Len(json_string)
    
    For i = 1 To json_string_size
        curc = Mid(json_string, i, 1)
        ncurc = Mid(json_string, i + 1, 1)
        
        If state_exp <> "" Then
            If curc = state_exp Then
                state_exp = ""
                curw = curw & curc
            ElseIf curc = "\" Then
                If InStr(1, "\""'bfnrtu", ncurc) Then
                    Select Case ncurc
                        Case "b": curw = curw & vbBack
                        Case "f": curw = curw & vbFormFeed
                        Case "n": curw = curw & vbNewLine
                        Case "r": curw = curw & vbCr
                        Case "t": curw = curw & vbTab
                        Case "u": curw = curw & ChrW(val("&h" & Mid(json_string, i + 2, 4))): i = i + 4
                        Case Else: curw = curw & ncurc
                    End Select
                    i = i + 1
                Else
                    curw = curw & curc
                End If
            Else
                curw = curw & curc
            End If
        Else
            If curc = """" Or curc = "'" Then
                state_exp = curc
                curw = curw & curc
            ElseIf InStr(1, "[]{}:,", curc) Then
                For i2 = i To json_string_size
                    If Mid(json_string, i2, 1) <> " " And Mid(json_string, i2, 1) <> ":" Then
                        i2 = -1
                        Exit For
                    End If
                    If Mid(json_string, i2, 1) = ":" Then Exit For
                Next i2
                If curw <> "" Then
                    args.Add handle_expression(curw, i2, json_string)
                End If
                curw = ""
                If curc <> "," Then args.Add curc
            ElseIf curc <> " " Then
                curw = curw & curc
            End If
        End If
    Next i
    
    If args.Length > 0 Then
        Select Case args(0)
            Case "["
                If args(1) = "]" Then
                    Set parse = New Collection
                Else
                    Set parse = parse_list(args.Slice(1, args.Length - 2))
                End If
            Case "{"
                If args(1) = "}" Then
                    Set parse = New Dictionary
                Else
                    Set parse = parse_dict(args.Slice(1, args.Length - 2))
                End If
        End Select
    Else
        Set parse = Nothing
    End If
End Function

Public Function stringify(ByVal json_object As Object) As String
    Dim json_object_type As String
    Dim json_object_size As Long
    
    json_object_type = TypeName(json_object)
    idx = 0
    
    Select Case json_object_type
        Case "Dictionary"
            stringify = IIf(json_object.Count > 0, "{" & stringify_dict(json_object) & "}", "{}")
        Case "List"
            stringify = IIf(json_object.Length > 0, "[" & stringify_list(json_object) & "]", "[]")
    End Select
End Function

Private Function handle_expression(curw As String, i As Long, json_string As String) As Variant
    Dim numeric_expr As String, quoted_curw As Boolean
    Dim unquoted_expr As String
    
    quoted_curw = InStr(1, """'", Left(curw, 1))
    numeric_expr = Replace(curw, ".", Format(0, "."))
    
    If i > -1 Then
        If Not quoted_curw Then
            Err.Raise 1, "JSON2.handle_expression()", "Unquoted key: " & vbCrLf & vbCrLf & curw & _
            vbCrLf & "^" & vbCrLf & "Expected: "" or '"
            Exit Function
        End If
    End If
    
    If Not quoted_curw Then
        Select Case curw
            Case "true", "false"
                handle_expression = curw = "true"
            Case "null"
                handle_expression = "null"
            Case Else
                If IsNumeric(numeric_expr) Then
                    handle_expression = CDbl(numeric_expr)
                End If
        End Select
    Else
        unquoted_expr = Mid(curw, 2, Len(curw) - 2)
        Select Case unquoted_expr
            Case "[", "{"
                handle_expression = "\/\/\S: " & unquoted_expr
            Case Else
                handle_expression = unquoted_expr
        End Select
    End If
End Function

Private Function parse_dict(args As List) As Dictionary
    Set parse_dict = New Dictionary
    
    Dim i As Long, i2 As Long, s As Long, key As String
    Dim args_size As Long
    
    args_size = args.Length - 1
    
    For i = 0 To args_size
        If i > 0 Then
            If args(i - 1) = ":" Then
                If args(i) = "{" Then
                    If args(i + 1) = "}" Then
                        parse_dict.Add key, New Dictionary
                        i = i + 1
                    Else
                        s = 0
                        For i2 = i To args_size
                            Select Case args(i2)
                                Case "{"
                                    s = s + 1
                                Case "}"
                                    If s > 0 Then s = s - 1
                                    If s = 0 Then Exit For
                            End Select
                        Next i2
                        parse_dict.Add key, parse_dict(args.Slice(i + 1, i2 - 1))
                        i = i2
                    End If
                ElseIf args(i) = "[" Then
                    If args(i + 1) = "]" Then
                        parse_dict.Add key, New List
                        i = i + 1
                    Else
                        s = 0
                        For i2 = i To args_size
                            Select Case args(i2)
                                Case "["
                                    s = s + 1
                                Case "]"
                                    If s > 0 Then s = s - 1
                                    If s = 0 Then Exit For
                            End Select
                        Next i2
                        parse_dict.Add key, parse_list(args.Slice(i + 1, i2 - 1))
                        i = i2
                    End If
                Else
                    parse_dict.Add key, args(i)
                End If
            End If
        End If
        
        If i < args_size Then
            If args(i + 1) = ":" Then
                key = args(i)
            End If
        End If
    Next i
End Function

Private Function parse_list(args As List) As List
    Set parse_list = New List

    Dim i As Long, i2 As Long, s As Long, key As String
    Dim args_size As Long
    
    args_size = args.Length - 1

    For i = 0 To args_size
        If args(i) = "[" Then
            If args(i + 1) = "]" Then
                parse_list.Add New List
                i = i + 1
            Else
                s = 0
                For i2 = i To args_size
                    Select Case args(i2)
                        Case "["
                            s = s + 1
                        Case "]"
                            If s > 0 Then s = s - 1
                            If s = 0 Then Exit For
                    End Select
                Next i2
                parse_list.Add parse_list(args.Slice(i + 1, i2 - 1))
                i = i2
            End If
        ElseIf args(i) = "{" Then
            If args(i + 1) = "}" Then
                parse_list.Add New Dictionary
                i = i + 1
            Else
                s = 0
                For i2 = i To args_size
                    Select Case args(i2)
                        Case "{"
                            s = s + 1
                        Case "}"
                            If s > 0 Then s = s - 1
                            If s = 0 Then Exit For
                    End Select
                Next i2
                parse_list.Add parse_dict(args.Slice(i + 1, i2 - 1))
                i = i2
            End If
        Else
            Dim formated_value As String
            formated_value = args(i)
            
            If Left(formated_value, 8) = "\/\/\S: " Then
                formated_value = Mid(formated_value, 9)
            End If
            parse_list.Add formated_value
        End If
    Next i
End Function

Private Function stringify_dict(ByVal json_object As Dictionary) As String
    Dim i As Long
    Dim json_object_size As Long
    Dim cur_json_object_key As String, cur_json_object_type As String
    Dim key_part As String
    
    idx = idx + 1
    json_object_size = UBound(json_object.Items)
    
    For i = 0 To json_object_size
        cur_json_object_key = json_object.Keys(i)
        cur_json_object_type = TypeName(json_object.Items(i))
        key_part = vbCrLf & GetIdx & """" & cur_json_object_key & """"

        Select Case cur_json_object_type
            Case "Dictionary"
                If json_object.Items(i).Count > 0 Then
                    stringify_dict = stringify_dict & key_part & ": {" & stringify_dict(json_object.Items(i)) & _
                    IIf(i = json_object_size, GetIdx & "}" & vbCrLf, GetIdx & "},")
                Else
                    stringify_dict = stringify_dict & key_part & ": {" & IIf(i = json_object_size, "}" & vbCrLf, "},")
                End If
            Case "List"
                If json_object.Items(i).Length > 0 Then
                    stringify_dict = stringify_dict & key_part & ": [" & stringify_list(json_object.Items(i)) & _
                    IIf(i = json_object_size, GetIdx & "]" & vbCrLf, GetIdx & "],")
                Else
                    stringify_dict = stringify_dict & key_part & ": [" & IIf(i = json_object_size, "]" & vbCrLf, "],")
                End If
            Case "Boolean"
                stringify_dict = stringify_dict & key_part & ": " & IIf(json_object.Items(i), "true", "false") & IIf(i = json_object_size, vbCrLf, ",")
            Case "String"
                stringify_dict = stringify_dict & key_part & ": """ & stringify_ue(json_object.Items(i)) & IIf(i = json_object_size, """" & vbCrLf, """,")
            Case Else
                stringify_dict = stringify_dict & key_part & ": " & Replace(json_object.Items(i), ",", ".") & IIf(i = json_object_size, vbCrLf, ",")
        End Select
    Next i
    
    idx = idx - 1
End Function

Private Function stringify_list(ByVal json_object As List) As String
    Dim i As Long
    Dim json_object_size As Long
    Dim cur_json_object_type As String
    Dim key_part As String
    
    idx = idx + 1
    json_object_size = json_object.Length - 1
    
    For i = 0 To json_object_size
        cur_json_object_type = TypeName(json_object(i))
        key_part = vbCrLf & GetIdx
        
        Select Case cur_json_object_type
            Case "Dictionary"
                If json_object(i).Count > 0 Then
                    stringify_list = stringify_list & key_part & "{" & stringify_dict(json_object(i)) & GetIdx & _
                    IIf(i = json_object_size, "}" & vbCrLf, "},")
                Else
                    stringify_list = stringify_list & key_part & "{" & IIf(i = json_object_size, "}" & vbCrLf, "},")
                End If
            Case "List"
                If json_object(i).Length > 0 Then
                    stringify_list = stringify_list & key_part & "[" & stringify_list(json_object(i)) & GetIdx & _
                    IIf(i = json_object_size, "]" & vbCrLf, "],")
                Else
                    stringify_list = stringify_list & key_part & "[" & IIf(i = json_object_size, "]" & vbCrLf, "],")
                End If
            Case "Boolean"
                stringify_list = stringify_list & key_part & IIf(json_object(i), "true", "false") & IIf(i = json_object_size, vbCrLf, ",")
            Case "String"
                stringify_list = stringify_list & key_part & """" & stringify_ue(json_object(i)) & IIf(i = json_object_size, """" & vbCrLf, """,")
            Case Else
                stringify_list = stringify_list & key_part & Replace(json_object(i), ",", ".") & IIf(i = json_object_size, vbCrLf, ",")
        End Select
    Next i
    
    idx = idx - 1
End Function

Private Function GetIdx() As String
    GetIdx = String(idx * 2, " ")
End Function

Private Function stringify_ue(ByVal expr As String) As String
    Dim i As Long
    For i = 1 To Len(expr)
        Select Case Mid(expr, i, 1)
            Case vbBack: stringify_ue = stringify_ue & "\b"
            Case vbFormFeed: stringify_ue = stringify_ue & "\f"
            Case vbLf: stringify_ue = stringify_ue & "\n"
            Case vbTab: stringify_ue = stringify_ue & "\t"
            Case "\": stringify_ue = stringify_ue & "\\"
            Case """": stringify_ue = stringify_ue & "\"""
            Case "'": stringify_ue = stringify_ue & "\'"
            Case Else: If Mid(expr, i + 1, 1) <> vbLf Then stringify_ue = stringify_ue & Mid(expr, i, 1)
        End Select
    Next i
End Function
