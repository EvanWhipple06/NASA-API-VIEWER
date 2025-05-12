Imports System.IO.Directory

Public Class Data
    Public Function parseData(ByRef var1 As String, Optional ByRef var2 As String = "")
        Dim pythonProcess As New Process()
        Dim pythonPath As String = $""""""          'path to the python installation on your machine
        Dim scriptPath As String = $""""""        'path to the python script on your machine

        Dim key As String = ""                        'your NASA api key here

        pythonProcess.StartInfo.FileName = pythonPath
        pythonProcess.StartInfo.Arguments = (scriptPath & " " & var1 & " " & var2 & " " & key)                       'pass arguments through to cmd prompt to be run in the python script
        pythonProcess.StartInfo.UseShellExecute = False                                                              'pass the console output bacck in a string
        pythonProcess.StartInfo.RedirectStandardOutput = True
        pythonProcess.StartInfo.CreateNoWindow = True
        pythonProcess.Start()
        Dim output As String = pythonProcess.StandardOutput.ReadToEnd()
        pythonProcess.WaitForExit()

        Return parseSingleArg(output)
    End Function

    Private Function parseSingleArg(ByVal output As String)
        Dim parsedData = New Dictionary(Of Object, Object)

        Dim state As Integer = 0                            'state: 0 = key, 1 = value, 2 = nested dictionary, 3 = list
        Dim key As String = "", value As String = "", valueD As Object

        For i = 1 To output.Length - 1                      'read the character and changes state based on the delimiter, and appends key/value to dictionary if relevant
            If (output(i).Equals(","c)) Then
                If Not key.Equals("") And Not value.Equals("") And state = 1 Then
                    cleanKey(key)
                    cleanVal(value)

                    parsedData.Add(key, value)
                Else
                    cleanKey(key)

                    parsedData.Add(key, valueD)
                End If

                key = ""
                value = ""
                valueD = ""
                state = 0
            ElseIf (output(i).Equals(":"c)) And Not state = 1 Then      'prevents links from breaking
                state = 1
            ElseIf output(i).Equals("{"c) Then
                state = 2
            ElseIf output(i).Equals("["c) Then
                state = 3
            End If

            If state = 0 Then                                       'if state = 0 then make the key
                key += output(i)
            ElseIf state = 1 Then                                   'if state = 1 then track the value (has a special case for if its at the end of the dict
                If output(i).Equals("}"c) Then
                    If value.Contains("False") Then                 'boolean handling instead of using cleanVal() if it reads a boolean it makes it a boolean in code
                        value = "False"
                    ElseIf value.Contains("True") Then
                        value = "True"
                    End If
                    cleanKey(key)
                    parsedData.Add(key, value)
                    key = ""
                    value = ""
                Else
                    value += output(i)
                End If
            ElseIf state = 2 Then                                   'run nested dictionary parsing
                valueD = parseNestedDict(i, output)
            ElseIf state = 3 Then
                valueD = parseArr(i, output)                        'run list parsing
            End If
        Next

        Return parsedData
    End Function

    Private Function parseArr(ByRef i As Integer, ByVal output As String)           'parses arrays for the close approach data
        Dim nestedList = New List(Of Object)
        Dim temp As Object
        For t = i + 1 To output.Length
            If output(t).Equals("{"c) Then                      'if it detects a dictionary it parses pested dict
                temp = parseNestedDict(t, output)

            ElseIf output(t).Equals(","c) Then                  'if it detects a change in index it adds the previous collected thing to the list
                nestedList.Add(temp)

            ElseIf output(t).Equals("]"c) Then                  'if it detects the end of an index it updates the counter variable and continues
                i = t

                Return nestedList
            End If
        Next
    End Function

    Private Function parseNestedDict(ByRef i As Integer, ByVal output As String)
        Dim nestedDict = New Dictionary(Of Object, Object)

        Dim state As Integer = 0                            'state: 0 = key, 1 = value, 2 = nested dictionary,
        Dim key As String = "", value As String = "", valueD As Object
        For t = i + 1 To output.Length
            If (output(t).Equals(","c)) Then                'read the character and changes state based on the delimiter, and appends key/value to dictionary if relevant
                If Not key.Equals("") And Not value.Equals("") And state = 1 Then       'if the key and value arent blank add to dictg
                    cleanKey(key)
                    cleanVal(value)

                    nestedDict.Add(key, value)
                Else                                    'if value is blank add the key and nested dict
                    cleanKey(key)

                    nestedDict.Add(key, valueD)
                End If
                key = ""
                value = ""
                state = 0
            ElseIf (output(t).Equals(":"c)) And Not state = 1 Then      'prevents links from breaking
                state = 1
            ElseIf output(t).Equals("{"c) Then                          'detects a nested nested dict
                state = 2
            ElseIf output(t) = "}"c Then                                'detects the end of the nested dict
                i = t

                If Not key.Equals("") And Not value.Equals("") And state = 1 Then   'same as before
                    cleanKey(key)
                    cleanVal(value)

                    nestedDict.Add(key, value)
                Else
                    cleanKey(key)

                    nestedDict.Add(key, valueD)
                End If

                Return nestedDict
            End If

            If state = 0 Then
                key += output(t)                'if state is zero, add to the key
            ElseIf state = 1 Then
                value += output(t)              'if state is one, add to the value
            ElseIf state = 2 Then
                valueD = parse2NDict(t, output)     'if the state is 2 then parse nested nested dict
            End If
        Next
    End Function

    Private Function parse2NDict(ByRef i As Integer, ByVal output As String)
        Dim nestedDict = New Dictionary(Of Object, Object)  'ignore this?? need to change this to a 2d list of dimensions 1, 2 because nested nested dictionaries exist fml

        Dim state As Integer = 0                            'state: 0 = key, 1 = value, 
        Dim key As String = "", value As String = "", valueD As Object

        For t = i + 1 To output.Length
            If (output(t).Equals(","c)) Then            'checks if it should go to state zero and appends to dictionary as needed
                cleanKey(key)
                cleanVal(value)
                nestedDict.Add(key, value)

                key = ""
                value = ""
                state = 0
            ElseIf (output(t).Equals(":"c)) And Not state = 1 Then      'prevents links from breaking
                state = 1
            ElseIf output(t) = "}"c Then            'detects if the nested nested dictionary should end and appends if necessary
                i = t
                cleanKey(key)
                cleanVal(value)
                nestedDict.Add(key, value)
                Return nestedDict
            End If

            If state = 0 Then           'if state is zero adds to key
                key += output(t)
            ElseIf state = 1 Then       'if state is one adds to value
                value += output(t)
            End If
        Next
    End Function

    Private Function cleanKey(ByRef str As String)  'takes the key and returns it as just the text, no ' or ,
        If str(0) = "'"c Then
            str = str.Substring(1, str.Length - 2)
        Else
            str = str.Substring(3, str.Length - 4)
        End If
    End Function

    Private Function cleanVal(ByRef str As String)   'takes the value and returns it as just the text, no ' or ,
        If str.Length <= 2 Then
            str = "error in cleanVal"
        Else
            str = str.Substring(2)
            If str(0).Equals("'"c) Then
                str = str.Substring(1, str.Length - 2)
            End If
        End If
    End Function

    Private Function parseZeroArg(ByVal output As String)
        Return "WIP"
    End Function

    Private Function parseDoubleArg(ByVal output As String)
        Return "WIP"
    End Function
End Class
