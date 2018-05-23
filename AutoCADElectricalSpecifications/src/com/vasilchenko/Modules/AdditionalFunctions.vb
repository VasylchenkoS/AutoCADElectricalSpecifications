Imports System.Text.RegularExpressions

Namespace com.vasilchenko.Modules
    Module AdditionalFunctions
        Public Function GetNumericFromString(s As String) As String
            If IsNothing(s) Then Return -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                Return matches(matches.Count - 1).Value
            Else
                Return -1
            End If
        End Function
    End Module
End Namespace
