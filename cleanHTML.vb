Imports System.Text.RegularExpressions

Public Class cleanHTML

    Public Shared Function Clean(ByVal HTML As String)

        Dim _html As String = HTML

        Try
            RemoveTag(_html, "font")
            RemoveTag(_html, "div")
            RemoveTag(_html, "span")
            RemoveTag(_html, "p style")
            RemoveTag(_html, "..:")
            RemoveTag(_html, ".:")
            RemoveTag(_html, "p")
            RemoveEntity(_html, "nbsp")
            RemoveEntity(_html, "quot")
            RemoveEntity(_html, "apos")
            RemoveEntity(_html, "amp")
            RemoveSection(_html, "style")
            _html = _html.Replace("·", "<li>")
            _html = _html.Replace("<", "&lt;")
            _html = _html.Replace(">", "&gt;")
            Return _html

        Catch ex As Exception
            Return ""
        End Try
        
    End Function

    Private Shared Sub RemoveTag(ByRef html As String, ByVal tag As String)
        Try
            html = Regex.Replace(html, String.Format("<{0}[^>]*>", tag), "", RegexOptions.IgnoreCase)
            html = Regex.Replace(html, String.Format("</{0}[^>]*>", tag), "", RegexOptions.IgnoreCase)
        Catch ex As Exception
            Exit Sub
        End Try

    End Sub

    Private Shared Sub RemoveEntity(ByRef html As String, ByVal tag As String)
        html = html.Replace("&" & tag & ";", "")
    End Sub

    Private Shared Sub RemoveSection(ByRef html As String, ByVal tag As String)
        html = Regex.Replace(html, String.Format("<{0}.+/{0}>", tag), "", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
    End Sub

    Public Shared Function htmlEncode(ByVal Str As String, Optional ByVal chars As String = "'""&") As String
        For i As Integer = 0 To chars.Length - 1
            Str = Str.Replace(chars.Substring(i, 1), String.Format("&#{0};", Asc(chars.Substring(i, 1)).ToString))
        Next
        Return Str
    End Function

    Public Shared Function FixedLen(ByVal Str As String, ByVal len As Integer) As String

        If Str.Length > len Then
            Str = Left(Str, len) & "..."
        Else
            For i As Integer = 0 To CInt((len - Str.Length) / 2)
                Str += " &nbsp;"
            Next
        End If
        Return Str
    End Function

End Class