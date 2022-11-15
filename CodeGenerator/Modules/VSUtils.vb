'**************************************************
' FILE      : vb
' AUTHOR    : Paulo.Santos
' CREATION  : 3/24/2009 9:35:37 PM
' COPYRIGHT : Copyright © 2009
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       TODO: Add file description
'
' Change log:
' 0.1   3/24/2009 9:35:37 PM
'       Paulo.Santos
'       Created.
'***************************************************

Module VSUtils

    Private Const __ProjectKindVB = "{F184B08F-C81C-45F6-A57F-5ABD9991F28F}"
    Private Const __ProjectKindCSharp = "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}"
    Private Const __ProjectKindCpp = "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}"
    Private Const __ProjectKindWeb = "{E24C65DC-7377-472b-9ABA-BC803B73C61A}"
    Private Const __WebAppProjectGuid = "{349c5851-65df-11da-9384-00065b846f21}"

    Friend Enum ProjectKind
        ' Fields
        Cpp = 3
        CSharp = 2
        Unknown = 0
        VB = 1
        Web = 4
    End Enum

    Friend Function GetProjectKind(ByVal project As EnvDTE.Project) As ProjectKind
        If (Not project Is Nothing) Then
            If (project.Kind = __ProjectKindVB) Then
                Return ProjectKind.VB
            End If
            If (project.Kind = __ProjectKindCSharp) Then
                Return ProjectKind.CSharp
            End If
            If (project.Kind = __ProjectKindCpp) Then
                Return ProjectKind.Cpp
            End If
            If (project.Kind = __ProjectKindWeb) Then
                Return ProjectKind.Web
            End If
        End If
        Return ProjectKind.Unknown
    End Function

    Friend Function GetProjectPropertyByName(ByVal project As EnvDTE.Project, ByVal propertyName As String) As Object
        Dim i As Integer = 1
        Do While (i <= project.Properties.Count)
            Dim [property] As EnvDTE.Property = project.Properties.Item(i)
            If (((Not [property] Is Nothing) AndAlso (Not [property].Name Is Nothing)) AndAlso [property].Name.Equals(propertyName)) Then
                Return [property].Value
            End If
            i += 1
        Loop
        Return Nothing
    End Function

End Module
