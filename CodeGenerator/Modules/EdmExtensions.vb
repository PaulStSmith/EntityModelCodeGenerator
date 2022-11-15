'**************************************************
' FILE      : EdmExtensions.vb
' AUTHOR    : paulo.santos
' CREATION  : 3/24/2009 8:21:27 AM
' COPYRIGHT : Copyright © 2009
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       TODO: Add file description
'
' Change log:
' 0.1   3/24/2009 8:21:27 AM
'       paulo.santos
'       Created.
'***************************************************

Imports System.Runtime.CompilerServices
Imports System.Data

Friend Module EdmExtensions

    Private Const __XmlnsGetterAccess As String = "http://schemas.microsoft.com/ado/2006/04/codegeneration:GetterAccess"
    Private Const __XmlnsSetterAccess As String = "http://schemas.microsoft.com/ado/2006/04/codegeneration:GetterAccess"

#Region " EdmProperty Public "

    <Extension()> _
    Public Function IsPublic(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return IsPublicGetter(edmProperty) AndAlso IsPublicSetter(edmProperty)
    End Function

    <Extension()> _
    Public Function IsPublicGetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Dim getterAccess As String = GetAccess(edmProperty, __XmlnsGetterAccess)
        Return (getterAccess = String.Empty OrElse getterAccess = "Public")
    End Function

    <Extension()> _
    Public Function IsPublicSetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Dim setterAccess As String = GetAccess(edmProperty, __XmlnsSetterAccess)
        Return (setterAccess = String.Empty OrElse setterAccess = "Public")
    End Function

#End Region

#Region " EdmProperty Internal "

    <Extension()> _
    Public Function IsInternal(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return IsInternalGetter(edmProperty) AndAlso IsInternalSetter(edmProperty)
    End Function

    <Extension()> _
    Public Function IsInternalGetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return (GetAccess(edmProperty, __XmlnsGetterAccess) = "Internal")
    End Function

    <Extension()> _
    Public Function IsInternalSetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return (GetAccess(edmProperty, __XmlnsSetterAccess) = "Internal")
    End Function

#End Region

#Region " EdmProperty Protected "

    <Extension()> _
    Public Function IsProtected(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return IsProtectedGetter(edmProperty) AndAlso IsProtectedSetter(edmProperty)
    End Function

    <Extension()> _
    Public Function IsProtectedGetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return (GetAccess(edmProperty, __XmlnsGetterAccess) = "Protected")
    End Function

    <Extension()> _
    Public Function IsProtectedSetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return (GetAccess(edmProperty, __XmlnsSetterAccess) = "Protected")
    End Function

#End Region

#Region " EdmProperty Private "

    <Extension()> _
    Public Function IsPrivate(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return IsPrivateGetter(edmProperty) AndAlso IsPrivateSetter(edmProperty)
    End Function

    <Extension()> _
    Public Function IsPrivateGetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return (GetAccess(edmProperty, __XmlnsGetterAccess) = "Private")
    End Function

    <Extension()> _
    Public Function IsPrivateSetter(ByVal edmProperty As Metadata.Edm.EdmProperty) As Boolean
        Return (GetAccess(edmProperty, __XmlnsSetterAccess) = "Private")
    End Function

#End Region

#Region " NavigationProperty Public "

    <Extension()> _
    Public Function IsPublic(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return IsPublicGetter(NavigationProperty) AndAlso IsPublicSetter(NavigationProperty)
    End Function

    <Extension()> _
    Public Function IsPublicGetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Dim getterAccess As String = GetAccess(NavigationProperty, __XmlnsGetterAccess)
        Return (getterAccess = String.Empty OrElse getterAccess = "Public")
    End Function

    <Extension()> _
    Public Function IsPublicSetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Dim getterAccess As String = GetAccess(NavigationProperty, __XmlnsGetterAccess)
        Return (getterAccess = String.Empty OrElse getterAccess = "Public")
    End Function

#End Region

#Region " NavigationProperty Internal "

    <Extension()> _
    Public Function IsInternal(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return IsInternalGetter(NavigationProperty) AndAlso IsInternalSetter(NavigationProperty)
    End Function

    <Extension()> _
    Public Function IsInternalGetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return (GetAccess(NavigationProperty, __XmlnsGetterAccess) = "Internal")
    End Function

    <Extension()> _
    Public Function IsInternalSetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return (GetAccess(NavigationProperty, __XmlnsSetterAccess) = "Internal")
    End Function

#End Region

#Region " NavigationProperty Protected "

    <Extension()> _
    Public Function IsProtected(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return IsProtectedGetter(NavigationProperty) AndAlso IsProtectedSetter(NavigationProperty)
    End Function

    <Extension()> _
    Public Function IsProtectedGetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return (GetAccess(NavigationProperty, __XmlnsGetterAccess) = "Protected")
    End Function

    <Extension()> _
    Public Function IsProtectedSetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return (GetAccess(NavigationProperty, __XmlnsSetterAccess) = "Protected")
    End Function

#End Region

#Region " NavigationProperty Private "

    <Extension()> _
    Public Function IsPrivate(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return IsPrivateGetter(NavigationProperty) AndAlso IsPrivateSetter(NavigationProperty)
    End Function

    <Extension()> _
    Public Function IsPrivateGetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return (GetAccess(NavigationProperty, __XmlnsGetterAccess) = "Private")
    End Function

    <Extension()> _
    Public Function IsPrivateSetter(ByVal NavigationProperty As Metadata.Edm.NavigationProperty) As Boolean
        Return (GetAccess(NavigationProperty, __XmlnsSetterAccess) = "Private")
    End Function

#End Region

#Region " Private Methods "

    Private Function GetAccess(ByVal edm As Metadata.Edm.EdmMember, ByVal identity As String) As String
        Dim access As System.Data.Metadata.Edm.MetadataProperty = Nothing
        edm.MetadataProperties.TryGetValue(identity, True, access)
        Return If(access Is Nothing, Nothing, access.Value)
    End Function

#End Region

End Module
