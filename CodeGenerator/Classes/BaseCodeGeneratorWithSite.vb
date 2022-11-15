'**************************************************
' FILE      : BaseCodeGeneratorWithSite.vb
' AUTHOR    : paulo.santos
' CREATION  : 3/18/2009 4:13:59 PM
' COPYRIGHT : Copyright © 2009
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       TODO: Add file description
'
' Change log:
' 0.1   3/18/2009 4:13:59 PM
'       paulo.santos
'       Created.
'***************************************************

Imports Microsoft.VisualStudio.OLE.Interop
Imports Microsoft.VisualStudio.Shell
Imports System.CodeDom.Compiler
Imports System.Runtime.InteropServices
Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Designer.Interfaces
Imports Microsoft.VisualStudio.Shell.Interop
Imports EnvDTE
Imports VSLangProj
Imports System.Diagnostics.CodeAnalysis
Imports System.Security.Permissions

<ComVisible(True)> _
<Guid("20090320-2612-2511-2201-000000000002")> _
Public MustInherit Class BaseCodeGeneratorWithSite
    Inherits BaseCodeGenerator
    Implements IObjectWithSite
    Implements IDisposable

    Private __Site As Object
    Private __CodeDomProvider As CodeDomProvider
    Private __ServiceProvider As ServiceProvider

    ''' <summary>Initializes an instance of the <see cref="BaseCodeGeneratorWithSite" /> class.
    ''' This is the default constructor for this class.</summary>
    Protected Sub New()
    End Sub

    ''' <summary>The method that does the actual work of generating code given the input file</summary>
    ''' <param name="inputFileContent">File contents as a string</param>
    ''' <returns>The generated code file as a byte-array</returns>
    Public MustOverride Overrides Function GenerateCode(ByVal inputFileContent As String) As Byte()

    ''' <summary>Gets the default extension for this generator</summary>
    ''' <returns>String with the default extension for this generator</returns>
    Public Overrides Function GetDefaultExtension() As String
        Dim codeDom = GetCodeProvider()
        Debug.Assert(codeDom IsNot Nothing, "CodeDomProvider is NULL.")
        Dim extension = codeDom.FileExtension
        If (Not String.IsNullOrEmpty(extension)) Then
            extension = "." & extension.TrimStart(".".ToCharArray())
        End If
        Return extension
    End Function

    ''' <summary>
    ''' GetSite method of IOleObjectWithSite
    ''' </summary>
    ''' <param name="riid">interface to get</param>
    ''' <param name="ppvSite">IntPtr in which to stuff return value</param>
    <SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes", Justification:="Method may be called by COM consumers.")> _
    Public Sub GetSite(ByRef riid As System.Guid, ByRef ppvSite As System.IntPtr) Implements Microsoft.VisualStudio.OLE.Interop.IObjectWithSite.GetSite

        If (__Site Is Nothing) Then
            Throw New COMException("object is not sited", VSConstants.E_FAIL)
        End If

        Dim pUnknownPointer As IntPtr = Marshal.GetIUnknownForObject(__Site)
        Dim intPointer As IntPtr = IntPtr.Zero
        Marshal.QueryInterface(pUnknownPointer, riid, intPointer)

        If (intPointer = IntPtr.Zero) Then
            Throw New COMException("site does not support requested interface", VSConstants.E_NOINTERFACE)
        End If

        ppvSite = intPointer

    End Sub

    ''' <summary>
    ''' SetSite method of IOleObjectWithSite
    ''' </summary>
    ''' <param name="pUnkSite">site for this object to use</param>
    Public Sub SetSite(ByVal pUnkSite As Object) Implements Microsoft.VisualStudio.OLE.Interop.IObjectWithSite.SetSite
        __Site = pUnkSite
        __CodeDomProvider = Nothing
        __ServiceProvider = Nothing
    End Sub

    ''' <summary>
    ''' Demand-creates a ServiceProvider
    ''' </summary>
    Private ReadOnly Property SiteServiceProvider() As ServiceProvider
        Get
            If (__ServiceProvider Is Nothing) Then
                __ServiceProvider = New ServiceProvider(CType(__Site, IServiceProvider))
                Debug.Assert(__ServiceProvider IsNot Nothing, "Unable to get ServiceProvider from site object.")
            End If
            Return __ServiceProvider
        End Get
    End Property

    ''' <summary>
    ''' Method to get a service by its GUID
    ''' </summary>
    ''' <param name="serviceId">GUID of service to retrieve</param>
    ''' <returns>An object that implements the requested service</returns>
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")> _
    Protected Function GetService(ByVal serviceId As Guid) As Object
        Return SiteServiceProvider.GetService(serviceId)
    End Function

    ''' <summary>
    ''' Method to get a service by its Type
    ''' </summary>
    ''' <param name="serviceType">Type of service to retrieve</param>
    ''' <returns>An object that implements the requested service</returns>
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")> _
    Protected Function GetService(ByVal serviceType As Type) As Object
        Return SiteServiceProvider.GetService(serviceType)
    End Function

    ''' <summary>
    ''' Returns a CodeDomProvider object for the language of the project containing
    ''' the project item the generator was called on
    ''' </summary>
    ''' <returns>A CodeDomProvider object</returns>
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification:="Method belongs to base class.")> _
    Protected Overridable Function GetCodeProvider() As CodeDomProvider

        If (__CodeDomProvider Is Nothing) Then
            '*
            '* Query for IVSMDCodeDomProvider/SVSMDCodeDomProvider for this project type
            '*
            Dim provider As IVSMDCodeDomProvider = CType(GetService(GetType(SVSMDCodeDomProvider)), IVSMDCodeDomProvider)
            If (provider IsNot Nothing) Then
                __CodeDomProvider = CType(provider.CodeDomProvider, CodeDomProvider)
            End If
        Else
            '*
            '* In the case where no language specific CodeDom is available, fall back to VB
            '*
            __CodeDomProvider = CodeDomProvider.CreateProvider("VB")
        End If

        Return __CodeDomProvider

    End Function

    ''' <summary>
    ''' Returns the EnvDTE.ProjectItem object that corresponds to the project item the code 
    ''' generator was called on
    ''' </summary>
    ''' <returns>The EnvDTE.ProjectItem of the project item the code generator was called on</returns>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="This method should not throw exceptions.")> _
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")> _
    Protected Function GetProjectItem() As ProjectItem
        Try
            Dim p As Object = GetService(GetType(ProjectItem))
            Debug.Assert(p IsNot Nothing, "Unable to get Project Item.")
            Return CType(p, ProjectItem)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns the EnvDTE.Project object of the project containing the project item the code 
    ''' generator was called on
    ''' </summary>
    ''' <returns>
    ''' The EnvDTE.Project object of the project containing the project item the code generator was called on
    ''' </returns>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="This method should not throw exceptions.")> _
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")> _
    Protected Function GetProject() As Project
        Try
            Return GetProjectItem().ContainingProject
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns the VSLangProj.VSProjectItem object that corresponds to the project item the code 
    ''' generator was called on
    ''' </summary>
    ''' <returns>The VSLangProj.VSProjectItem of the project item the code generator was called on</returns>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="This method should not throw exceptions.")> _
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")> _
    Protected Function GetVSProjectItem() As VSProjectItem
        Try
            Return CType(GetProjectItem().Object, VSProjectItem)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Returns the VSLangProj.VSProject object of the project containing the project item the code
    ''' generator was called on
    ''' </summary>
    ''' <returns>
    ''' The VSLangProj.VSProject object of the project containing the project item 
    ''' the code generator was called on
    ''' </returns>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="This method should not throw exceptions.")> _
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")> _
    Protected Function GetVSProject() As VSProject
        Try
            Return CType(GetProject().Object, VSProject)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Gets the default namespace of the project.
    ''' </summary>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="This method should not throw exceptions.")> _
    Protected ReadOnly Property DefaultNamespace() As String
        Get
            Try
                Return GetProject().Properties.Item("DefaultNamespace").Value
            Catch ex As Exception
                Return String.Empty
            End Try
        End Get
    End Property

#Region " IDisposable Support "

    Private __Disposed As Boolean

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.__Disposed Then
            If disposing Then
                If (__ServiceProvider IsNot Nothing) Then
                    __ServiceProvider.Dispose()
                End If

                If (__CodeDomProvider IsNot Nothing) Then
                    __CodeDomProvider.Dispose()
                End If
            End If
        End If
        Me.__Disposed = True
    End Sub

    ''' <summary>Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.</summary>
    ''' <filterpriority>2</filterpriority>
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
