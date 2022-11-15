'**************************************************
' FILE      : BaseCodeGenerator.vb
' AUTHOR    : paulo.santos
' CREATION  : 3/18/2009 2:30:02 PM
' COPYRIGHT : Copyright © 2009
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       TODO: Add file description
'
' Change log:
' 0.1   3/18/2009 2:30:02 PM
'       paulo.santos
'       Created.
'***************************************************

Imports Microsoft.VisualStudio
Imports Microsoft.VisualStudio.Shell.Interop
Imports System.Runtime.InteropServices
Imports System.Diagnostics.CodeAnalysis
Imports System.Security.Permissions

<ComVisible(True)> _
<Guid("20090320-2612-2511-2201-000000000001")> _
Public MustInherit Class BaseCodeGenerator
    Implements IVsSingleFileGenerator

    Private __CodeGeneratorProgress As IVsGeneratorProgress
    Private __CodeFileNameSpace As String = String.Empty
    Private __CodeFilePath As String = String.Empty

    ''' <summary>Initializes an instance of the <see cref="BaseCodeGenerator" /> class.
    ''' This is the default constructor for this class.</summary>
    Protected Sub New()
    End Sub

    ''' <summary>
    ''' Implements the IVsSingleFileGenerator.DefaultExtension method. 
    ''' Returns the extension of the generated file
    ''' </summary>
    ''' <param name="pbstrDefaultExtension">Out parameter, will hold the extension that is to be given to the output file name. The returned extension must include a leading period</param>
    ''' <returns>S_OK if successful, E_FAIL if not</returns>
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="Method should not thow exceptions.")> _
    Public Function DefaultExtension(ByRef pbstrDefaultExtension As String) As Integer Implements Microsoft.VisualStudio.Shell.Interop.IVsSingleFileGenerator.DefaultExtension

        Try
            pbstrDefaultExtension = GetDefaultExtension()
            Return VSConstants.S_OK
        Catch ex As Exception
            Trace.WriteLine(ex.Message)
            pbstrDefaultExtension = String.Empty
            Return VSConstants.E_FAIL
        End Try

    End Function

    '''  <summary>
    '''  Implements the IVsSingleFileGenerator.Generate method.
    '''  Executes the transformation and returns the newly generated output file, whenever a custom tool is loaded, or the input file is saved
    '''  </summary>
    '''  <param name="wszInputFilePath">The full path of the input file. May be a null reference (Nothing in Visual Basic) in future releases of Visual Studio, so generators should not rely on this value</param>
    '''  <param name="bstrInputFileContents">The contents of the input file. This is either a UNICODE BSTR (if the input file is text) or a binary BSTR (if the input file is binary). If the input file is a text file, the project system automatically converts the BSTR to UNICODE</param>
    '''  <param name="wszDefaultNamespace">This parameter is meaningful only for custom tools that generate code. It represents the namespace into which the generated code will be placed. If the parameter is not a null reference (Nothing in Visual Basic) and not empty, the custom tool can use the following syntax to enclose the generated code</param>
    '''  <param name="rgbOutputFileContents">[out] Returns an array of bytes to be written to the generated file. You must include UNICODE or UTF-8 signature bytes in the returned byte array, as this is a raw stream. The memory for rgbOutputFileContents must be allocated using the .NET Framework call, System.Runtime.InteropServices.AllocCoTaskMem, or the equivalent Win32 system call, CoTaskMemAlloc. The project system is responsible for freeing this memory</param>
    '''  <param name="pcbOutput">[out] Returns the count of bytes in the rgbOutputFileContents array</param>
    '''  <param name="pGenerateProgress">A reference to the IVsGeneratorProgress interface through which the generator can report its progress to the project system</param>
    '''  <returns>If the method succeeds, it returns S_OK. If it fails, it returns E_FAIL</returns>
    Public Function Generate(ByVal wszInputFilePath As String, ByVal bstrInputFileContents As String, ByVal wszDefaultNamespace As String, ByVal rgbOutputFileContents() As System.IntPtr, ByRef pcbOutput As UInteger, ByVal pGenerateProgress As Microsoft.VisualStudio.Shell.Interop.IVsGeneratorProgress) As Integer Implements Microsoft.VisualStudio.Shell.Interop.IVsSingleFileGenerator.Generate

        If (String.IsNullOrEmpty(bstrInputFileContents)) Then
            Throw New ArgumentNullException("bstrInputFileContents")
        End If

        __CodeFilePath = wszInputFilePath
        __CodeFileNameSpace = wszDefaultNamespace
        __CodeGeneratorProgress = pGenerateProgress

        Dim bytes = GenerateCode(bstrInputFileContents)

        If (bytes Is Nothing) OrElse (bytes.Length = 0) Then
            '*
            '* This signals that GenerateCode() has failed. Tasklist items have been put up in GenerateCode()
            '*
            rgbOutputFileContents = Nothing
            pcbOutput = 0

            '*
            '* Return E_FAIL to inform Visual Studio that the generator has failed (so that no file gets generated)
            '*
            Return VSConstants.E_FAIL
        End If

        '*
        '* The contract between IVsSingleFileGenerator implementors and consumers is that
        '* any output returned from IVsSingleFileGenerator.Generate() is returned through
        '* memory allocated via CoTaskMemAlloc(). Therefore, we have to convert the 
        '* byte[] array returned from GenerateCode() into an unmanaged blob.  
        '*
        Dim outputLength As Integer = bytes.Length
        rgbOutputFileContents(0) = Marshal.AllocCoTaskMem(outputLength)
        Marshal.Copy(bytes, 0, rgbOutputFileContents(0), outputLength)
        pcbOutput = CUInt(outputLength)
        Return VSConstants.S_OK

    End Function

    ''' <summary>
    ''' Namespace for the file
    ''' </summary>
    Protected ReadOnly Property FileNamespace() As String
        Get
            Return __CodeFileNameSpace
        End Get
    End Property

    ''' <summary>
    ''' File-path for the input file
    ''' </summary>
    Protected ReadOnly Property InputFilePath() As String
        Get
            Return __CodeFilePath
        End Get
    End Property

    ''' <summary>
    ''' Method that will communicate an error via the shell callback mechanism
    ''' </summary>
    ''' <param name="level">Level or severity</param>
    ''' <param name="message">Text displayed to the user</param>
    ''' <param name="line">Line number of error</param>
    ''' <param name="column">Column number of error</param>
    <SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="result", Justification:="Variable used to avoid memory leak.")> _
    Protected Overridable Sub GeneratorError(ByVal level As Integer, ByVal message As String, ByVal line As Integer, ByVal column As Integer)
        Dim progress As IVsGeneratorProgress = CodeGeneratorProgress
        If (progress IsNot Nothing) Then
            Dim result = progress.GeneratorError(0, CUInt(level), message, CUInt(line), CUInt(column))
        End If
    End Sub

    ''' <summary>
    ''' Method that will communicate a warning via the shell callback mechanism
    ''' </summary>
    ''' <param name="level">Level or severity</param>
    ''' <param name="message">Text displayed to the user</param>
    ''' <param name="line">Line number of warning</param>
    ''' <param name="column">Column number of warning</param>
    <SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="result", Justification:="Variable used to avoid memory leak.")> _
    Protected Overridable Sub GeneratorWarning(ByVal level As Integer, ByVal message As String, ByVal line As Integer, ByVal column As Integer)
        Dim progress As IVsGeneratorProgress = CodeGeneratorProgress
        If (progress IsNot Nothing) Then
            Dim result = progress.GeneratorError(1, CUInt(level), message, CUInt(line), CUInt(column))
        End If
    End Sub

    ''' <summary>
    ''' Interface to the VS shell object we use to tell our progress while we are generating
    ''' </summary>
    Protected Friend ReadOnly Property CodeGeneratorProgress() As IVsGeneratorProgress
        Get
            Return __CodeGeneratorProgress
        End Get
    End Property

    ''' <summary>
    ''' Gets the default extension for this generator
    ''' </summary>
    ''' <returns>String with the default extension for this generator</returns>
    <SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate", Justification:="Method must be implemented by derived classes.")> _
    Public MustOverride Function GetDefaultExtension() As String

    ''' <summary>
    ''' The method that does the actual work of generating code given the input file
    ''' </summary>
    ''' <param name="inputFileContent">File contents as a string</param>
    ''' <returns>The generated code file as a byte-array</returns>
    Public MustOverride Function GenerateCode(ByVal inputFileContent As String) As Byte()

End Class
