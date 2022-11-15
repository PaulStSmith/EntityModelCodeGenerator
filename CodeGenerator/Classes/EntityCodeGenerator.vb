'**************************************************
' FILE      : PJonDevelopmentEntityCodeGenerator.vb
' AUTHOR    : paulo.santos
' CREATION  : 3/18/2009 4:36:54 PM
' COPYRIGHT : Copyright © 2009
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       TODO: Add file description
'
' Change log:
' 0.1   3/18/2009 4:36:54 PM
'       paulo.santos
'       Created.
'***************************************************

Imports System.Runtime.InteropServices
Imports System.CodeDom
Imports System.Data.Metadata.Edm
Imports System.Globalization
Imports System.ComponentModel
Imports System.Data.Entity.Design
Imports System.IO
Imports System.Diagnostics.CodeAnalysis
Imports System.Collections.ObjectModel

''' <summary>
''' This is the generator class. 
''' When the 'Custom Tool' property of a .edmx file in a C# or VB project item is set to "EntityCodeGenerator",
''' Visual Studio will call the GenerateCode function which will return the contents of the generated file back
''' to the project system
''' </summary>
<SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling", justification:="This class is quite complex, due the GenerateCode method.")> _
<ComVisible(True)> _
<Guid("20090320-2612-2511-2201-000000000003")> _
Public Class EntityCodeGenerator
    Inherits BaseCodeGeneratorWithSite

    Private __Types As New Dictionary(Of EntityType, CodeTypeDeclaration)
    Private __StructCodeNamespace As CodeNamespace
    Private __CodeNamespace As String = String.Empty

    ''' <summary>Gets the default extension for this generator</summary>
    ''' <returns>String with the default extension for this generator</returns>
    Public Overrides Function GetDefaultExtension() As String
        Return ".Designer" & MyBase.GetDefaultExtension()
    End Function

    ''' <summary>The method that does the actual work of generating code given the input file</summary>
    ''' <param name="inputFileContent">File contents as a string</param>
    ''' <returns>The generated code file as a byte-array</returns>
    <SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling", Justification:="It's a complex code.")> _
    <SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId:="result", Justification:="Using H_RESULT to avoid memory leak.")> _
    <SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Justification:="This method should not throw any exception.")> _
    Public Overrides Function GenerateCode(ByVal inputFileContent As String) As Byte()

        Dim generatedCodeAsBytes As Byte() = Nothing

        Try
            '*
            '* Gets the CSDL contents from the file
            '*
            Dim csdlContent = ExtractCsdlContent(inputFileContent)
            If (csdlContent Is Nothing) Then
                Throw New InvalidOperationException("No CSDL content in input file")
            End If

            '*
            '* Sets the code namespace
            '*
            SetCodeNamespace(csdlContent)

            '*
            '* Prepare the XMLReader
            '*
            Dim csdlReader = csdlContent.CreateReader()
            Using csdlReader
                '*
                '* Prepare the CodeWriter
                '*
                Dim codeWriter As New StringWriter(CultureInfo.InvariantCulture)
                Using codeWriter
                    '*
                    '* Define the Code Language, based on the file Extension
                    '*
                    Dim lang As LanguageOption = LanguageOption.GenerateVBCode
                    Dim fileExtension = MyBase.GetCodeProvider().FileExtension
                    If (fileExtension IsNot Nothing AndAlso fileExtension.Length > 0) Then
                        fileExtension = "." + fileExtension.TrimStart(".".ToCharArray())
                    End If

                    '*
                    '* Check if it's a valid language
                    '*
                    If (fileExtension.EndsWith(".vb", StringComparison.OrdinalIgnoreCase)) Then
                        lang = LanguageOption.GenerateVBCode
                    ElseIf (fileExtension.EndsWith(".cs", StringComparison.OrdinalIgnoreCase)) Then
                        lang = LanguageOption.GenerateCSharpCode
                    Else
                        Throw New InvalidOperationException("Unsupported project language. Only C# and VB are supported.")
                    End If

                    '*
                    '* Reports as 25% done
                    '*
                    If (MyBase.CodeGeneratorProgress IsNot Nothing) Then
                        Dim result = MyBase.CodeGeneratorProgress.Progress(25, 100)
                    End If

                    '*
                    '* Create the CodeGenerator and hookup the proper events
                    '*
                    Dim classGenerator = New EntityClassGenerator(lang)
                    classGenerator.EdmToObjectNamespaceMap.Add(csdlContent.Attribute("Namespace").Value, __CodeNamespace)
                    AddHandler classGenerator.OnTypeGenerated, AddressOf OnTypeGenerated

                    '*
                    '* Reports as 50% done
                    '*
                    If (MyBase.CodeGeneratorProgress IsNot Nothing) Then
                        Dim result = MyBase.CodeGeneratorProgress.Progress(50, 100)
                    End If

                    '*
                    '* Clean the previous namespace
                    '*
                    __Types.Clear()
                    __StructCodeNamespace = Nothing

                    '*
                    '* Generate the code
                    '*
                    Dim errors = classGenerator.GenerateCode(csdlReader, codeWriter)

                    '*
                    '* Reports the errors to the Visual Studio
                    '*
                    If (errors IsNot Nothing) Then
                        For Each ex In errors
                            Dim line As UInteger = If(ex.Line = 0, 0, ex.Line - 1)
                            Dim column As UInteger = If(ex.Column = 0, 0, ex.Column - 1)

                            If (ex.Severity = Metadata.Edm.EdmSchemaErrorSeverity.Warning) Then
                                MyBase.GeneratorWarning(0, ex.Message, line, column)
                            Else
                                MyBase.GeneratorError(4, ex.Message, line, column)
                            End If
                        Next
                    End If

                    '*
                    '* Generate the TypeConverter
                    '*
                    GenerateConverters(codeWriter)

                    '*
                    '* Reports as 75% done
                    '*
                    If (MyBase.CodeGeneratorProgress IsNot Nothing) Then
                        Dim result = MyBase.CodeGeneratorProgress.Progress(75, 100)
                    End If

                    '*
                    '* Replace the partial placeholders
                    '*
                    Dim generatedCode As String = Text.RegularExpressions.Regex.Replace(codeWriter.ToString(), "\r?\n\s*?(?:'|//)\#\#Partial\#\#\r?\n(.*?)private", vbCrLf & "$1" & If(lang = LanguageOption.GenerateCSharpCode, "partial private", "Partial Private"), Text.RegularExpressions.RegexOptions.Singleline Or Text.RegularExpressions.RegexOptions.IgnoreCase)

                    '*
                    '* Converts the generated code into bytes
                    '*
                    generatedCodeAsBytes = Text.Encoding.UTF8.GetBytes(generatedCode)

                    '*
                    '* Reports as 100% done
                    '*
                    If (MyBase.CodeGeneratorProgress IsNot Nothing) Then
                        Dim result = MyBase.CodeGeneratorProgress.Progress(100, 100)
                    End If

                End Using
            End Using
        Catch ex As Exception
            MyBase.GeneratorError(4, ex.Message, 1, 1)
            Debug.WriteLine(ex.Message)
            Debug.WriteLine(ex.StackTrace)
            generatedCodeAsBytes = Nothing
        End Try

        Return generatedCodeAsBytes

    End Function

    ''' <summary>
    ''' This method is called when a type is created by the code generator.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The arguments of the event.</param>
    Private Sub OnTypeGenerated(ByVal sender As Object, ByVal e As TypeGeneratedEventArgs)

        If (e.TypeSource.BuiltInTypeKind <> BuiltInTypeKind.EntityType) Then
            Exit Sub
        End If

        If (__StructCodeNamespace Is Nothing) Then
            __StructCodeNamespace = New CodeNamespace(GetNamespaceName(__STR_Structures))
        End If

        Dim key = e.TypeSource
        Dim ctm As CodeTypeDeclaration = GenerateStructureFromEntity(key)
        If (ctm IsNot Nothing) Then
            __Types.Add(key, ctm)
            __StructCodeNamespace.Types.Add(ctm)
            e.AdditionalMembers.AddRange(GenerateToStructureMethod(key))
        End If

    End Sub

    ''' <summary>
    ''' Generates a <see cref="CodeTypeMember"/> that contains an structure based on the specified <paramref name="edmSource"/>.
    ''' </summary>
    ''' <param name="edmSource">The source of the generated structure.</param>
    ''' <returns>A <see cref="CodeTypeMember"/> that contains an structure based on the specified <paramref name="edmSource"/>.</returns>
    Private Function GenerateStructureFromEntity(ByVal edmSource As MetadataItem) As CodeTypeMember

        '*
        '* Check if the entitySource has a name
        '*
        Dim nameProperty = edmSource.MetadataProperties("Name")
        Dim membersProperty = edmSource.MetadataProperties("Members")
        If (nameProperty Is Nothing OrElse membersProperty Is Nothing) Then
            '*
            '* No name, no code generation
            '*
            Return Nothing
        End If

        '*
        '* Create the Structure
        '*
        Dim entityName As String = nameProperty.Value
        Dim struct As New CodeTypeDeclaration(entityName) ' & STR_Structure)
        With struct
            .IsClass = True
            .IsPartial = True
            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(SerializableAttribute))))
            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(TypeConverter)), New CodeAttributeArgument(New CodeTypeOfExpression("Converters." & entityName & __STR_Converter))))

            '*
            '* Iterates through every property
            '*
            Dim properties = GetPrimitiveProperties(edmSource)
            If (properties IsNot Nothing) Then
                For Each prop As EdmProperty In properties
                    If (prop.IsPublicGetter OrElse prop.IsInternalGetter) Then
                        AddPrimitiveProperty(struct, prop)
                    End If
                Next
            End If

            '*
            '* Iterates through every navigation properties
            '*
            Dim navProperties = GetNavigationProperties(edmSource)
            If (navProperties IsNot Nothing) Then
                For Each nav As NavigationProperty In navProperties
                    If (nav.IsPublicGetter OrElse nav.IsInternalGetter) Then
                        AddNavigationalProperty(struct, nav)
                    End If
                Next
            End If
        End With

        Return struct

    End Function

    Private Function GetStructType(ByVal entityName As String) As CodeTypeReference
        Return New CodeTypeReference(GetNamespaceName(__STR_Structures) & "." & entityName)
    End Function

    ''' <summary>
    ''' Returns a <see cref="CodeTypeMember"/> that contains the method ToStructure().
    ''' </summary>
    ''' <param name="edmSource">The source of the generated method.</param>
    ''' <returns>A <see cref="CodeTypeMember"/> that contains the method ToStructure().</returns>
    Private Function GenerateToStructureMethod(ByVal edmSource As MetadataItem) As CodeTypeMember()

        Dim nameProperty = edmSource.MetadataProperties("Name")
        If (nameProperty Is Nothing) Then
            '*
            '* No name, no code generation
            '*
            Return Nothing
        End If

        Dim entityName As String = nameProperty.Value

        Dim toStructureMethod As New CodeMemberMethod()
        Dim structType = GetStructType(entityName)

        With toStructureMethod
            .Name = __STR_ToStructure
            .Attributes = MemberAttributes.Public
            .ReturnType = structType

            .Statements.Add(New CodeMethodReturnStatement(New CodeMethodInvokeExpression(New CodeThisReferenceExpression(), __STR_ToStructure, New CodePrimitiveExpression(1))))

            .Comments.Add(New CodeCommentStatement("<summary>Converts this instance of the entity to a simple class that is detached from any database.</summary>", True))
        End With

        Dim toStructureMethod2 As New CodeMemberMethod()
        With toStructureMethod2
            .Name = __STR_ToStructure
            .Attributes = MemberAttributes.Public
            .ReturnType = structType
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(Integer), "numLevels"))

            .Statements.Add(New CodeVariableDeclarationStatement(GetType(TypeConverter), "structConverter", New CodeObjectCreateExpression(New CodeTypeReference("Converters." & entityName & __STR_Converter), New CodeArgumentReferenceExpression("numLevels"))))
            .Statements.Add(New CodeMethodReturnStatement(New CodeMethodInvokeExpression(New CodeVariableReferenceExpression("structConverter"), "ConvertFrom", New CodeThisReferenceExpression())))

            .Comments.Add(New CodeCommentStatement("<summary>Converts this instance of the entity to a simple class that is detached from any database, specifying the number of levels to drill down in the entity hierarchy.</summary>", True))
            .Comments.Add(New CodeCommentStatement("<param name=""numLevels"">Number of levels to read down the entity hierarchy.</param>", True))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(SuppressMessageAttribute)), _
                                                               New CodeAttributeArgument(New CodePrimitiveExpression("Microsoft.Naming")), _
                                                               New CodeAttributeArgument(New CodePrimitiveExpression("CA1704:IdentifiersShouldBeSpelledCorrectly")), _
                                                               New CodeAttributeArgument("MessageId", New CodePrimitiveExpression("num"))))
        End With

        Return New CodeTypeMember() {toStructureMethod, toStructureMethod2}

    End Function

    ''' <summary>
    ''' Generates TypeConverter classes for each type creaded.
    ''' </summary>
    ''' <param name="writer">The <see cref="IO.TextWriter"></see> that will recieve the generated code.</param>
    Private Sub GenerateConverters(ByVal writer As IO.TextWriter)

        If (__Types Is Nothing) OrElse (__Types.Count = 0) Then
            Return
        End If

        Dim ns As New CodeNamespace(GetNamespaceName("Converters"))

        '*
        '* Gets the namespace for the converters
        '*
        For Each edmSource As EntityType In __Types.Keys
            Dim struct = __Types(edmSource)
            Dim cvt As New CodeTypeDeclaration(edmSource.Name & __STR_Converter)
            With cvt
                '*
                '* Creates the type delcaration
                '*
                Dim structType = GetStructType(edmSource.Name)

                .IsPartial = True
                .BaseTypes.Add(New CodeTypeReference(GetType(TypeConverter)))
                .Comments.Add(New CodeCommentStatement(String.Format(CultureInfo.InvariantCulture, "<summary>Provides methods to convert from {0} to {1}.</summary>", edmSource.Name, struct.Name), True))

                '*
                '* Create the variable to hold the numbers of levels down to convert (default 1)
                '*
                .Members.Add(New CodeMemberField(GetType(Integer), "__NumLevels"))

                '*
                '* Create the default constructor
                '*
                .Members.Add(GetConverterDefaultCtor(.Name))

                '*
                '* Create the default constructor
                '*
                .Members.Add(GetCtor(.Name))

                '*
                '* Create the NumLevels property
                '*
                .Members.Add(GetNumLevelsProperty())

                '*
                '* Overrides the CanConvertFrom method
                '*
                .Members.Add(GetCanConvertFromMethod(edmSource))

                '*
                '* Overrides the CanConvertTo method
                '*
                .Members.Add(GetCanConvertToMethod(structType.BaseType))

                '*
                '* Overrides the ConvertFrom method
                '*
                .Members.Add(GetConvertFromMethod(edmSource, structType))

                '*
                '* Create the partial method
                '*
                .Members.Add(GetAdditionalConversionMethod(edmSource, structType))
            End With
            ns.Types.Add(cvt)
        Next

        With MyBase.GetCodeProvider()
            '*
            '* Writes the Structures generated earlier
            '*
            .GenerateCodeFromNamespace(__StructCodeNamespace, writer, Nothing)

            '*
            '* Writes the code to the code writer
            '*
            .GenerateCodeFromNamespace(ns, writer, Nothing)
        End With

    End Sub

    ''' <summary>
    ''' Gets a Namespace name with the specified name.
    ''' </summary>
    ''' <param name="name">The name of the new namespace.</param>
    ''' <returns>The name of the new namespace.</returns>
    ''' <remarks></remarks>
    Private Function GetNamespaceName(ByVal name As String) As String
        Return (__CodeNamespace & "." & Me.GetCodeProvider().CreateEscapedIdentifier(name)).Trim(".".ToCharArray())
    End Function

    ''' <summary>
    ''' Adds a primitive property in specified structure. 
    ''' </summary>
    ''' <param name="struct">The structure to add the primitive property.</param>
    ''' <param name="primProperty">The primitive property to add.</param>
    Private Sub AddPrimitiveProperty(ByVal struct As CodeTypeDeclaration, ByVal primProperty As EdmProperty)

        '*
        '* Check if the getter is either public or internal
        '*
        If (Not (primProperty.IsPublicGetter OrElse primProperty.IsInternalGetter)) Then
            Return
        End If

        '*
        '* Check if the type is a primitive type
        '*
        Dim primType = TryCast(primProperty.TypeUsage.EdmType, PrimitiveType)
        If primType Is Nothing Then
            MyBase.GeneratorWarning(0, "Property " & primProperty.Name & " is of type " & primProperty.TypeUsage.EdmType.FullName & " which is not a PrimitiveType. This property is being ignored when generating the structure type.", 0, 0)
            Exit Sub
        End If

        '*
        '* Add the property field
        '*
        struct.Members.Add(New CodeMemberField(primType.ClrEquivalentType, "__" & primProperty.Name))

        '*
        '* Add the property getter and setter
        '*
        Dim propMember As New CodeMemberProperty()
        With propMember
            .HasGet = True
            .HasSet = True
            .Attributes = MemberAttributes.Public Or MemberAttributes.Final
            .Name = Me.GetCodeProvider().CreateEscapedIdentifier(primProperty.Name)
            .Type = New CodeTypeReference(primType.ClrEquivalentType)

            AddGetterAndSetterToProperty(primProperty, propMember)
        End With
        AddXmlSummary(primProperty, propMember)
        struct.Members.Add(propMember)

    End Sub

    ''' <summary>
    ''' Adds a naviagational property to the specified strucutre.
    ''' </summary>
    ''' <param name="struct">The structure to add the navigational property.</param>
    ''' <param name="navProperty">The navigational property to add.</param>
    Private Sub AddNavigationalProperty(ByVal struct As CodeTypeDeclaration, ByVal navProperty As NavigationProperty)

        '*
        '* Check if the getter is either public or internal
        '*
        If (Not (navProperty.IsPublicGetter OrElse navProperty.IsInternalGetter)) Then
            Return
        End If

        '*
        '* Add the memberField
        '*
        Dim isCollection = (TypeOf navProperty.TypeUsage.EdmType Is CollectionType) ' navToMember.RelationshipMultiplicity = RelationshipMultiplicity.Many)
        Dim navType = If(isCollection, DirectCast(navProperty.TypeUsage.EdmType, CollectionType).TypeUsage.EdmType, navProperty.TypeUsage.EdmType)
        Dim baseType As New CodeTypeReference(GetNamespaceName(__STR_Structures) & "." & navType.Name)
        Dim typeRef = baseType
        If isCollection Then
            typeRef = New CodeTypeReference(typeRef, 1)
        End If
        struct.Members.Add(New CodeMemberField(typeRef, "__" & navProperty.Name))

        '*
        '* Add the property getter and setter
        '*
        Dim propMember As New CodeMemberProperty()
        With propMember
            .Attributes = MemberAttributes.Public Or MemberAttributes.Final
            .Name = Me.GetCodeProvider().CreateEscapedIdentifier(navProperty.Name)

            '*
            '* Define the proper type for the property
            '*
            If (Not isCollection) Then
                .Type = typeRef
                AddGetterAndSetterToProperty(navProperty, propMember)
            Else
                .Type = GetGenericTypeReference(GetType(ReadOnlyCollection(Of Object)), baseType)

                '*
                '* Create the return statement
                '*
                .HasSet = False
                .GetStatements.Add(New CodeMethodReturnStatement(New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeTypeReferenceExpression(GetType(Array)), "AsReadOnly"), New CodeExpression() {New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__" & navProperty.Name)})))

                '*
                '* Identifies that it's an automatically generated code
                '*
                .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
            End If
        End With
        AddXmlSummary(navProperty, propMember)
        struct.Members.Add(propMember)

        If (isCollection) Then
            AddSetterMethodForCollection(struct, navProperty, baseType)
        End If

    End Sub

    ''' <summary>
    ''' Returns a <see cref="CodeMemberMethod"/> that overrides the <see cref="TypeConverter.ConvertFrom"/> method.
    ''' </summary>
    ''' <param name="edmSource">The <see cref="EntityType"/> that is the source of the Conversion.</param>
    ''' <param name="structType">A <see cref="CodeTypeReference"/> for the structure type.</param>
    ''' <returns>A <see cref="CodeMemberMethod"/> that overrides the <see cref="TypeConverter.ConvertFrom"/> method.</returns>
    <SuppressMessage("Microsoft.Maintainability", "CA1506:AvoidExcessiveClassCoupling", Justification:="This is a complex method.")> _
    Private Function GetConvertFromMethod(ByVal edmSource As EntityType, ByVal structType As CodeTypeReference) As CodeMemberMethod

        Dim convertFromMethod As New CodeMemberMethod
        With convertFromMethod
            .Name = "ConvertFrom"
            .Attributes = MemberAttributes.Override Or MemberAttributes.Public
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(ITypeDescriptorContext), "context"))
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(CultureInfo), "culture"))
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(Object), "value"))
            .ReturnType = New CodeTypeReference(GetType(Object))

            '*
            '* Check if it's the proper type
            '*
            .Statements.Add(New CodeConditionStatement(New CodeBinaryOperatorExpression(New CodeArgumentReferenceExpression("value"), CodeBinaryOperatorType.IdentityEquality, New CodePrimitiveExpression(Nothing)), New CodeMethodReturnStatement(New CodePrimitiveExpression(Nothing))))
            Dim ifStatement = New CodeConditionStatement(New CodeMethodInvokeExpression(New CodeThisReferenceExpression(), "CanConvertFrom", New CodeMethodInvokeExpression(New CodeArgumentReferenceExpression("value"), "GetType")))
            With ifStatement.TrueStatements
                .Add(New CodeVariableDeclarationStatement(structType, "struct", New CodeObjectCreateExpression(structType)))

                '*
                '* List all the properties
                '*
                Dim properties = GetPrimitiveProperties(edmSource)
                Dim structReference As CodeVariableReferenceExpression = New CodeVariableReferenceExpression("struct")
                If (properties IsNot Nothing) Then
                    For Each prop As EdmProperty In properties
                        If (prop.IsPublicGetter OrElse prop.IsInternalGetter) Then
                            .Add(New CodeAssignStatement(New CodePropertyReferenceExpression(structReference, Me.GetCodeProvider().CreateEscapedIdentifier(prop.Name)), _
                                                         New CodePropertyReferenceExpression(New CodeArgumentReferenceExpression("value"), Me.GetCodeProvider().CreateEscapedIdentifier(prop.Name))))
                        End If
                    Next
                End If

                '*
                '* List all the naviagional properties
                '*
                Dim navProperties = GetNavigationProperties(edmSource)
                If (navProperties IsNot Nothing) Then
                    Dim listConverters As New Dictionary(Of String, CodeTypeReference)

                    '*
                    '* Check if it's to go down into the navigational three
                    '*
                    Dim ifDown As New CodeConditionStatement()
                    ifDown.Condition = New CodeBinaryOperatorExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__NumLevels"), CodeBinaryOperatorType.GreaterThanOrEqual, New CodePrimitiveExpression(1))
                    With ifDown.TrueStatements
                        For Each nav As NavigationProperty In navProperties
                            Debug.WriteLine(edmSource.FullName & " --> " & nav.TypeUsage.EdmType.FullName)

                            Dim isCollection = (TypeOf nav.TypeUsage.EdmType Is CollectionType)
                            Dim navType = If(isCollection, DirectCast(nav.TypeUsage.EdmType, CollectionType).TypeUsage.EdmType, nav.TypeUsage.EdmType)

                            If (nav.IsPublicGetter OrElse nav.IsInternalGetter) Then
                                '*
                                '* Declare the TypeConverter
                                '*
                                If (Not listConverters.ContainsKey(navType.Name)) Then
                                    Dim typeRef = New CodeTypeReference(navType.Name & "Converter")
                                    listConverters.Add(navType.Name, typeRef)
                                    .Add(New CodeVariableDeclarationStatement(typeRef, "structConverterFor" & navType.Name, New CodeObjectCreateExpression(typeRef, New CodeBinaryOperatorExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__NumLevels"), CodeBinaryOperatorType.Subtract, New CodePrimitiveExpression(1)))))
                                End If

                                '*
                                '* Determine the type of Conversion
                                '*
                                Dim valueNavProperty As CodePropertyReferenceExpression = New CodePropertyReferenceExpression(New CodeArgumentReferenceExpression("value"), nav.Name)
                                If (isCollection) Then
                                    '*
                                    '* Convert a Multiple Items
                                    '*
                                    Dim listStructure As CodeTypeReference = GetGenericTypeReference(GetType(List(Of Object)), structType)
                                    .Add(New CodeVariableDeclarationStatement(listStructure, "item" & nav.Name, New CodeObjectCreateExpression(listStructure)))

                                    '*
                                    '* Get the enumerator
                                    '*
                                    .Add(New CodeVariableDeclarationStatement(GetType(IEnumerator), "enum" & nav.Name, New CodeMethodInvokeExpression(valueNavProperty, "GetEnumerator")))

                                    '*
                                    '* Build the while loop
                                    '*
                                    Dim whileStatement = New CodeIterationStatement()
                                    With whileStatement
                                        .TestExpression = New CodeMethodInvokeExpression(New CodeVariableReferenceExpression("enum" & nav.Name), "MoveNext")
                                        .InitStatement = New CodeSnippetStatement("")
                                        .IncrementStatement = New CodeSnippetStatement("")

                                        '*
                                        '* Convert each item
                                        '*
                                        .Statements.Add(New CodeMethodInvokeExpression( _
                                                            New CodeVariableReferenceExpression("item" & nav.Name), "Add", _
                                                                New CodeMethodInvokeExpression(New CodeVariableReferenceExpression("structConverterFor" & navType.Name), "ConvertFrom", _
                                                                                                   New CodePropertyReferenceExpression(New CodeVariableReferenceExpression("enum" & nav.Name), "Current"))))
                                    End With
                                    .Add(whileStatement)
                                    .Add(New CodeMethodInvokeExpression(structReference, "Set" & nav.Name, New CodeVariableReferenceExpression("item" & nav.Name)))
                                Else
                                    '*
                                    '* Convert a single Item
                                    '*
                                    .Add(New CodeAssignStatement(New CodePropertyReferenceExpression(structReference, nav.Name), _
                                                                 New CodeMethodInvokeExpression(New CodeVariableReferenceExpression("structConverterFor" & navType.Name), "ConvertFrom", valueNavProperty)))
                                End If
                            End If
                        Next

                    End With
                    .Add(ifDown)
                End If

                '*
                '* Provides aditional Conversion
                '*
                .Add(New CodeMethodInvokeExpression(New CodeThisReferenceExpression(), "AdditionalConversion", New CodeArgumentReferenceExpression("value"), New CodeVariableReferenceExpression("struct")))

                '*
                '* Return the converted structure
                '*
                .Add(New CodeMethodReturnStatement(structReference))
            End With
            .Statements.Add(ifStatement)
            .Statements.Add(New CodeMethodReturnStatement(New CodePrimitiveExpression(Nothing)))

            '*
            '* Add the XML comment
            '*
            .Comments.Add(New CodeCommentStatement(" <summary>Converts the given object to the type of this converter, using the specified context and culture information.</summary>", True))
            .Comments.Add(New CodeCommentStatement(" <returns>An <see cref=""T:System.Object"" /> that represents the converted value.</returns>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""context"">An <see cref=""T:System.ComponentModel.ITypeDescriptorContext"" /> that provides a format context.</param>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""culture"">The <see cref=""T:System.Globalization.CultureInfo"" /> to use as the current culture.</param>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""value"">The <see cref=""T:System.Object"" /> to convert.</param>", True))
            .Comments.Add(New CodeCommentStatement(" <exception cref=""T:System.NotSupportedException"">The conversion cannot be performed.</exception>", True))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With
        Return convertFromMethod

    End Function

    ''' <summary>
    ''' Returns a <see cref="CodeMemberMethod"/> that overrides the <see cref="TypeConverter.CanConvertFrom"/> method.
    ''' </summary>
    ''' <param name="source">The entity source for the conversion.</param>
    ''' <returns>A <see cref="CodeMemberMethod"/> that overrides the <see cref="TypeConverter.CanConvertFrom"/> method.</returns>
    Function GetCanConvertFromMethod(ByVal source As EdmType) As CodeMemberMethod

        Dim canConvertFromMethod As New CodeMemberMethod
        With canConvertFromMethod
            .Name = "CanConvertFrom"
            .Attributes = MemberAttributes.Override Or MemberAttributes.Public
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(ITypeDescriptorContext), "context"))
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(Type), "sourceType"))
            .ReturnType = New CodeTypeReference(GetType(Boolean))

            '*
            '* Check if the parameter is valid
            '*
            .Statements.Add(New CodeMethodReturnStatement( _
                                    New CodeBinaryOperatorExpression( _
                                        New CodeArgumentReferenceExpression("sourceType"), _
                                        CodeBinaryOperatorType.IdentityEquality, _
                                        New CodeTypeOfExpression(GetNamespaceName(source.Name)))))

            '*
            '* Add the XML comment
            '*
            .Comments.Add(New CodeCommentStatement(" <summary>Returns whether this converter can convert an object of the given type to the type of this converter, using the specified context.</summary>", True))
            .Comments.Add(New CodeCommentStatement(" <returns>true if this converter can perform the conversion; otherwise, false.</returns>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""context"">An <see cref=""T:System.ComponentModel.ITypeDescriptorContext"" /> that provides a format context.</param>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""sourceType"">A <see cref=""T:System.Type"" /> that represents the type you want to convert from.</param>", True))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With
        Return canConvertFromMethod

    End Function

    ''' <summary>
    ''' Sets the code namespace.
    ''' </summary>
    ''' <param name="schema">The <see cref="XElement"/> that represents the CSDL schema read from the EDMX file.</param>
    Private Sub SetCodeNamespace(ByVal schema As XElement)
        Dim str As String = Me.FileNamespace
        If (str Is Nothing) Then
            str = String.Empty
        End If
        If Not (str = String.Empty) Then
            __CodeNamespace = str
        End If
        Dim proj = Me.GetProject()
        If (VSUtils.GetProjectKind(proj) = ProjectKind.VB) Then
            Dim projectPropertyByName As String = TryCast(VSUtils.GetProjectPropertyByName(proj, "RootNamespace"), String)
            If String.IsNullOrEmpty(projectPropertyByName) Then
                __CodeNamespace = schema.Attribute("Namespace")
            End If
            Return
        End If
        __CodeNamespace = schema.Attribute("Namespace")
    End Sub

    ''' <summary>
    ''' Returns a <see cref="CodeMemberMethod"/> that represents the partial methos AditionalConversion
    ''' </summary>
    ''' <param name="source">The source of the aditional parameter.</param>
    ''' <param name="destination">The destination of the aditional parameter.</param>
    ''' <returns>A <see cref="CodeMemberMethod"/> that represents the partial methos AditionalConversion</returns>
    Function GetAdditionalConversionMethod(ByVal source As EdmType, ByVal destination As CodeTypeReference) As CodeMemberMethod

        Dim cmm = New CodeMemberMethod
        With cmm
            .Name = "AdditionalConversion"
            .Parameters.Add(New CodeParameterDeclarationExpression(New CodeTypeReference(GetNamespaceName(source.Name)), "source"))
            .Parameters.Add(New CodeParameterDeclarationExpression(destination, "destination"))

            .Comments.Add(New CodeCommentStatement(" <summary>Provides aditional Conversions between the <paramref name=""source""/> and the <paramref name=""destination""/>.</summary>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""source"">The entity being converted.</param>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""destination"">The destination object being created.</param>", True))
            .Comments.Add(New CodeCommentStatement("##Partial##", False))
        End With

        Return cmm

    End Function

End Class
