'**************************************************
' FILE      : Utilities.vb
' AUTHOR    : paulo.santos
' CREATION  : 3/24/2009 8:24:46 AM
' COPYRIGHT : Copyright © 2009
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       TODO: Add file description
'
' Change log:
' 0.1   3/24/2009 8:24:46 AM
'       paulo.santos
'       Created.
'***************************************************

Imports System.CodeDom
Imports System.ComponentModel
Imports System.Data.Metadata.Edm
Imports System.IO
Imports System.Globalization
Imports System.Collections.ObjectModel
Imports System.Diagnostics.CodeAnalysis

Friend Module Utilities

    Friend Const __STR_Converter As String = "Converter"
    Friend Const __STR_Structures As String = "SimpleClasses"
    Friend Const __STR_ToStructure As String = "ToSimpleClass"

    ''' <summary>
    ''' Returns a <see cref="CodeConstructor"/> that represents the default constructor for the <see cref="TypeConverter"/> this class generate for each entity.
    ''' </summary>
    ''' <returns>A <see cref="CodeConstructor"/> that represents the default constructor for the <see cref="TypeConverter"/> this class generate for each entity.</returns>
    Function GetConverterDefaultCtor(ByVal className As String) As CodeConstructor
        Dim ctor1 As New CodeConstructor()
        With ctor1
            .Attributes = MemberAttributes.Public
            .Statements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__NumLevels"), New CodePrimitiveExpression(1)))

            .Comments.Add(New CodeCommentStatement(" <summary>Initializes an instance of the <see cref=""" & className & """ /> class.", True))
            .Comments.Add(New CodeCommentStatement(" This is the default constructor for this class.</summary>", True))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With
        Return ctor1
    End Function

    ''' <summary>
    ''' Returns a <see cref="CodeConstructor"/> that recieves as a parameter the number of levels to convert down from the entity.
    ''' </summary>
    ''' <returns>A <see cref="CodeConstructor"/> that recieves as a parameter the number of levels to convert down from the entity.</returns>
    Function GetCtor(ByVal className As String) As CodeConstructor
        Dim ctor2 As New CodeConstructor()
        With ctor2
            .Attributes = MemberAttributes.Public
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(Integer), "numLevels"))
            .Statements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__NumLevels"), New CodeArgumentReferenceExpression("numLevels")))

            .Comments.Add(New CodeCommentStatement(" <summary>Initializes an instance of the <see cref=""" & className & """ /> class.</summary>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""numLevels"">The number of levels down to convert.</param>", True))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With
        Return ctor2
    End Function

    ''' <summary>
    ''' Returns a <see cref="CodeMemberProperty"/> that contains the method to get and set the number of levels deep into the hierarchy that the converter will drill down.
    ''' </summary>
    ''' <returns>A <see cref="CodeMemberProperty"/> that contains the method to get and set the number of levels deep into the hierarchy that the converter will drill down.</returns>
    Function GetNumLevelsProperty() As CodeMemberProperty
        Dim numLevelsProperty As New CodeMemberProperty()
        With numLevelsProperty
            .Attributes = MemberAttributes.Public
            .Name = "NumLevels"
            .Type = New CodeTypeReference(GetType(Integer))

            .GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__NumLevels")))
            .SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__NumLevels"), New CodeArgumentReferenceExpression("value")))

            .Comments.Add(New CodeCommentStatement(" <summary>Gets or sets the number of levels that the converter will dril down when it reaches navigational properties.</summary>", True))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With
        Return numLevelsProperty
    End Function

    ''' <summary>
    ''' Returns a <see cref="CodeMemberMethod"/> that overrides the <see cref="TypeConverter.CanConvertTo"/> method.
    ''' </summary>
    ''' <param name="structName">The name of the structure that the entity generates.</param>
    ''' <returns>A <see cref="CodeMemberMethod"/> that overrides the <see cref="TypeConverter.CanConvertTo"/> method.</returns>
    Function GetCanConvertToMethod(ByVal structName As String) As CodeMemberMethod

        Dim canConvertToMethod As New CodeMemberMethod
        With canConvertToMethod
            .Name = "CanConvertTo"
            .Attributes = MemberAttributes.Override Or MemberAttributes.Public
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(ITypeDescriptorContext), "context"))
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(Type), "destinationType"))
            .ReturnType = New CodeTypeReference(GetType(Boolean))

            '*
            '* Check if the parameter is valid
            '*
            .Statements.Add(New CodeMethodReturnStatement( _
                                    New CodeBinaryOperatorExpression( _
                                        New CodeArgumentReferenceExpression("destinationType"), _
                                        CodeBinaryOperatorType.IdentityEquality, _
                                        New CodeTypeOfExpression(structName))))

            '*
            '* Add the XML comment
            '*
            .Comments.Add(New CodeCommentStatement(" <summary>Returns whether this converter can convert the object to the specified type, using the specified context.</summary>", True))
            .Comments.Add(New CodeCommentStatement(" <returns>true if this converter can perform the conversion; otherwise, false.</returns>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""context"">An <see cref=""T:System.ComponentModel.ITypeDescriptorContext"" /> that provides a format context.</param>", True))
            .Comments.Add(New CodeCommentStatement(" <param name=""destinationType"">A <see cref=""T:System.Type"" /> that represents the type you want to convert to.</param>", True))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With
        Return canConvertToMethod

    End Function

    ''' <summary>
    ''' Returns all the navigational properties for the specified entity.
    ''' </summary>
    ''' <param name="edmSource">The entity to retrieve the navigational properties.</param>
    ''' <returns>An <see cref="IEnumerable(Of Object)"/> that contains all the navigational properties for the specified entity.</returns>
    Function GetNavigationProperties(ByVal edmSource As MetadataItem) As IEnumerable(Of Object)
        Dim membersProperty = edmSource.MetadataProperties("Members")
        If (membersProperty Is Nothing) Then
            Return Nothing
        End If
        Return CType(membersProperty.Value, IEnumerable).OfType(Of Object)().Where(Function(x) x.GetType().FullName.Equals(GetType(NavigationProperty).FullName))
    End Function

    ''' <summary>
    ''' Returns all the primitive properties for the specified entity.
    ''' </summary>
    ''' <param name="edmSource">The entity to retrieve the primitive properties.</param>
    ''' <returns>An <see cref="IEnumerable(Of Object)"/> that contains all the primitive properties for the specified entity.</returns>
    Function GetPrimitiveProperties(ByVal edmSource As MetadataItem) As IEnumerable(Of Object)
        Dim membersProperty = edmSource.MetadataProperties("Members")
        If (membersProperty Is Nothing) Then
            Return Nothing
        End If
        Return CType(membersProperty.Value, IEnumerable).OfType(Of Object)().Where(Function(x) x.GetType().FullName.Equals(GetType(EdmProperty).FullName))
    End Function

    ''' <summary>
    ''' Adds a setter method of a collection for the specified navigationak property.
    ''' </summary>
    ''' <param name="struct">The structure where to add the setter method.</param>
    ''' <param name="navProperty">The navigational property to add.</param>
    ''' <param name="baseType">The base type of the navigational property.</param>
    Sub AddSetterMethodForCollection(ByVal struct As CodeTypeDeclaration, ByVal navProperty As NavigationProperty, ByVal baseType As CodeTypeReference)

        '*
        '* Create the Set Method
        '*
        Dim setMethod As New CodeMemberMethod()
        With setMethod
            .Name = "Set" & navProperty.Name
            .Attributes = MemberAttributes.FamilyAndAssembly
            .Parameters.Add(New CodeParameterDeclarationExpression(GetGenericTypeReference(GetType(IEnumerable(Of Object)), baseType), "value"))

            .Statements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__" & navProperty.Name), New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeArgumentReferenceExpression("value"), "ToArray"))))

            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With
        struct.Members.Add(setMethod)

    End Sub

    ''' <summary>
    ''' Gets a generic type reference based of the specified types.
    ''' </summary>
    ''' <param name="genericType">The generic base type.</param>
    ''' <param name="ofTypeReference">The referenced types for the generic type references.</param>
    ''' <returns>A <see cref="CodeTypeReference"/> that contains a generic type reference with the referenced types as its parameters.</returns>
    Function GetGenericTypeReference(ByVal genericType As Type, ByVal ParamArray ofTypeReference() As CodeTypeReference) As CodeTypeReference

        Dim typeRef As New CodeTypeReference(genericType)
        typeRef.TypeArguments.Clear()
        typeRef.TypeArguments.AddRange(ofTypeReference)
        Return typeRef

    End Function

    ''' <summary>
    ''' Add XML comment summary based on the enmtity information.
    ''' </summary>
    ''' <param name="edmSource">The entity information where to base the XML comments.</param>
    ''' <param name="codeMember">The code memeber where to add the XML comments.</param>
    Sub AddXmlSummary(ByVal edmSource As EdmMember, ByVal codeMember As CodeMemberProperty)

        Dim doc = edmSource.Documentation
        Dim xmlDoc As New Xml.XmlDocument

        '*
        '* Create the <summary> node
        '*
        Dim xmlEl = xmlDoc.CreateElement("summary")
        If (doc IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(doc.Summary)) Then
            xmlEl.InnerText = doc.Summary
        Else
            xmlEl.InnerText = "There are no comments for " & edmSource.Name & " in the schema"
        End If
        codeMember.Comments.Add(New CodeCommentStatement(xmlEl.OuterXml, True))

        '*
        '* Create the <remarks> node
        '*
        If (doc IsNot Nothing) AndAlso (Not String.IsNullOrEmpty(doc.LongDescription)) Then
            xmlEl = xmlDoc.CreateElement("remarks")
            xmlEl.InnerText = doc.LongDescription
            codeMember.Comments.Add(New CodeCommentStatement(xmlEl.OuterXml, True))
        End If

    End Sub

    ''' <summary>
    ''' Adds a getter and setter method for the specified peoprty.
    ''' </summary>
    ''' <param name="primProperty">The primitive property in question.</param>
    ''' <param name="propertyMember">The property memeber to add the getter and setter methods.</param>
    Sub AddGetterAndSetterToProperty(ByVal primProperty As EdmMember, ByVal propertyMember As CodeMemberProperty)

        With propertyMember
            '*
            '* Create the return statement
            '*
            .GetStatements.Add(New CodeMethodReturnStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__" & primProperty.Name)))

            '*
            '* Create the assign statement
            '*
            .SetStatements.Add(New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression(), "__" & primProperty.Name), New CodeArgumentReferenceExpression("value")))

            '*
            '* Identifies that it's an automatically generated code
            '*
            .CustomAttributes.Add(New CodeAttributeDeclaration(New CodeTypeReference(GetType(DebuggerNonUserCodeAttribute))))
        End With

    End Sub

    ''' <summary>
    ''' Extracts CSDL content from the EDMX file content and returns it as an XElement
    ''' </summary>
    ''' <param name="inputFileContent">EDMX file conmtent</param>
    ''' <returns>CSDL content XElement</returns>
    Function ExtractCsdlContent(ByVal inputFileContent As String) As XElement

        Dim csdlContent As XElement = Nothing
        Dim edmxns As XNamespace = "http://schemas.microsoft.com/ado/2007/06/edmx"
        Dim edmns As XNamespace = "http://schemas.microsoft.com/ado/2006/04/edm"

        '*
        '* Reads the XML Documento
        '*
        Dim edmxDoc = XDocument.Load(New StringReader(inputFileContent))
        If (edmxDoc IsNot Nothing) Then
            '*
            '* Gets the <edmx:Edmx> node
            '*
            Dim edmxNode = edmxDoc.Element(edmxns + "Edmx")
            If (edmxNode IsNot Nothing) Then
                '*
                '* Gets the <edmx:Runtime> node
                '*
                Dim runtimeNode = edmxNode.Element(edmxns + "Runtime")
                If (runtimeNode IsNot Nothing) Then
                    '*
                    '* Gets the <edmx:ConceptualModels> node
                    '*
                    Dim conceptualModelsNode = runtimeNode.Element(edmxns + "ConceptualModels")
                    If (conceptualModelsNode IsNot Nothing) Then
                        '*
                        '* Gets the <Schema> node
                        '*
                        csdlContent = conceptualModelsNode.Element(edmns + "Schema")
                    End If
                End If
            End If
        End If

        Return csdlContent

    End Function

End Module
