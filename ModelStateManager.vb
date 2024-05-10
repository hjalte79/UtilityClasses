Public Class ModelStateManager
	Implements IDisposable
	
	' Example use:
	' http://www.hjalte.nl/tutorials/72-wichmodelstate
    
    Private _doc As Document
    Private _documentType As DocumentTypeEnum
    Private _startActiveModelState As ModelState
    Private _startMemberEditScope As MemberEditScopeEnum

    Public Sub New(ByVal doc As Document)
        If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            Throw New ArgumentException("This is not a part or assembly document")
        End If

        If doc Is Nothing Then
            Throw New ArgumentException("Document may not be NULL")
        End If

        _doc = doc
        _documentType = _doc.DocumentType
        _doc = FactoryDocument


        If ModelStates.Count <> 0 Then
            _startActiveModelState = ModelStates.ActiveModelState
        End If

        _startMemberEditScope = ModelStates.MemberEditScope
    End Sub

    Public Sub New(ByVal doc As Document, ByVal editScope As MemberEditScopeEnum)
        Me.New(doc)
        Me.EditScope = editScope
    End Sub

    Public ReadOnly Property FactoryDocument As Document
        Get
            If (_doc.ComponentDefinition.IsModelStateMember) Then
                Return _doc.ComponentDefinition.FactoryDocument
            End If

            Return _doc
        End Get
    End Property

    Public ReadOnly Property ModelStates As ModelStates
        Get
            Return _doc.ComponentDefinition.ModelStates
        End Get
    End Property

    Public Property EditScope As MemberEditScopeEnum
        Get
            Return ModelStates.MemberEditScope
        End Get
        Set(ByVal value As MemberEditScopeEnum)
            ModelStates.MemberEditScope = value
        End Set
    End Property

    Public Sub Dispose() Implements IDisposable.Dispose
        If (_startActiveModelState IsNot Nothing) Then
            _startActiveModelState.Activate()
        End If
        ModelStates.MemberEditScope = _startMemberEditScope
    End Sub
	
	' Copyright 2024
    ' 
    ' This code was written by Jelte de Jong, and published on www.hjalte.nl/https://github.com/hjalte79
    '
    ' Permission Is hereby granted, free of charge, to any person obtaining a copy of this 
    ' software And associated documentation files (the "Software"), to deal in the Software 
    ' without restriction, including without limitation the rights to use, copy, modify, merge, 
    ' publish, distribute, sublicense, And/Or sell copies of the Software, And to permit persons 
    ' to whom the Software Is furnished to do so, subject to the following conditions:
    '
    ' The above copyright notice And this permission notice shall be included In all copies Or
    ' substantial portions Of the Software.
    ' 
    ' THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY Of ANY KIND, EXPRESS Or IMPLIED, 
    ' INCLUDING BUT Not LIMITED To THE WARRANTIES Of MERCHANTABILITY, FITNESS For A PARTICULAR 
    ' PURPOSE And NONINFRINGEMENT. In NO Event SHALL THE AUTHORS Or COPYRIGHT HOLDERS BE LIABLE 
    ' For ANY CLAIM, DAMAGES Or OTHER LIABILITY, WHETHER In AN ACTION Of CONTRACT, TORT Or 
    ' OTHERWISE, ARISING FROM, OUT Of Or In CONNECTION With THE SOFTWARE Or THE USE Or OTHER 
    ' DEALINGS In THE SOFTWARE.
End Class