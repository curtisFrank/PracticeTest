'**************************************************************************
' Name: Problem.vb
' Programmer: Curtis N Frank
' Date: 4/2/2016
' Assignment: Advanced VB.NET ITSE 2349 Individual Project
' Purpose: Specification file for the Problem class.
'**************************************************************************

Public Class Problem

    ' private member variables
    Private _question As String
    Private _answer As String

    ' public property
    Public Property Question As String
        Get
            Return _question
        End Get
        Set(value As String)
            _question = value
        End Set
    End Property

    ' public property
    Public Property Answer As String
        Get
            Return _answer
        End Get
        Set(value As String)
            _answer = value
        End Set
    End Property

    ' default constructor
    Public Sub New()
        _question = ""
        _answer = ""
    End Sub

    ' overloaded constructor
    Public Sub New(ByVal q As String, ByVal a As String)
        _question = q
        _answer = a
    End Sub

End Class
