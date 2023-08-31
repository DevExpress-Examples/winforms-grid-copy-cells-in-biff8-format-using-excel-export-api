Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace gridCopyToClipboardExample

    Friend Class Employee

        Private _Id As Integer, _Name As String, _Salary As Double, _Bonus As Double, _Department As String, _Hired As DateTime

        Public Sub New(ByVal id As Integer, ByVal name As String, ByVal salary As Double, ByVal bonus As Double, ByVal department As String, ByVal hired As System.DateTime)
            Me.Id = id
            Me.Name = name
            Me.Salary = salary
            Me.Bonus = bonus
            Me.Department = department
            Me.Hired = hired
        End Sub

        Public Property Id As Integer
            Get
                Return _Id
            End Get

            Private Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property Name As String
            Get
                Return _Name
            End Get

            Private Set(ByVal value As String)
                _Name = value
            End Set
        End Property

        Public Property Salary As Double
            Get
                Return _Salary
            End Get

            Private Set(ByVal value As Double)
                _Salary = value
            End Set
        End Property

        Public Property Bonus As Double
            Get
                Return _Bonus
            End Get

            Private Set(ByVal value As Double)
                _Bonus = value
            End Set
        End Property

        Public Property Department As String
            Get
                Return _Department
            End Get

            Private Set(ByVal value As String)
                _Department = value
            End Set
        End Property

        Public Property Hired As DateTime
            Get
                Return _Hired
            End Get

            Private Set(ByVal value As DateTime)
                _Hired = value
            End Set
        End Property
    End Class
End Namespace
