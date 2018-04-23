Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace gridCopyToClipboardExample
    Friend Class Employee
        Public Sub New(ByVal id As Integer, ByVal name As String, ByVal salary As Double, ByVal bonus As Double, ByVal department As String, ByVal hired As Date)
            Me.Id = id
            Me.Name = name
            Me.Salary = salary
            Me.Bonus = bonus
            Me.Department = department
            Me.Hired = hired
        End Sub

        Private privateId As Integer
        Public Property Id() As Integer
            Get
                Return privateId
            End Get
            Private Set(ByVal value As Integer)
                privateId = value
            End Set
        End Property
        Private privateName As String
        Public Property Name() As String
            Get
                Return privateName
            End Get
            Private Set(ByVal value As String)
                privateName = value
            End Set
        End Property
        Private privateSalary As Double
        Public Property Salary() As Double
            Get
                Return privateSalary
            End Get
            Private Set(ByVal value As Double)
                privateSalary = value
            End Set
        End Property
        Private privateBonus As Double
        Public Property Bonus() As Double
            Get
                Return privateBonus
            End Get
            Private Set(ByVal value As Double)
                privateBonus = value
            End Set
        End Property
        Private privateDepartment As String
        Public Property Department() As String
            Get
                Return privateDepartment
            End Get
            Private Set(ByVal value As String)
                privateDepartment = value
            End Set
        End Property
        Private privateHired As Date
        Public Property Hired() As Date
            Get
                Return privateHired
            End Get
            Private Set(ByVal value As Date)
                privateHired = value
            End Set
        End Property
    End Class
End Namespace
