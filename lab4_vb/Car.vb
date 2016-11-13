Public Class Car
    Public Number As String
    Public Label As String
    Public Production_Year As String
    Public Second_Name As String

    Sub New(ByVal number As String, label As String, production_year As String, second_name As String)
        Me.Number = number
        Me.Label = label
        Me.Production_Year = production_year
        Me.Second_Name = second_name
    End Sub

    Sub New()
        Me.Number = ""
        Me.Label = ""
        Me.Production_Year = ""
        Me.Second_Name = ""
    End Sub

    Public Sub Input()
        Console.WriteLine("Введіть номер авто: ")
        Me.Number = Console.ReadLine()

        Console.WriteLine("Введіть марку авто: ")
        Me.Label = Console.ReadLine()

        Console.WriteLine("Введіть рік випуску авто: ")
        Me.Production_Year = Console.ReadLine()

        Console.WriteLine("Введіть прізвище власника авто: ")
        Me.Second_Name = Console.ReadLine()

    End Sub

    Public Overridable Function toString() As String

        Return String.Format("Номер: {0}, Марка: {1}, Рік випуску: {2}, Прізвище: {3}", Number, Label, Production_Year, Second_Name)

    End Function
End Class
