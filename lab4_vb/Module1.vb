Imports System.Data.OleDb
Imports System.Data.SqlClient

Module Module1

    Sub Main()
        Dim menu As String
        menu = "1) зчитати з бд;" + Environment.NewLine +
                "2) додати у бд;" + Environment.NewLine +
                "3) видалити з бд;" + Environment.NewLine +
                "4) додати введення з консолi(до хеш таб);" + Environment.NewLine +
                "5) роздрукувати все;" + Environment.NewLine +
                "6) записати у бд(вiдсортоване);" + Environment.NewLine +
                "7) ;" + Environment.NewLine +
                "8) закiнчити роботу." + Environment.NewLine
        Dim choose As Integer

        Dim HashTabb As New Hashtable()


        Do
            Console.Write(menu)
            choose = Console.ReadLine()
            Select Case choose
                Case 1
                    HashTabb.Clear()
                    Dim f As String = "E:\mathfuck\вба\db_for_vb.accdb"
                    Dim t As String = "tab1"
                    Dim Conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & f)
                    Dim CommandSQL As New OleDbCommand("Select * From " & t, Conn)
                    Dim data As OleDbDataReader
                    Conn.Open()
                    data = CommandSQL.ExecuteReader
                    While data.Read()
                        HashTabb.Add(data.Item(3), New Car(data.Item(0), data.Item(1), data.Item(2), data.Item(3)))

                    End While
                    Conn.Close()
                Case 2
                    Dim new_elem As New Car()
                    new_elem.Input()
                    Dim f As String = "E:\mathfuck\вба\db_for_vb.accdb"
                    Dim t As String = "tab1"
                    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & f)
                    Dim cmd As New OleDbCommand
                    con.Open()
                    cmd.Connection = con
                    cmd.CommandText = "INSERT INTO tab1( [label], [pruduction_year], [second_name], [number]) VALUES('" + new_elem.Label + "','" + new_elem.Production_Year + "','" + new_elem.Second_Name + "','" + new_elem.Number + "')"
                    cmd.ExecuteNonQuery()
                    con.Close()
                Case 3
                    Console.WriteLine("Введіть номер машини за яким хочете видалити запис: ")
                    Dim s As String = Console.ReadLine()
                    Dim f As String = "E:\mathfuck\вба\db_for_vb.accdb"
                    Dim t As String = "tab1"
                    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & f)
                    Dim cmd As New OleDbCommand
                    con.Open()
                    cmd.Connection = con
                    cmd.CommandText = "DELETE FROM " + t + " WHERE number='" + s + "'"
                    cmd.ExecuteNonQuery()
                    con.Close()

                Case 4
                    Dim new_elem As New Car()
                    new_elem.Input()
                    HashTabb.Add(new_elem.Second_Name, new_elem)
                Case 5
                    Dim value_of_Hashtabb As ICollection = HashTabb.Values
                    For Each d As Car In value_of_Hashtabb
                        Console.WriteLine("{0}", d.toString())
                    Next
                Case 6
                    If HashTabb.Count > 1 Then
                        Dim keyColl As ICollection = HashTabb.Keys
                        Dim list As New ArrayList
                        For Each s As String In keyColl
                            list.Add(s)
                        Next s
                        list.Sort()

                        Dim f As String = "E:\mathfuck\вба\db_for_vb.accdb"
                        Dim t As String = "tab1"
                        Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & f)
                        Dim cmd As New OleDbCommand
                        con.Open()
                        cmd.Connection = con
                        cmd.CommandText = "DELETE * FROM " + t
                        cmd.ExecuteNonQuery()
                        For i = 0 To list.Count - 1
                            cmd.CommandText = "INSERT INTO tab1( [label], [pruduction_year], [second_name], [number]) VALUES('" + HashTabb(list(i)).Label + "','" + HashTabb(list(i)).Production_Year + "','" + HashTabb(list(i)).Second_Name + "','" + HashTabb(list(i)).Number + "')"
                            cmd.ExecuteNonQuery()
                        Next
                        con.Close()
                    Else
                        Console.WriteLine("хеш таблиця пуста")
                    End If
                Case 7
            End Select
        Loop While (choose <> 8)


    End Sub

End Module
