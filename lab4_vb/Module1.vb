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
                "7) відбір власників ZAZ;" + Environment.NewLine +
                "8) закiнчити роботу." + Environment.NewLine
        Dim choose As Integer

        Dim HashTabb As New Hashtable()
        Dim f As String = "E:\mathfuck\вба\db_for_vb.accdb"
        Dim t As String = "tab1"
        Dim con
        Try
            con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & f)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try



        Do
            Console.Write(menu)
            choose = Console.ReadLine()
            Select Case choose
                Case 1
                    HashTabb.Clear()
                    Try
                        Dim CommandSQL As New OleDbCommand("Select * From " & t, con)
                        Dim data As OleDbDataReader
                        con.Open()
                        data = CommandSQL.ExecuteReader
                        While data.Read()
                            HashTabb.Add(data.Item(3), New Car(data.Item(0), data.Item(1), data.Item(2), data.Item(3)))

                        End While
                        con.Close()
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try

                Case 2
                    Dim new_elem As New Car()
                    new_elem.Input()
                    Try
                        Dim cmd As New OleDbCommand
                        con.Open()
                        cmd.Connection = con
                        cmd.CommandText = "INSERT INTO tab1( [label], [pruduction_year], [second_name], [number]) VALUES('" + new_elem.Label + "','" + new_elem.Production_Year + "','" + new_elem.Second_Name + "','" + new_elem.Number + "')"
                        cmd.ExecuteNonQuery()
                        con.Close()
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try

                Case 3
                    Console.WriteLine("Введіть номер машини за яким хочете видалити запис: ")
                    Dim s As String = Console.ReadLine()
                    Try
                        Dim cmd As New OleDbCommand
                        con.Open()
                        cmd.Connection = con
                        cmd.CommandText = "DELETE FROM " + t + " WHERE number='" + s + "'"
                        cmd.ExecuteNonQuery()
                        con.Close()
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try


                Case 4
                    Dim myEx As New ArgumentException("невірно введені дані")
                    Dim new_elem As New Car()
                    new_elem.Input()
                    Try
                        If (new_elem.Label = "VAZ") Then
                            Throw myEx
                        Else
                            HashTabb.Add(new_elem.Second_Name, new_elem)
                        End If
                    Catch ex As Exception
                        Console.WriteLine(ex.Message)
                    End Try




                Case 5
                    Console.WriteLine("a) вивести таблицею" + Environment.NewLine + "b) вивести не таблицею")
                    Dim ch As String = Console.ReadLine()
                    If ch = "b" Then

                        For Each d As Car In HashTabb.Values
                            Console.WriteLine("{0}", d.toString())
                        Next
                    ElseIf ch = "a" Then
                        Dim s As String = "|______|_________|_____|_____________|"
                        Console.WriteLine("{0}{1}{0}", s, vbCrLf + "|Number|__Label__|Year_|_Second Name_|" + Chr(10) + Chr(13))

                        For Each elem As Car In HashTabb.Values

                            Console.WriteLine("|{0,-6}|{1,9}|{2,5}|{3,13}|", elem.Number, elem.Label, elem.Production_Year, elem.Second_Name)
                            Console.WriteLine(s)

                        Next


                    End If

                Case 6
                    If HashTabb.Count > 0 Then
                        Dim keyColl As ICollection = HashTabb.Keys
                        Dim list As New ArrayList
                        For Each s As String In keyColl
                            list.Add(s)
                        Next s
                        list.Sort()
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
                    If HashTabb.Count > 0 Then
                        Dim new_list As New ArrayList()
                        For Each elem As Car In HashTabb.Values
                            If (elem.Label = "ZAZ") Then
                                new_list.Add(elem)
                            End If
                        Next

                        For Each elem As Car In new_list
                            Console.WriteLine(elem.toString())
                        Next
                    Else
                        Console.WriteLine("хеш таблиця пуста")
                    End If


            End Select
        Loop While (choose <> 8)


    End Sub

End Module
