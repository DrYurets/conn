
' ЭКСПОРТ на 9 СКЛАДОВ. Последнее обновление от 08/01/18 с созданием таблиц для extmedia. Yurets

Imports System
Imports MySql.Data
Imports MySql.Data.MySqlClient
Module connector

    Dim logf = My.Application.Info.DirectoryPath & "\export.log"
    Private Function preparer(strin As String)
        strin = Replace(strin, "'", "")
        strin = Replace(strin, "/", "")
        strin = Replace(strin, "#", "")
        strin = Replace(strin, "$", "")
        strin = Replace(strin, "", "")
        Return strin
    End Function

    Sub Main()
        Try
            Dim connectionstring, constr, LogFile, nadpis(40), topPath(40), frompath(40), top(40), setting(40) As String
            Dim cat As ADOX.Catalog = New ADOX.Catalog
            Dim FileNameSystemDataBase As String

            Dim SettingsFile = My.Application.Info.DirectoryPath & "\settings.txt"
            Dim fileExists, topsExists As Boolean
            Dim totalTops As Integer
            Dim TopsFile = My.Application.Info.DirectoryPath & "\tops.txt"
            Dim TopFrom = My.Application.Info.DirectoryPath & "\from.txt"

            fileExists = My.Computer.FileSystem.FileExists(SettingsFile)
            If fileExists = True Then
                Dim sett As IO.StreamReader
                Dim i As Integer
                sett = IO.File.OpenText(SettingsFile)
                For i = 1 To 10
                    nadpis(i) = sett.ReadLine
                    If Not IsNothing(nadpis(i)) Then
                        setting(i) = nadpis(i)
                    End If
                Next
                sett.Close()
                sett = Nothing
            End If
            LogFile = setting(1)                            'c:\export\export\export.log
            FileNameSystemDataBase = setting(3)             'c:\export\export\svgrp.mdw       
            topsExists = My.Computer.FileSystem.FileExists(TopsFile)

            connectionstring = setting(4)
            Dim conn As New MySqlConnection
            conn.ConnectionString = connectionstring
            conn.Open()

            System.IO.File.AppendAllText(logf, "" & Now + Environment.NewLine + "Соединение установлено" + Environment.NewLine)

            Dim cmd As New MySqlCommand
            cmd.Connection = conn

            Dim SqlQry = "DROP TABLE IF EXISTS `conn_tovar`;" & _
                            "CREATE TABLE IF NOT EXISTS `conn_tovar` (" & _
                            "`id_tovar` bigint(20) NOT NULL DEFAULT '0'," & _
                            "  `code` text NOT NULL," & _
                            "  `demo` text NOT NULL," & _
                            "  `cross_num` mediumtext NOT NULL," & _
                            "  `proizv` text NOT NULL," & _
                            "  `qnt` bigint(20) DEFAULT '0'," & _
                            "  `hidecode` varchar(10) DEFAULT '0'," & _
                            "  `price` varchar(10) DEFAULT NULL," & _
                            "  `roznica` varchar(10) DEFAULT NULL," & _
                            "  `comment` text NOT NULL," & _
                            "  `zakup` varchar(10) DEFAULT NULL," & _
                            "  `red` mediumtext NOT NULL," & _
                            "  `sklad` mediumint(10) NOT NULL," & _
                            "  UNIQUE KEY `id` (`id_tovar`)," & _
                            "  UNIQUE KEY `id_tovar` (`id_tovar`)," & _
                            "  UNIQUE KEY `id_tovar_2` (`id_tovar`)," & _
                            "  FULLTEXT KEY `comm` (`comment`)" & _
                            ") ENGINE=MyISAM DEFAULT CHARSET=cp1251;" & _
                            "DROP TABLE IF EXISTS `conn_analogi`;" & _
                            "CREATE TABLE IF NOT EXISTS `conn_analogi` (" & _
                            "`id_tovar` bigint(20) NOT NULL DEFAULT '0'," & _
                            "`id_analog` bigint(20) NOT NULL DEFAULT '0'," & _
                            "`sklad` mediumint(10) NOT NULL," & _
                            "KEY `id` (`id_tovar`)" & _
                            ") ENGINE=MyISAM DEFAULT CHARSET=cp1251;"

            cmd.CommandText = SqlQry
            cmd.Prepare()
            cmd.ExecuteNonQuery()

            System.IO.File.AppendAllText(logf, "" & Now + Environment.NewLine + "Таблицы успешно созданы" + Environment.NewLine)


            If topsExists = True Then
                Dim topFile As IO.StreamReader
                Dim DestTop As IO.StreamReader
                Dim i As Integer
                totalTops = 0
                topFile = IO.File.OpenText(TopsFile)
                DestTop = IO.File.OpenText(TopFrom)

                For i = 1 To 10
                    topPath(i) = topFile.ReadLine
                    frompath(i) = DestTop.ReadLine
                    If Len(topPath(i)) > 0 Then
                        If System.IO.File.Exists(topPath(i)) Then
                            System.IO.File.Delete(topPath(i))
                        End If
                        System.IO.File.Copy(frompath(i), topPath(i))
                        If Not IsNothing(topPath(i)) Then
                            totalTops = totalTops + 1
                            top(totalTops) = topPath(i)
                        End If
                    End If
                Next
                topFile.Close()
                topFile = Nothing
                topsExists = Nothing


            End If
            System.IO.File.AppendAllText(logf, "Складов " + totalTops.ToString + Environment.NewLine)
            Console.WriteLine("Складов: " & totalTops.ToString)

            Dim itera As Integer
            Dim cn, cn1 As ADODB.Connection, rst, ost As ADODB.Recordset

            Dim FileNameSource, InsertValues, ValuesArr, comma, comma2 As String
            Dim TotalTovarov, id_tovara, id_analoga As Integer

            
            For itera = 1 To totalTops 'начинаем цикл экспорта
                FileNameSource = top(itera)
                cn1 = New ADODB.Connection
                constr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & FileNameSource & "; Jet OLEDB:System database=" & FileNameSystemDataBase
                cn1.Open(constr, "Yurets", "13421342")
                ost = New ADODB.Recordset
                ost.Open("Экспорт_остатков_аналоги", cn1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                ost.MoveFirst()
                ValuesArr = ""
                TotalTovarov = 0
                comma = ""
                InsertValues = "INSERT INTO `conn_analogi` (id_tovar, id_analog, sklad) VALUES "
                While Not ost.EOF

                    If TotalTovarov = 800 Then
                        comma2 = ";"
                    Else
                        comma2 = ""
                    End If
                    If itera = 3 Then
                        id_tovara = ost.Fields(0).Value - 29894
                        id_analoga = ost.Fields(1).Value - 29894
                    Else
                        id_tovara = ost.Fields(0).Value
                        id_analoga = ost.Fields(1).Value
                    End If

                    ValuesArr = ValuesArr & comma & " (" & ost.Fields(0).Value & "," & ost.Fields(1).Value & "," & itera & ")" & comma2
                    comma = ","
                    If TotalTovarov = 800 Then
                        cmd.CommandText = InsertValues & ValuesArr
                        cmd.Prepare()
                        cmd.ExecuteNonQuery()
                        Console.Write(".")
                        ValuesArr = ""
                        TotalTovarov = 0
                        comma = ""
                        comma2 = ""
                    End If
                    ost.MoveNext()
                    TotalTovarov = TotalTovarov + 1
                End While
                cmd.CommandText = InsertValues & ValuesArr & ";"
                cmd.Prepare()
                cmd.ExecuteNonQuery()

                ost = Nothing

                InsertValues = Nothing
                ValuesArr = ""
                comma = Nothing
                System.IO.File.AppendAllText(logf, "a" + itera.ToString + " / ")

                ost = New ADODB.Recordset
                ost.Open("Export_tovari1", cn1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                ost.MoveFirst()
                ValuesArr = ""
                TotalTovarov = 0
                comma = ""
                InsertValues = "INSERT INTO `conn_tovar` (`id_tovar`, `code`, `demo`, `cross_num`, `proizv`, `qnt`, `hidecode`, `price`, `roznica`, `comment`, `zakup`, `red`, `sklad`) VALUES "

                Dim code, demo, cross_num, proizv, comment, red, price, roznica As String
                Dim id_tovar, hidecode, sklad, qnt As Integer

                While Not ost.EOF
                    If TotalTovarov = 500 Then
                        comma2 = ";"
                    Else
                        comma2 = ""
                    End If

                    If itera = 3 Then
                        id_tovar = ost.Fields(0).Value - 29894 'id_tovar
                    Else
                        id_tovar = ost.Fields(0).Value 'id_tovar
                    End If


                    code = LTrim(ost.Fields(1).Value) 'code

                    code = preparer(code)

                    demo = ""
                    If ost.Fields(2).Value.Equals(System.DBNull.Value) Then 'comment
                        comment = ""
                    Else
                        comment = Replace(Left(ost.Fields(2).Value, 255), "'", "")
                    End If
                    If ost.Fields(3).Value.Equals(System.DBNull.Value) Then 'cross
                        cross_num = ""
                    Else
                        cross_num = Replace(Left(ost.Fields(3).Value, 255), "'", "")
                    End If
                    If ost.Fields(4).Value.Equals(System.DBNull.Value) Then 'price_eur
                        price = 0
                    Else
                        price = ost.Fields(4).Value
                    End If

                    If ost.Fields(5).Value.Equals(System.DBNull.Value) Then 'qnt
                        qnt = 0
                    Else
                        qnt = ost.Fields(5).Value
                    End If

                    If ost.Fields(6).Value.Equals(System.DBNull.Value) Then 'rozn
                        roznica = 0
                    Else
                        roznica = ost.Fields(6).Value
                    End If
                    If ost.Fields(7).Value.Equals(System.DBNull.Value) Then 'proizv
                        proizv = "LIC"
                    Else
                        proizv = ost.Fields(7).Value
                    End If
                    If ost.Fields(8).Value.Equals(System.DBNull.Value) Then 'hidecode
                        hidecode = ""
                    Else
                        hidecode = ost.Fields(8).Value
                    End If

                    If ost.Fields(10).Value.Equals(System.DBNull.Value) Then 'red
                        red = ""
                    Else
                        red = Left(ost.Fields(10).Value, 255)
                    End If
                    sklad = itera

                    ValuesArr = ValuesArr & comma & " ('" & id_tovar & "','" & code & "','" & demo & "','" & cross_num & "','" & proizv & "','" & qnt & "','" & hidecode & "','" & price & "','" & roznica & "','" & comment & "','0','" & red & "','" & sklad & "')" & comma2
                    comma = ","
                    If TotalTovarov = 500 Then
                        cmd.CommandText = InsertValues & ValuesArr
                        cmd.Prepare()
                        cmd.ExecuteNonQuery()
                        Console.Write(".")
                        ValuesArr = ""
                        TotalTovarov = 0
                        comma = ""
                        comma2 = ""
                    End If
                    ost.MoveNext()
                    TotalTovarov = TotalTovarov + 1
                End While
                cmd.CommandText = InsertValues & ValuesArr & ";"
                cmd.Prepare()
                cmd.ExecuteNonQuery()
                System.IO.File.AppendAllText(logf, "t" + itera.ToString + Environment.NewLine)
                ost.Close()
                cn = Nothing
                cn1 = Nothing
                rst = Nothing
                ost = Nothing
            Next

            cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = `code`;" & _
                                "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '-','');" & _
                                "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '.','');" & _
                                "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '/','');" & _
                                "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, ' ','');" & _
                                "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '+','');" & _
                                "DROP TABLE `tovar`;" & _
                                "RENAME TABLE  `conn_tovar` TO `tovar`;" & _
                                "DROP TABLE `analogi`;" &
                                "RENAME TABLE  `conn_analogi` TO `analogi`;"
            
            cmd.Prepare()
            cmd.ExecuteNonQuery()
            Console.WriteLine("Импорт завершен")
            System.IO.File.AppendAllText(logf, Now + Environment.NewLine + "Export Success" + Environment.NewLine)
            conn.Close()

        Catch MySqlException As MySqlException
            System.IO.File.AppendAllText(logf, Now & MySqlException.ToString + Environment.NewLine)
        Catch MyException As Exception
            System.IO.File.AppendAllText(logf, Now & MyException.ToString + Environment.NewLine)
        End Try
    End Sub
End Module
