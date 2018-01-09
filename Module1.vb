


' РАБОЧИЙ ЭКСПОРТ на 9 СКЛАДОВ. Последнее обновление от 14/12/17 с созданием таблиц для AZURE. Yurets


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
            'Dim conn As New MySqlConnection
            'conn.ConnectionString = connectionstring
            'conn.Open()

            'System.IO.File.AppendAllText(logf, "" & Now + Environment.NewLine + "Соединение установлено" + Environment.NewLine)

            'Dim cmd As New MySqlCommand
            'cmd.Connection = conn

            'Dim SqlQry = "DROP TABLE IF EXISTS `conn_tovar`;" & _
            '                "CREATE TABLE IF NOT EXISTS `conn_tovar` (" & _
            '                "`id_tovar` bigint(20) NOT NULL DEFAULT '0'," & _
            '                "  `code` text NOT NULL," & _
            '                "  `demo` text NOT NULL," & _
            '                "  `cross_num` mediumtext NOT NULL," & _
            '                "  `proizv` text NOT NULL," & _
            '                "  `qnt` bigint(20) DEFAULT '0'," & _
            '                "  `hidecode` varchar(10) DEFAULT '0'," & _
            '                "  `price` varchar(10) DEFAULT NULL," & _
            '                "  `roznica` varchar(10) DEFAULT NULL," & _
            '                "  `comment` text NOT NULL," & _
            '                "  `zakup` varchar(10) DEFAULT NULL," & _
            '                "  `red` mediumtext NOT NULL," & _
            '                "  `sklad` mediumint(10) NOT NULL," & _
            '                "  UNIQUE KEY `id` (`id_tovar`)," & _
            '                "  UNIQUE KEY `id_tovar` (`id_tovar`)," & _
            '                "  UNIQUE KEY `id_tovar_2` (`id_tovar`)," & _
            '                "  FULLTEXT KEY `comm` (`comment`)" & _
            '                ") ENGINE=InnoDB COLLATE='utf8_general_ci';" & _
            '                "DROP TABLE IF EXISTS `conn_analogi`;" & _
            '                "CREATE TABLE IF NOT EXISTS `conn_analogi` (" & _
            '                "`id_tovar` bigint(20) NOT NULL DEFAULT '0'," & _
            '                "`id_analog` bigint(20) NOT NULL DEFAULT '0'," & _
            '                "`sklad` mediumint(10) NOT NULL," & _
            '                "KEY `id` (`id_tovar`)" & _
            '                ") ENGINE=InnoDB COLLATE='utf8_general_ci';" & _
            '                "DROP TABLE IF EXISTS `conn_analogi_old`;" & _
            '                "CREATE TABLE IF NOT EXISTS `conn_analogi_old` (" & _
            '                "`id_tovar` bigint(20) NOT NULL DEFAULT '0'," & _
            '                "`id_analog` bigint(20) NOT NULL DEFAULT '0'," & _
            '                "`sklad` mediumint(10) NOT NULL," & _
            '                "KEY `id` (`id_tovar`)" & _
            '                ") ENGINE=InnoDB COLLATE='utf8_general_ci';"

            'cmd.CommandText = SqlQry
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()



            'System.IO.File.AppendAllText(logf, "" & Now + Environment.NewLine + "Таблицы успешно созданы" + Environment.NewLine)


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
                        If System.IO.File.Exists(topPath(i) & "top_prev.mdb") Then
                            System.IO.File.Delete(topPath(i) & "top_prev.mdb") ' удаляем старый топ
                            System.IO.File.Move(topPath(i) & "top.mdb", topPath(i) & "\top_prev.mdb") 'переименовываем топ
                            System.IO.File.Copy(frompath(i), topPath(i) & "top.mdb") ' кладем новый топ
                        Else
                            System.IO.File.Copy(frompath(i), topPath(i) & "top.mdb")
                        End If
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
            Dim InsertValues, ValuesArr, comma, comma2 As String
            Dim TotalTovarov, id_tovara, id_analoga As Integer

            Dim latest_mdb As String
            latest_mdb = My.Application.Info.DirectoryPath & "\merged.accdb"

            If System.IO.File.Exists(latest_mdb) Then
                System.IO.File.Delete(My.Application.Info.DirectoryPath & "\merged.accdb")
                ' System.IO.File.Move(latest_mdb, My.Application.Info.DirectoryPath & "\merged_prev.mdb")
            End If
            If System.IO.File.Exists(My.Application.Info.DirectoryPath & "\merged.ldb") Then
                System.IO.File.Delete(My.Application.Info.DirectoryPath & "\merged.ldb")
            End If

            Dim merged As Catalog = New Catalog()
            merged.Create("Provider=Microsoft.ACE.OLEDB.12.0;" & _
                        "Data Source=" & My.Application.Info.DirectoryPath & "\merged.accdb;" & _
                        "Jet OLEDB:Engine Type=5")

            Dim con_merg As New ADODB.Connection
            Dim merged_db_path As String
            merged_db_path = My.Application.Info.DirectoryPath & "\merged.accdb"

            con_merg.Open("Provider=Microsoft.ACE.OLEDB.12.0; data source=" & merged_db_path & ";")
            Console.WriteLine("Database Created Successfully")

            con_merg.Execute("CREATE TABLE old_a (ID_Tovar INTEGER, ID_Analog INTEGER, sklad INTEGER);")
            con_merg.Execute("CREATE TABLE new_a (ID_Tovar INTEGER, ID_Analog INTEGER, sklad INTEGER);")
            Console.WriteLine("Tables Created Successfully")

            Dim conn_a As ADODB.Connection
            Dim new_a, old_a, differ As ADODB.Recordset
            Dim Old_FileNameSource, FileNameSource As String

            Dim fieldsArray(2) As Object
            fieldsArray(0) = "id_tovar"
            fieldsArray(1) = "id_analog"
            fieldsArray(2) = "sklad"

            Dim values(1) As Object

            Dim analogi_query, added_analogs As String
            Dim corrector_id As Integer

            added_analogs = "SELECT nw.id_tovar AS nwt, nw.id_analog AS nwa, ol.id_tovar AS olt, ol.id_analog AS ola, nw.sklad AS nws, ol.sklad AS ols " & _
                                "FROM new_a AS nw " & _
                                "LEFT JOIN old_a AS ol " & _
                                "ON nw.id_tovar = ol.id_tovar AND nw.id_analog = ol.id_analog " & _
                                "WHERE ol.id_tovar Is NULL " & _
                                "UNION ALL " & _
                                "SELECT nw.id_tovar AS nwt, nw.id_analog AS nwa, ol.id_tovar AS olt, ol.id_analog AS ola, nw.sklad AS nws, ol.sklad AS ols " & _
                                "FROM old_a AS ol " & _
                                "LEFT JOIN new_a AS nw " & _
                                "ON nw.id_tovar = ol.id_tovar AND nw.id_analog = ol.id_analog " & _
                                "WHERE nw.id_tovar IS NULL;"

            corrector_id = 0

            constr = "Provider=Microsoft.Jet.OLEDB.4.0;"
            differ = New ADODB.Recordset

            For itera = 1 To totalTops 'ПОЕХАЛИ!
                If itera = 2 Then
                    corrector_id = 170000
                End If
                analogi_query = "SELECT Tovari.itemid +" & corrector_id & " AS ID_Tovar, Tovari_1.itemid +" & corrector_id & " AS ID_Analog FROM (Analogi INNER JOIN Tovari ON Analogi.Analog = Tovari.Tovar) INNER JOIN Tovari AS Tovari_1 ON Analogi.Tovar = Tovari_1.Tovar ORDER BY Tovari.itemid +" & corrector_id & ", Tovari_1.itemid +" & corrector_id & ""

                'получили старую таблицу аналогов в oost
                Old_FileNameSource = top(itera) & "top_prev.mdb"
                conn_a = New ADODB.Connection
                constr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Old_FileNameSource & "; Jet OLEDB:System database=" & FileNameSystemDataBase
                conn_a.Open(constr, "Yurets", "13421342")

                'conn_a.Execute("SELECT * INTO new_a FROM (" & analogi_query & ");")

                'Dim query As String
                'query = "INSERT INTO new_azz SELECT * FROM new_a;"

                '                System.IO.File.AppendAllText(logf, "" & Now + Environment.NewLine + query + Environment.NewLine)


                'conn_a.Execute(query)

                old_a = New ADODB.Recordset
                old_a.Open(analogi_query, conn_a, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)


                'con_merg.Execute("INSERT INTO new_a SELECT * FROM ();")
                'old_a.Open("Экспорт_остатков_аналоги", conn_a, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

                'получили новую таблицу аналогов
                FileNameSource = top(itera) & "top.mdb"
                conn_a = New ADODB.Connection
                constr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & FileNameSource & "; Jet OLEDB:System database=" & FileNameSystemDataBase
                conn_a.Open(constr, "Yurets", "13421342")
                new_a = New ADODB.Recordset

                new_a.Open(analogi_query, conn_a, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)


                'new_a.Open("Экспорт_остатков_аналоги", conn_a, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)


                new_a.MoveFirst()
                While Not new_a.EOF
                    con_merg.Execute("INSERT INTO new_a (ID_Tovar, ID_Analog, sklad) VALUES (" & new_a.Fields(0).Value & ", " & new_a.Fields(1).Value & ", " & itera & ");")
                    new_a.MoveNext()
                End While
                new_a = Nothing
                Console.WriteLine("Import new Successful")

                old_a.MoveFirst()
                While Not old_a.EOF
                    con_merg.Execute("INSERT INTO old_a (ID_Tovar, ID_Analog, sklad) VALUES (" & old_a.Fields(0).Value & ", " & old_a.Fields(1).Value & ", " & itera & ");")
                    old_a.MoveNext()
                End While

                Console.WriteLine("Import old Successful")

                old_a = Nothing

                'Dim ass, ass2 As Integer
                'Do Until (new_a.EOF)
                '    old_a.MoveFirst()
                '    old_a.Find(new_a.Fields(0).Value = old_a.Fields(0).Value)
                '    If (old_a.EOF) And (ass2 < ass) Then
                '        Console.WriteLine("Найдено: " & new_a.Fields(0).Value)
                '        ' Debug.Print("Найдено: " & ost("ID_Tovar"))
                '        ass = ass + 1
                '    End If
                '    new_a.MoveNext()
                '    ass2 = ass2 + 1
                'Loop
                'Console.WriteLine("Всего: " & ass - 1)

                'ValuesArr = ""
                'TotalTovarov = 0
                'comma = ""
                'InsertValues = "INSERT INTO " & tbl & " (id_tovar, id_analog, sklad) VALUES "


                'While Not new_a.EOF
                '    If old_a.EOF = True Then 'старый уже кончился, а новый - еще нет
                '        values(0) = new_a.Fields(0).Value
                '        values(1) = new_a.Fields(1).Value
                '        differ.AddNew(fieldsArray, values)
                '        differ.Update()
                '    End If

                '    If ((new_a.Fields(0).Value & new_a.Fields(1).Value <> old_a.Fields(0).Value & old_a.Fields(1).Value) Or (old_a.Fields(0).Value = Nothing)) Then

                '        ' Console.WriteLine(data.Fields(0).Value & "/" & data.Fields(1).Value)
                '    End If


                '    'If TotalTovarov = 800 Then
                '    '    comma2 = ";"
                '    'Else
                '    '    comma2 = ""
                '    'End If
                '    'If itera = 3 Then
                '    '    id_tovara = data.Fields(0).Value - 29894
                '    '    id_analoga = data.Fields(1).Value - 29894
                '    'Else
                '    '    id_tovara = data.Fields(0).Value
                '    '    id_analoga = data.Fields(1).Value
                '    'End If
                '    'ValuesArr = ValuesArr & comma & " (" & data.Fields(0).Value & "," & data.Fields(1).Value & "," & itera & ")" & comma2
                '    'comma = ","
                '    'If TotalTovarov = 800 Then
                '    '    cmd.CommandText = InsertValues & ValuesArr
                '    '    cmd.Prepare()
                '    '    'cmd.ExecuteNonQuery()

                '    '    Console.Write(".")
                '    '    ValuesArr = ""
                '    '    TotalTovarov = 0
                '    '    comma = ""
                '    '    comma2 = ""
                '    'End If
                '    new_a.MoveNext()
                '    old_a.MoveNext()
                '    TotalTovarov = TotalTovarov + 1
                'End While

                'cmd.CommandText = InsertValues & ValuesArr & ";"
                'cmd.Prepare()
                'cmd.ExecuteNonQuery()

                '   InsertValues = Nothing
                '   ValuesArr = ""
                '   comma = Nothing

                '  Console.WriteLine("")
                '  Console.WriteLine("Аналоги склада " & itera & " обновлены")
                ' Console.WriteLine("")
                'System.IO.File.AppendAllText(logf, "a" + itera.ToString + Environment.NewLine)



                'ost = New ADODB.Recordset
                'ost.Open("Export_tovari1", cn1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                'ost.MoveFirst()
                'ValuesArr = ""
                'TotalTovarov = 0
                'comma = ""
                'InsertValues = "INSERT INTO `conn_tovar` (`id_tovar`, `code`, `demo`, `cross_num`, `proizv`, `qnt`, `hidecode`, `price`, `roznica`, `comment`, `zakup`, `red`, `sklad`) VALUES "

                'Dim code, demo, cross_num, proizv, comment, red, price, roznica, zakup As String
                'Dim id_tovar, hidecode, sklad, qnt As Integer

                'While Not ost.EOF
                '    If TotalTovarov = 500 Then
                '        comma2 = ";"
                '    Else
                '        comma2 = ""
                '    End If

                '    If itera = 3 Then
                '        id_tovar = ost.Fields(0).Value - 29894 'id_tovar
                '    Else
                '        id_tovar = ost.Fields(0).Value 'id_tovar
                '    End If


                '    code = LTrim(ost.Fields(1).Value) 'code

                '    code = preparer(code)

                '    demo = ""
                '    If ost.Fields(2).Value.Equals(System.DBNull.Value) Then 'comment
                '        comment = ""
                '    Else
                '        comment = Replace(Left(ost.Fields(2).Value, 255), "'", "")
                '    End If
                '    If ost.Fields(3).Value.Equals(System.DBNull.Value) Then 'cross
                '        cross_num = ""
                '    Else
                '        cross_num = Replace(Left(ost.Fields(3).Value, 255), "'", "")
                '    End If
                '    If ost.Fields(4).Value.Equals(System.DBNull.Value) Then 'price_eur
                '        price = 0
                '    Else
                '        price = ost.Fields(4).Value
                '    End If

                '    If ost.Fields(5).Value.Equals(System.DBNull.Value) Then 'qnt
                '        qnt = 0
                '    Else
                '        qnt = ost.Fields(5).Value
                '    End If

                '    If ost.Fields(6).Value.Equals(System.DBNull.Value) Then 'rozn
                '        roznica = 0
                '    Else
                '        roznica = ost.Fields(6).Value
                '    End If
                '    If ost.Fields(7).Value.Equals(System.DBNull.Value) Then 'proizv
                '        proizv = "LIC"
                '    Else
                '        proizv = ost.Fields(7).Value
                '    End If
                '    If ost.Fields(8).Value.Equals(System.DBNull.Value) Then 'hidecode
                '        hidecode = ""
                '    Else
                '        hidecode = ost.Fields(8).Value
                '    End If
                '    If ost.Fields(9).Value.Equals(System.DBNull.Value) Then 'zakup
                '        zakup = ""
                '    Else
                '        zakup = ost.Fields(9).Value
                '    End If
                '    If ost.Fields(10).Value.Equals(System.DBNull.Value) Then 'red
                '        red = ""
                '    Else
                '        red = Left(ost.Fields(10).Value, 255)
                '    End If
                '    sklad = itera

                '    ValuesArr = ValuesArr & comma & " ('" & id_tovar & "','" & code & "','" & demo & "','" & cross_num & "','" & proizv & "','" & qnt & "','" & hidecode & "','" & price & "','" & roznica & "','" & comment & "','" & zakup & "','" & red & "','" & sklad & "')" & comma2
                '    comma = ","
                '    If TotalTovarov = 500 Then

                '        cmd.CommandText = InsertValues & ValuesArr
                '        cmd.Prepare()
                '        cmd.ExecuteNonQuery()

                '        Console.Write(".")
                '        ValuesArr = ""
                '        TotalTovarov = 0
                '        comma = ""
                '        comma2 = ""
                '    End If
                '    ost.MoveNext()
                '    TotalTovarov = TotalTovarov + 1
                'End While

                'cmd.CommandText = InsertValues & ValuesArr & ";"
                'cmd.Prepare()
                'cmd.ExecuteNonQuery()

                'Console.WriteLine("")
                'Console.WriteLine("Товары склада " & itera & " обновлены")
                'Console.WriteLine("")
                'System.IO.File.AppendAllText(logf, " / t" + itera.ToString + Environment.NewLine)
                'ost.Close()
                'cn = Nothing
                'cn1 = Nothing
                'rst = Nothing
                'ost = Nothing
            Next
            con_merg.Close()

            con_merg.Open("Provider=Microsoft.ACE.OLEDB.12.0; data source=" & merged_db_path & ";")
            differ.Open(added_analogs, con_merg, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If differ.RecordCount > 0 Then

                ' появились новые аналоги



                'nw.id_tovar | nw.id_analog | ol.id_tovar | ol.id_analog
                '_______________________________________________________
                '   171534   |   172755     |             |
                '   172755   |   171534     |             |
                '_______________________________________________________


                ' удален аналог

                'nw.id_tovar	nw.id_analog	ol.id_tovar	ol.id_analog
                '                                   171534	    172755
                '                                   172755	    171534


                ' 1 добавлен и 1 удален

                'nw.id_tovar	nw.id_analog	ol.id_tovar	ol.id_analog
                '   171250	        174779		
                '   174779	        171250		
                '                                   172755	    179382
                '                                   179382	    172755


                'nw.id_tovar	nw.id_analog	ol.id_tovar	ol.id_analog	nw.sklad	ol.sklad
                '   170002	        171534			                            2	
                '   171534	        170002			                            2	
                '   171534	        175933			                            2	
                '   175933	        171534			                            2	




                System.IO.File.AppendAllText(logf, differ.RecordCount.ToString + " новых аналогов добавлено " + Environment.NewLine)


                Dim conn As New MySqlConnection
                conn.ConnectionString = connectionstring
                conn.Open()
                Dim cmd As New MySqlCommand
                cmd.Connection = conn

                Dim SqlQry = "DROP TABLE IF EXISTS `last_export`;" & _
                                "CREATE TABLE IF NOT EXISTS `last_export` (" & _
                                "`datetime` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP) ENGINE=InnoDB;"

                cmd.CommandText = SqlQry
                cmd.Prepare()
                cmd.ExecuteNonQuery()

                Dim nw_tovar, nw_analog, ol_tovar, ol_analog, nws, ols As Integer

                While Not differ.EOF

                    If differ.Fields(1).Value.Equals(System.DBNull.Value) Then
                        nw_tovar = 0
                    Else
                        nw_tovar = differ.Fields(1).Value
                    End If
                    If differ.Fields(2).Value.Equals(System.DBNull.Value) Then
                        nw_analog = 0
                    Else
                        nw_analog = differ.Fields(2).Value
                    End If
                    If differ.Fields(3).Value.Equals(System.DBNull.Value) Then
                        ol_tovar = 0
                    Else
                        ol_tovar = differ.Fields(3).Value
                    End If
                    If differ.Fields(4).Value.Equals(System.DBNull.Value) Then
                        ol_analog = 0
                    Else
                        ol_analog = differ.Fields(4).Value
                    End If
                    If differ.Fields(5).Value.Equals(System.DBNull.Value) Then
                        nws = 0
                    Else
                        nws = differ.Fields(5).Value
                    End If
                    If differ.Fields(6).Value.Equals(System.DBNull.Value) Then
                        ols = 0
                    Else
                        ols = differ.Fields(6).Value
                    End If

                    Dim analogi_update_qry As String

                    If (nw_tovar + nw_analog) > 0 Then ' добавлены аналоги
                        analogi_update_qry = "INSERT INTO `conn_analogi` (id_tovar, id_analog, sklad) VALUES (" & nw_tovar & ", " & nw_analog & ", " & nws & ");"
                        
                    ElseIf (ol_tovar + ol_analog) > 0 Then ' удалены аналоги
                        analogi_update_qry = "DELETE FROM `conn_analogi` WHERE `id_tovar` = '" & ol_tovar & "' AND `id_analog` = '" & ol_analog & "' AND `sklad` = '" & ols & "';"

                    Else ' нихуя непонятно что

                    End If

                    cmd.CommandText = analogi_update_qry
                    cmd.Prepare()
                    cmd.ExecuteNonQuery()

                End While

                Console.WriteLine("Данные обновлены")
            Else
                System.IO.File.AppendAllText(logf, "Нечего добавить" + Environment.NewLine)
            End If


            'If differ.RecordCount > 0 Then
            '    System.IO.File.AppendAllText(logf, Environment.NewLine + "Че-то найдено: " + differ.RecordCount.ToString + Environment.NewLine)
            '    differ.MoveFirst()

            '    While Not differ.EOF
            '        Console.WriteLine(differ.Fields(0).Value)
            '        'System.IO.File.AppendAllText(logf, differ.Fields(0).Value + " / " + differ.Fields(1).Value + Environment.NewLine)
            '        differ.MoveNext()
            '    End While

            '    System.IO.File.AppendAllText(logf, Environment.NewLine + "Всего записей: " + differ.RecordCount.ToString + Environment.NewLine)

            'Else
            '    System.IO.File.AppendAllText(logf, Environment.NewLine + "Данные не требуют обновления. Ахуеть!" + Environment.NewLine)
            'End If

            differ.Close()


            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = `code`;"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()
            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '-','');"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()
            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '.','');"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()
            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '/','');"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()
            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, ' ','');"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()
            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '+','');"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()
            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '-','');"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()
            'cmd.CommandText = "UPDATE `conn_tovar` SET `demo` = REPLACE(`demo`, '-','');" & _
            '                    "DROP TABLE `tovar`;" & _
            '                    "RENAME TABLE  `conn_tovar` TO `tovar`;" & _
            '                    "DROP TABLE `analogi`;" &
            '                    "RENAME TABLE  `conn_analogi` TO `analogi`;"
            'cmd.Prepare()
            'cmd.ExecuteNonQuery()

            'System.IO.File.AppendAllText(logf, "Импорт завершен" + Environment.NewLine)

        Catch ex As Exception
            System.IO.File.AppendAllText(logf, Now & " " & ex.Message.ToString + Environment.NewLine)
        End Try
    End Sub
End Module
