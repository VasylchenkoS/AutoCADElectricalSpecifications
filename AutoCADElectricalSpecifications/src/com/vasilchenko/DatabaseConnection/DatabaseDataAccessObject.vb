Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules

Namespace com.vasilchenko.DatabaseConnection
    Module DatabaseDataAccessObject
        Private ReadOnly DatabaseTableList As New List(Of String)
        Public Sub FillItemData(ByVal inputList As List(Of DrawingSpecificationItem), Optional strSqlQuery As String = "")
            If strSqlQuery = "" Then
                strSqlQuery = "SELECT [TAGNAME], [MFG], [CAT], [INST], [LOC], [FAMILY], [CNT], [UM], [DESC3], [PAR1_CHLD2] " &
                              "FROM [COMP] " &
                              "WHERE [TAGNAME] IN (" &
                                "SELECT [TAGNAME] " &
                                "FROM [COMP] " &
                                "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) AND [TAGNAME] IS NOT NULL) " &
                              "AND (([PAR1_CHLD2] = '1' OR [PAR1_CHLD2] BETWEEN '301' AND '399') AND [MFG] <> 'IGNORED') " &
                              "ORDER BY [INST], [MFG], [TAGNAME], [PAR1_CHLD2]"
            End If

            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        If Not IsDBNull(.item("TAGNAME")) Then
                            If .item("TAGNAME").ToString.ToLower.Contains("ignor") Then
                                currentItem.AddTag("-")
                            Else
                                currentItem.AddTag(.item("TAGNAME"))
                            End If
                        End If
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("CNT")) AndAlso .item("CNT") <> "" Then
                            currentItem.Count = .item("CNT")
                        Else
                            currentItem.Count = 1
                        End If
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetLastNumericFromString(.item("INST")) Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.CatalogName & ". Не указана Ф-я группа")
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC") Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.Tag & ". Не указано местоположение")
                        If Not IsDBNull(.item("DESC3")) Then currentItem.Note = .item("DESC3")
                        If Not IsDBNull(.item("FAMILY")) Then currentItem.Family = .item("FAMILY")
                        If Not IsDBNull(.item("PAR1_CHLD2")) AndAlso CInt(.item("PAR1_CHLD2")) > 300 Then
                            If Not IsNothing(inputList.Find(Function(x) x.Tag = currentItem.Tag AndAlso Not x.Note.Equals(""))) Then _
                                currentItem.Note = inputList.Find(Function(x) x.Tag = currentItem.Tag AndAlso Not x.Note.Equals("")).Note
                        End If
                    End With
                    FillItemDescription(currentItem, inputList)
                    If Not currentItem.Description Is Nothing Then AddItemToList(inputList, currentItem)
                Next objRow
            End If
        End Sub

        Friend Sub FillMountItemData(ByRef inputList As List(Of DrawingSpecificationItem), Optional strSqlQuery As String = "")
            If strSqlQuery = "" Then
                strSqlQuery = "SELECT [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT], [UM], [DESC3], [TAGNAME], [PAR1_CHLD2] " &
                              "FROM [PNLCOMP] " &
                              "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " &
                                "AND ([PAR1_CHLD2] = '1' OR [PAR1_CHLD2] BETWEEN '301' AND '399') " &
                                "AND [MFG] <> 'IGNORED' " &
                                "AND ([TAGNAME] IS NOT NULL " &
                                    "OR ([TAGNAME] IS NULL AND [WDBLKNAM] <> 'DIN' AND [WDBLKNAM] <> 'WW' AND [WDBLKNAM] <> 'DI')) " &
                              "AND NOT EXISTS ( " &
                                "SELECT [TAGNAME] " &
                                "FROM [COMP]  " &
                                "WHERE [TAGNAME] IN (" &
                                    "SELECT [TAGNAME]  " &
                                    "FROM [COMP]  " &
                                    "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) AND [TAGNAME] IS NOT NULL) " &
                                "AND (([PAR1_CHLD2] = '1' OR [PAR1_CHLD2] BETWEEN '301' AND '399')  " &
                                "AND [MFG] <> 'IGNORED') " &
                                "AND [PNLCOMP].[TAGNAME] = [COMP].[TAGNAME])" &
                                "ORDER BY [INST], [MFG], [TAGNAME], [PAR1_CHLD2]"
            End If

            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        'currentItem.AddTag("-")
                        'currentItem.Family = "EN"
                        'If Not IsDBNull(.item("WDBLKNAM")) Then currentItem.Family = .item("WDBLKNAM")
                        If Not IsDBNull(.item("TAGNAME")) Then currentItem.AddTag(.item("TAGNAME")) Else currentItem.AddTag("-")
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetLastNumericFromString(.item("INST")) Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.CatalogName & ". Не указана Ф-я группа")
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC") Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.Tag & ". Не указано местоположение")
                        If Not IsDBNull(.item("DESC3")) Then currentItem.Note = .item("DESC3")
                        If Not IsDBNull(.item("CNT")) AndAlso .item("CNT") <> "" Then
                            currentItem.Count = .item("CNT")
                        Else
                            currentItem.Count = 1
                        End If
                        If Not IsDBNull(.item("PAR1_CHLD2")) AndAlso CInt(.item("PAR1_CHLD2")) > 300 Then
                            If Not IsNothing(inputList.Find(Function(x) x.Tag = currentItem.Tag AndAlso Not x.Note.Equals(""))) Then _
                                currentItem.Note = inputList.Find(Function(x) Not x.Tag.Equals("-") AndAlso x.Tag = currentItem.Tag AndAlso Not x.Note.Equals("")).Note
                        End If
                    End With
                    FillItemDescription(currentItem, inputList)
                    If Not currentItem.Description Is Nothing Then AddItemToList(inputList, currentItem)
                Next objRow
            End If
        End Sub

        Public Sub FillPageTerminalData(ByRef inputList As List(Of DrawingSpecificationItem), Optional strSqlQuery As String = "")

            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        If Not IsDBNull(.item("TAGSTRIP")) Then currentItem.AddTag(.item("TAGSTRIP"))
                        currentItem.Family = "TRMS"
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetLastNumericFromString(.item("INST")) Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.CatalogName & ". Не указана Ф-я группа")
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC") Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.Tag & ". Не указано местоположение")
                        If Not IsDBNull(.item("CNT")) AndAlso .item("CNT") <> "" Then
                            currentItem.Count = .item("CNT")
                        Else
                            currentItem.Count = 1
                        End If
                    End With
                    FillItemDescription(currentItem, inputList)
                    If Not currentItem.Description Is Nothing Then AddItemToList(inputList, currentItem)
                Next objRow
            End If
        End Sub

        Public Sub FillMountTerminalData(ByRef inputList As List(Of DrawingSpecificationItem), Optional strSqlQuery As String = "")
            If strSqlQuery = "" Then
                strSqlQuery = "SELECT [TAGSTRIP], [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT] " &
                              "FROM [PNLTERM] " &
                              "WHERE [INST] NOT LIKE '9[0-9]%' " &
                                "AND [PAR1_CHLD2] = '1' " &
                                "AND [MFG] <> 'IGNORED'"
            End If
            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        If Not IsDBNull(.item("TAGSTRIP")) Then currentItem.AddTag(.item("TAGSTRIP"))
                        If Not IsDBNull(.item("WDBLKNAM")) Then currentItem.Family = .item("WDBLKNAM")
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetLastNumericFromString(.item("INST")) Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.CatalogName & ". Не указана Ф-я группа")
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC") Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.Tag & ". Не указано местоположение")
                        If Not IsDBNull(.item("CNT")) AndAlso .item("CNT") <> "" Then
                            currentItem.Count = .item("CNT")
                        Else
                            currentItem.Count = 1
                        End If
                    End With
                    FillItemDescription(currentItem, inputList)
                    If Not currentItem.Description Is Nothing Then AddItemToList(inputList, currentItem)
                Next objRow
            End If

            If inputList.Count = 0 Then Exit Sub

            strSqlQuery = "SELECT [TAGSTRIP], [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT] " &
                          "FROM [TERMS] " &
                          "WHERE [INST] NOT LIKE '9[0-9]%' " &
                            "AND Left([PAR1_CHLD2],1) LIKE '3' " &
                            "AND [MFG] <> 'IGNORED'"

            objDataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        If Not IsDBNull(.item("TAGSTRIP")) Then currentItem.AddTag(.item("TAGSTRIP"))
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetLastNumericFromString(.item("INST")) Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.CatalogName & ". Не указана Ф-я группа")
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC") Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.Tag & ". Не указано местоположение")
                        If Not IsDBNull(.item("CNT")) AndAlso .item("CNT") <> "" Then
                            currentItem.Count = .item("CNT")
                        Else
                            currentItem.Count = 1
                        End If
                    End With
                    FillItemDescription(currentItem, inputList)
                    If Not currentItem.Description Is Nothing Then AddItemToList(inputList, currentItem)
                Next objRow
            End If

            If inputList.Exists(Function(x) x.CatalogName.StartsWith("UT") Or x.CatalogName.StartsWith("ST 2,5") And x.CatalogName.Contains("2,5") And x.Instance = 61) Then
                Dim shtCount = 0.0
                Dim taglist As New List(Of String)
                Dim item As New DrawingSpecificationItem
                For Each temp In inputList.FindAll(Function(x) x.CatalogName.StartsWith("UT") Or x.CatalogName.StartsWith("ST") And x.CatalogName.Contains("2,5") And x.Instance = 61)
                    If temp.CatalogName.StartsWith("UT 2,5") Or temp.CatalogName.StartsWith("ST 2,5") Then
                        shtCount += temp.Count * 2
                    ElseIf temp.CatalogName.StartsWith("UTTB 2,5") Or temp.CatalogName.StartsWith("STTB 2,5") Then
                        shtCount += temp.Count * 4
                    End If

                    TerminalMarkingTagAdd(taglist, temp)

                Next
                item.Location = inputList.Find(Function(x) x.CatalogName.StartsWith("UT") Or x.CatalogName.StartsWith("ST") And x.CatalogName.Contains("2,5") And x.Instance = 61).Location
                item.Instance = AdditionalFunctions.GetLastNumericFromString("64.КЛЕММЫ-МАРКИРОВКА")
                taglist.ForEach(Sub(x) item.AddTag(x))
                item.Family = "TRMS"
                item.Article = "0828734"
                item.CatalogName = "UCT-TM 5"
                item.Description = "Маркувальна пластина (72 ел.) для клем"
                item.Unit = "шт."
                item.Manufacture = "Phoenix Contact"
                item.Count = System.Math.Ceiling(shtCount / 72)
                FillItemDescription(item, inputList)
                If Not item.Description Is Nothing Then AddItemToList(inputList, item)
            End If

            If inputList.Exists(Function(x) x.CatalogName.StartsWith("UT") And x.CatalogName.Contains("4") And x.Instance = 61) Then
                Dim shtCount = 0.0
                Dim taglist As New List(Of String)
                Dim item As New DrawingSpecificationItem
                For Each temp In inputList.FindAll(Function(x) x.CatalogName.StartsWith("UT") And x.CatalogName.Contains("4") And x.Instance = 61)
                    If temp.CatalogName.StartsWith("UT 4") Then
                        shtCount += temp.Count * 2
                    ElseIf temp.CatalogName.StartsWith("UTTB 4") Then
                        shtCount += temp.Count * 4
                    End If
                    TerminalMarkingTagAdd(taglist, temp)
                Next
                item.Location = inputList.Find(Function(x) x.CatalogName.StartsWith("UT") And x.CatalogName.Contains("4") And x.Instance = 61).Location
                item.Instance = AdditionalFunctions.GetLastNumericFromString("64.КЛЕММЫ-МАРКИРОВКА")
                taglist.ForEach(Sub(x) item.AddTag(x))
                item.Family = "TRMS"
                item.Article = "0828736"
                item.CatalogName = "UCT-TM 6"
                item.Description = "Маркувальна пластина (60 ел.) для клем"
                item.Unit = "шт."
                item.Manufacture = "Phoenix Contact"
                item.Count = System.Math.Ceiling(shtCount / 60)
                FillItemDescription(item, inputList)
                If Not item.Description Is Nothing Then AddItemToList(inputList, item)

            End If

            If inputList.Exists(Function(x) (x.CatalogName.StartsWith("UT") Or x.CatalogName.StartsWith("URT")) And x.CatalogName.Contains("6") And x.Instance = 61) Then
                Dim shtCount As Double = 0.0
                Dim taglist As New List(Of String)
                Dim item As New DrawingSpecificationItem
                For Each temp In inputList.FindAll(Function(x) (x.CatalogName.StartsWith("UT") Or x.CatalogName.StartsWith("URT")) And x.CatalogName.Contains("6") And x.Instance = 61)
                    shtCount += temp.Count * 2
                    TerminalMarkingTagAdd(taglist, temp)
                Next
                item.Location = inputList.Find(Function(x) (x.CatalogName.StartsWith("UT") Or x.CatalogName.StartsWith("URT")) And x.CatalogName.Contains("6") And x.Instance = 61).Location
                item.Instance = AdditionalFunctions.GetLastNumericFromString("64.КЛЕММЫ-МАРКИРОВКА")
                taglist.ForEach(Sub(x) item.AddTag(x))
                item.Family = "TRMS"
                item.Article = "0828740"
                item.CatalogName = "UCT-TM 8"
                item.Description = "Маркувальна пластина (42 ел.) для клем"
                item.Unit = "шт."
                item.Manufacture = "Phoenix Contact"
                item.Count = System.Math.Ceiling(shtCount / 42)
                FillItemDescription(item, inputList)
                If Not item.Description Is Nothing Then AddItemToList(inputList, item)
            End If

            If inputList.Exists(Function(x) x.CatalogName.StartsWith("UT") And x.CatalogName.Contains("16") And x.Instance = 61) Then
                Dim shtCount = 0.0
                Dim taglist As New List(Of String)
                Dim item As New DrawingSpecificationItem
                For Each temp In inputList.FindAll(Function(x) x.CatalogName.StartsWith("UT") And x.CatalogName.Contains("16") And x.Instance = 61)
                    shtCount += temp.Count * 2
                    TerminalMarkingTagAdd(taglist, temp)
                Next
                item.Location = inputList.Find(Function(x) x.CatalogName.StartsWith("UT") And x.CatalogName.Contains("16") And x.Instance = 61).Location
                item.Instance = AdditionalFunctions.GetLastNumericFromString("64.КЛЕММЫ-МАРКИРОВКА")
                taglist.ForEach(Sub(x) item.AddTag(x))
                item.Family = "TRMS"
                item.Article = "0829144"
                item.CatalogName = "UCT-TM 12"
                item.Description = "Маркувальна пластина (30 ел.) для клем"
                item.Unit = "шт."
                item.Manufacture = "Phoenix Contact"
                item.Count = System.Math.Ceiling(shtCount / 30)
                FillItemDescription(item, inputList)
                If Not item.Description Is Nothing Then AddItemToList(inputList, item)
            End If
        End Sub

        Private Sub TerminalMarkingTagAdd(taglist As List(Of String), temp As DrawingSpecificationItem)
            Dim tagArray() As String = Split(temp.Tag, ", ",)
            For Each str As String In tagArray
                taglist.Add(str)
            Next
        End Sub

        Private Sub FillItemDescription(ByRef objInput As DrawingSpecificationItem, ByRef inputList As List(Of DrawingSpecificationItem))
            Dim objDataTable As Data.DataTable
            Dim strSqlQuery = ""

            DatabaseTableListCheck()

            If objInput.Family = "" OrElse objInput.Family.ToLower = "cbl" Then
                For Each strCurtable In DatabaseTableList
                    strSqlQuery = "select [CATALOG] from " & strCurtable & " where [CATALOG] = '" & objInput.CatalogName & "'"
                    objDataTable = DatabaseConnections.GetSqlDataTable(strSqlQuery, My.Settings.default_catSQLConnectionString, True)
                    If Not IsNothing(objDataTable) AndAlso objDataTable.Rows.Count <> 0 Then
                        objInput.Family = strCurtable
                        Exit For
                    End If
                Next
                If objInput.Family = Nothing OrElse objInput.Family.Equals("") Then
                    MsgBox("Строка в базе не найдена. Строка запроса:" & strSqlQuery, MsgBoxStyle.Critical)
                    Exit Sub
                End If
            End If

            If objInput.Family.ToLower().Contains("plcio") And Not objInput.Family.ToLower().Equals("plcio") Then
                objInput.Family = "PLCIO"
            End If

            strSqlQuery = "SELECT [DESCRIPTION], [USER1], [USER2], [USER3],[ASSEMBLYCODE] " &
                          "FROM [" & objInput.Family & "] " &
                          "WHERE [CATALOG] = '" & objInput.CatalogName & "'"

            objDataTable = DatabaseConnections.GetSqlDataTable(strSqlQuery, My.Settings.default_catSQLConnectionString, False)
            If Not IsNothing(objDataTable) Then

                objInput.Description = objDataTable.Rows(0).Item("DESCRIPTION")
                If Not IsDBNull(objDataTable.Rows(0).Item("USER1")) Then objInput.Article = objDataTable.Rows(0).Item("USER1")
                If Not IsDBNull(objDataTable.Rows(0).Item("USER2")) Then
                    objInput.Unit = objDataTable.Rows(0).Item("USER2")
                Else
                    MsgBox("В Элементе " & objInput.Manufacture & " " & objInput.CatalogName & " не прописана ед. измерения. Установлено по-умолчанию")
                    objInput.Unit = "шт."
                End If

                If Not IsDBNull(objDataTable.Rows(0).Item("USER3")) Then objInput.Note = objDataTable.Rows(0).Item("USER3") & ". " & objInput.Note
                If Not IsDBNull(objDataTable.Rows(0).Item("ASSEMBLYCODE")) Then
                    Dim strAssemblyCode = objDataTable.Rows(0).Item("ASSEMBLYCODE")
                    inputList = AddAssembly(objInput, inputList, strAssemblyCode)
                End If
            Else
                objInput.Family = ""
                FillItemDescription(objInput, inputList)
            End If
        End Sub

        Private Function AddAssembly(objInput As DrawingSpecificationItem, inputList As List(Of DrawingSpecificationItem), strAssemblyCode As String) As List(Of DrawingSpecificationItem)
            Dim strSqlQuery As String
            Dim objDataTable As Data.DataTable

            strSqlQuery = "SELECT CATALOG, MANUFACTURER, ASSEMBLYQUANTITY " &
                          "FROM " & objInput.Family & " " &
                          "WHERE ASSEMBLYLIST like '%" & strAssemblyCode & "%'"

            objDataTable = DatabaseConnections.GetSqlDataTable(strSqlQuery, My.Settings.default_catSQLConnectionString, False)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim objNewItem As New DrawingSpecificationItem
                    With objRow
                        If Not IsDBNull(.item("MANUFACTURER")) Then objNewItem.Manufacture = .item("MANUFACTURER")
                        If Not IsDBNull(.item("CATALOG")) Then objNewItem.CatalogName = .item("CATALOG")
                        objNewItem.AddTag(objInput.Tag)
                        objNewItem.Instance = objInput.Instance
                        objNewItem.Location = objInput.Location
                        objNewItem.Note = objInput.Note
                        objNewItem.Family = objInput.Family
                        If Not IsDBNull(.item("ASSEMBLYQUANTITY")) Then
                            objNewItem.Count = objInput.Count * .item("ASSEMBLYQUANTITY")
                        Else
                            Throw New Exception("Не указано количество сборки", New ArgumentNullException)
                        End If
                    End With
                    FillItemDescription(objNewItem, inputList)
                    If Not objNewItem.Description Is Nothing Then AddItemToList(inputList, objNewItem)
                Next objRow
            End If
            Return inputList
        End Function

        Private Sub AddItemToList(inputList As List(Of DrawingSpecificationItem), objInput As DrawingSpecificationItem)
            If inputList.Exists(Function(x) x.CatalogName.Equals(objInput.CatalogName) And x.Location.Equals(objInput.Location) AndAlso x.Description.Equals(objInput.Description)) Then
                Dim findItem = inputList.Find(Function(x) x.CatalogName.Equals(objInput.CatalogName) And x.Location.Equals(objInput.Location))
                If objInput.Family = "DI" Then
                    findItem.AddLength(objInput.CountList(0))
                    findItem.Count += objInput.Count
                Else
                    findItem.AddTag(objInput.Tag)
                    findItem.Count += objInput.Count
                    If objInput.Note <> "" And Not findItem.Note.Contains(objInput.Note) Then
                        If findItem.Note <> "" Then
                            findItem.Note = findItem.Note & ". " & objInput.Note
                        Else
                            findItem.Note = objInput.Note
                        End If
                    End If
                End If
            Else
                inputList.Add(objInput)
            End If
        End Sub

        Private Sub DatabaseTableListCheck()
            Dim objDataTable As Data.DataTable
            If DatabaseTableList.Count = 0 Then
                Const strSqlQuery As String = "select name " &
                                              "From sys.tables " &
                                              "where name Not Like '#_%' ESCAPE '#' and name not like 'DataModel'"
                objDataTable = DatabaseConnections.GetSqlDataTable(strSqlQuery, My.Settings.default_catSQLConnectionString, False)
                If Not IsNothing(objDataTable) Then
                    For Each objRow In objDataTable.Rows
                        DatabaseTableList.Add(objRow.Item("name"))
                    Next
                End If
            End If
        End Sub

        Friend Sub FillSheetDinAndDuctItemData(ByRef inputList As List(Of DrawingSpecificationItem), Optional strSqlQuery As String = "")
            If strSqlQuery = "" Then
                strSqlQuery = "SELECT [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [BLOCK], [CNT] " &
                              "FROM [PNLCOMP] " &
                              "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " &
                                "AND ([WDBLKNAM] = 'DIN' OR [WDBLKNAM] = 'WW' OR [WDBLKNAM] = 'DI') " &
                                "AND ([CNT] IS NOT NULL)"
            End If
            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        currentItem.AddTag("-")
                        If Not IsDBNull(.item("BLOCK")) Then currentItem.BlockName = .item("BLOCK")
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetLastNumericFromString(.item("INST")) Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.Tag & " " & currentItem.CatalogName & ". Не указана Ф-я группа")
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC") Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.Tag & " " & currentItem.CatalogName & ". Не указано местоположение")
                        currentItem.Family = "DI"
                        currentItem.AddLength(.item("CNT"))
                        'currentItem.AddLength(GetLengthFromBlock(acDatabase, currentItem, acTransaction))
                    End With
                    FillItemDescription(currentItem, inputList)
                    If Not currentItem.Description Is Nothing Then AddItemToList(inputList, currentItem)
                Next objRow
            End If
            FillDinAndDuctCount(inputList)
        End Sub

        Private Sub FillDinAndDuctCount(ByRef inputList As List(Of DrawingSpecificationItem))
            For Each curItem In inputList.FindAll(Function(x) x.Family = "DI")
                Dim count As Short = 0
                Dim curCount = 0
                Dim maxLength = CDbl(AdditionalFunctions.GetLastNumericFromString(curItem.Description)) * 1000
                Dim countList = curItem.CountList

                For i As Short = 0 To countList.Count - 1
                    If countList.Count = 0 Then
                        If curCount > 0 Then count += 1
                        curItem.Count = count
                        Exit For
                    ElseIf curCount + countList(i) <= maxLength Then
                        curCount += countList(i)
                        countList.Remove(countList(i))
                        i -= 1
                    End If

                    If i = countList.Count - 1 Then
                        count += 1
                        curCount = 0
                        i = -1
                    End If
                Next
            Next
        End Sub

        Public Function FillCableItemData(Optional strSqlQuery As String = "") As SortedList(Of Short, List(Of CableItemClass))
            Dim acCableList As New SortedList(Of Short, List(Of CableItemClass))

            If strSqlQuery = "" Then
                strSqlQuery = "SELECT [FAMILY], [TAGNAME], [MFG], [CAT], [CNT], [INST], [LOC], [DESC2] " &
                              "FROM [COMP] " &
                              "WHERE [PAR1_CHLD2] = '1' AND [INST] LIKE '8[0-9]%' "
            End If

            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New CableItemClass
                    With objRow
                        If Not IsDBNull(.item("TAGNAME")) Then currentItem.TAG = .item("TAGNAME")
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("DESC2")) Then currentItem.TableDescription = .item("DESC2")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetLastNumericFromString(.item("INST")) Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.TAG & " " & currentItem.CatalogName & ". Не указана Ф-я группа")
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC") Else _
                            Throw New ArgumentNullException("Елемент " & currentItem.TAG & " " & currentItem.CatalogName & ". Не указано местоположение")
                        If Not IsDBNull(.item("CNT")) Then currentItem.Length = CInt(.item("CNT")) Else _
                            Throw New ArgumentNullException("Не указана длинна в кабеле " & currentItem.TAG)
                        currentItem.Unit = "м"
                    End With

                    strSqlQuery = "SELECT [DESCRIPTION], [USER1], [TYPE], [GAUGE], [MISCELLANEOUS1] " &
                                  "FROM [W0] " &
                                  "WHERE [CATALOG] = '" & currentItem.CatalogName & "'"

                    Dim objSqlDataTable = DatabaseConnections.GetSqlDataTable(strSqlQuery, My.Settings.default_catSQLConnectionString, False)
                    If Not IsNothing(objDataTable) Then
                        currentItem.Description = objSqlDataTable.Rows(0).Item("DESCRIPTION")
                        If Not IsDBNull(objSqlDataTable.Rows(0).Item("USER1")) Then currentItem.Article = objSqlDataTable.Rows(0).Item("USER1")
                        If Not IsDBNull(objSqlDataTable.Rows(0).Item("TYPE")) Then currentItem.Type = objSqlDataTable.Rows(0).Item("TYPE")
                        If Not IsDBNull(objSqlDataTable.Rows(0).Item("GAUGE")) Then currentItem.Section = objSqlDataTable.Rows(0).Item("GAUGE")
                        If Not IsDBNull(objSqlDataTable.Rows(0).Item("MISCELLANEOUS1")) Then currentItem.Cores = objSqlDataTable.Rows(0).Item("MISCELLANEOUS1")
                    End If

                    strSqlQuery = "SELECT [WIRENO], [LOC1], [LOC2], [NAM1], [PIN1], [NAM2], [PIN2] " &
                                  "FROM [WFRM2ALL] " &
                                  "WHERE [CBL] = '" & currentItem.TAG & "'"

                    Dim cblDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
                    If Not IsNothing(cblDataTable) Then
                        IdentifyCable(currentItem, cblDataTable)
                    End If

                    If IsNothing(currentItem.Source) OrElse IsNothing(currentItem.Destination) Then
                        MsgBox("Нет источника или приемника для кабеля " & currentItem.TAG)
                        Throw New Exception()
                    End If

                    If Not acCableList.ContainsKey(currentItem.Instance) Then acCableList.Add(currentItem.Instance, New List(Of CableItemClass))
                    acCableList.Values(acCableList.IndexOfKey(currentItem.Instance)).Add(currentItem)
                Next objRow
            End If

            Return acCableList
        End Function

        Private Sub IdentifyCable(currentItem As CableItemClass, cblDataTable As Data.DataTable)
            For Each cblRow In cblDataTable.Rows
                With cblRow

                    If Not IsDBNull(.item("WIRENO")) Then currentItem.AddWire(.item("WIRENO"))
                    If IsDBNull(.item("LOC1")) OrElse .item("LOC1").Equals("??") Then
                        MsgBox("Кабель " & currentItem.TAG & " не соединен со стороны " & .item("NAM1") & .item("PIN1"))
                        Throw New Exception()
                    ElseIf IsDBNull(.item("LOC2")) OrElse .item("LOC2").Equals("??") Then
                        MsgBox("Кабель " & currentItem.TAG & " не соединен со стороны " & .item("NAM2") & .item("PIN2"))
                        Throw New Exception()
                    End If

                    If currentItem.Source = "" And currentItem.Destination = "" Then
                        currentItem.Source = Trim(.item("LOC1"))
                        currentItem.Destination = Trim(.item("LOC2"))
                    Else
                        If (Not currentItem.Source.ToLower().Equals(Trim(.item("LOC1")).ToString().ToLower()) And
                                Not currentItem.Source.ToLower().Equals(Trim(.item("LOC2")).ToString().ToLower())) OrElse
                           (Not currentItem.Destination.ToLower().Equals(Trim(.item("LOC1")).ToString().ToLower()) And
                                Not currentItem.Destination.ToLower().Equals(Trim(.item("LOC2")).ToString().ToLower())) Then
                            MsgBox("Кабель " & currentItem.TAG & " идет в несколько панелей")
                            Throw New Exception()
                        End If
                    End If
                End With
            Next
        End Sub

    End Module
End Namespace