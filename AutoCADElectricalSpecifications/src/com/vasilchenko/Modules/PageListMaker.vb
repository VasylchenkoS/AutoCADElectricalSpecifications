Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.DatabaseConnection
Imports AutoCADElectricalSpecifications.com.vasilchenko.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.Modules
    Module PageListMaker
        Public Function SelectPages() As List(Of Integer)
            Dim pageList As New List(Of Integer)
            Dim uf As New PageSelector
            uf.ShowDialog()
            For Each item As Object In uf.lvSheets.SelectedItems
                pageList.Add(CInt(item.text))
            Next

            If pageList.Count = 0 Then Throw New Exception("Не выбраны листы для обработки", New NullReferenceException())

            uf.Dispose()

            Return pageList
        End Function

        Public Function MakePagesItemList(acDatabase As Database, acTransaction As Transaction, pageList As List(Of Integer)) As SortedList(Of String, List(Of DrawingSpecificationItem))
            Dim itemList As New SortedList(Of String, List(Of DrawingSpecificationItem))
            For Each i As Integer In pageList
                itemList = MakePageItemList(acDatabase, acTransaction, i, itemList)
            Next

            If itemList Is Nothing Then
                Throw New Exception
            End If
            Return itemList
        End Function

        Private Function MakePageItemList(acDatabase As Database, acTransaction As Transaction, i As Integer, resultList As SortedList(Of String, List(Of DrawingSpecificationItem))) As SortedList(Of String, List(Of DrawingSpecificationItem))
            Dim specItemList As New List(Of DrawingSpecificationItem)
            Dim tempSpecItemList As New List(Of DrawingSpecificationItem)
            Dim strSqlQuery As String
            'strSqlQuery = "SELECT [TAGNAME], [MFG], [CAT], [INST], [LOC], [FAMILY], [CNT], [UM], [DESC3], [PAR1_CHLD2] " & "FROM [COMP] " & "WHERE [TAGNAME] IN (" & "SELECT [TAGNAME] " & "FROM [COMP] " & "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " & "AND [TAGNAME] IS NOT NULL) " & "AND (([PAR1_CHLD2] = '1' OR [PAR1_CHLD2] BETWEEN '301' AND '399') AND [MFG] <> 'IGNORED') " & "AND [DWGIX] = '" & i & "'"
            strSqlQuery = "SELECT [TAGNAME], [MFG], [CAT], [INST], [LOC], [FAMILY], [CNT], [UM], [DESC3], [PAR1_CHLD2] FROM [COMP] WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) AND [TAGNAME] IS NOT NULL AND([PAR1_CHLD2] = '1' OR [PAR1_CHLD2] BETWEEN '301' AND '399') AND [MFG] <> 'IGNORED' AND [DWGIX] = '" & i & "' UNION SELECT [COMP].[TAGNAME], [COMP].[MFG], [COMP].[CAT], [COMP].[INST], [COMP].[LOC], [COMP].[FAMILY], [COMP].[CNT], [COMP].[UM], [COMP].[DESC3], [COMP].[PAR1_CHLD2] FROM [COMP] INNER JOIN [COMP] AS A 
ON [COMP].[TAGNAME] = A.[TAGNAME] WHERE A.[PAR1_CHLD2]='2' AND A.[DWGIX] = '" & i & "' AND [COMP].[PAR1_CHLD2]='1'"
            DatabaseDataAccessObject.FillItemData(specItemList, strSqlQuery)

            strSqlQuery = "SELECT [TAGNAME], [MFG], [CAT], [INST], [LOC], [FAMILY], [CNT], [UM], [DESC3], [PAR1_CHLD2] " & "FROM [COMP] " & "WHERE [TAGNAME] IN (" & "SELECT [TAGNAME] " & "FROM [COMP] " & "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " & "AND [TAGNAME] IS NOT NULL) " & "AND (([PAR1_CHLD2] = '2' OR [PAR1_CHLD2] BETWEEN '301' AND '399') AND [MFG] <> 'IGNORED') " & "AND [DWGIX] = '" & i & "'"
            DatabaseDataAccessObject.FillItemData(specItemList, strSqlQuery)

            strSqlQuery = "SELECT [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT], [UM], [DESC3], [TAGNAME] " & "FROM [PNLCOMP] " & "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " & "AND ([PAR1_CHLD2] = '1' OR [PAR1_CHLD2] = '301') " & "AND [MFG] <> 'IGNORED' " & "AND (([TAGNAME] IS NOT NULL AND [WDBLKNAM] = 'EN') " & "OR ([TAGNAME] IS NULL AND [WDBLKNAM] <> 'DIN' AND [WDBLKNAM] <> 'WW' AND [WDBLKNAM] <> 'DI') " & "OR ([TAGNAME] IS NOT NULL AND [WDBLKNAM] = 'WW' AND [CNT] IS NULL)) " & "AND [DWGIX] = '" & i & "'"
            DatabaseDataAccessObject.FillMountItemData(specItemList, strSqlQuery)

            strSqlQuery = "SELECT [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [BLOCK], [CNT] " & "FROM [PNLCOMP] " & "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " & "AND ([WDBLKNAM] = 'DIN' OR [WDBLKNAM] = 'WW' OR [WDBLKNAM] = 'DI') " & "AND ([CNT] IS NOT NULL) " & "AND [DWGIX] = '" & i & "'"
            DatabaseDataAccessObject.FillSheetDinAndDuctItemData(specItemList, strSqlQuery)

            strSqlQuery = "SELECT [TAGSTRIP], [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT] " & "FROM [PNLTERM] " & "WHERE [INST] NOT LIKE '9[0-9]%' " & "AND [PAR1_CHLD2] = '1' " & "AND [MFG] <> 'IGNORED' " & "AND [DWGIX] = '" & i & "'"
            DatabaseDataAccessObject.FillMountTerminalData(tempSpecItemList, strSqlQuery)

            If tempSpecItemList.Count = 0 Then
                strSqlQuery = "SELECT [TAGSTRIP], [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT] " & "FROM [TERMS] " & "WHERE [INST] NOT LIKE '9[0-9]%' " & "AND [PAR1_CHLD2] = '1' " & "AND [MFG] <> 'IGNORED' " & "AND [DWGIX] = '" & i & "'"
                DatabaseDataAccessObject.FillPageTerminalData(specItemList, strSqlQuery)
            Else
                DatabaseDataAccessObject.FillPageTerminalData(specItemList, strSqlQuery)
            End If

            Return MakeListForSheet(specItemList, resultList)
        End Function

        Private Function MakeListForSheet(inputList As List(Of DrawingSpecificationItem), Optional resultList As SortedList(Of String, List(Of DrawingSpecificationItem)) = Nothing) As SortedList(Of String, List(Of DrawingSpecificationItem))
            If IsNothing(resultList) Then resultList = New SortedList(Of String, List(Of DrawingSpecificationItem))

            For Each locations In inputList.GroupBy(Function(x) x.Location)
                Dim key As Short = locations(0).Instance
                Dim place As String = ""

                Dim locList = locations.ToList
                If key >= 800 AndAlso key >= 899 Then
                    For Each item In locList
                        item.Location = "Кабельно-провідникова продукція"
                    Next
                ElseIf key >= 700 AndAlso key >= 799 Then
                    For Each item In locList
                        item.Location = "Обладнання на ПС"
                    Next
                ElseIf Not resultList.ContainsKey(place) Then
                    place = locations(0).Location
                    Dim Response = MsgBox("Описание места в верхнем регистре? Пример:" & locations(0).Location, MsgBoxStyle.YesNo, "Описание места")
                    If Response = MsgBoxResult.No Then
                        For Each item In locList
                            item.Location = item.Location.Chars(0) & item.Location.ToLower.Remove(0, 1)
                        Next
                    End If
                End If

                If resultList.ContainsKey(place) Then
                    Dim curList = resultList.Item(place)
                    For Each element In locations
                        If curList.Exists(Function(x) x.CatalogName.Equals(element.CatalogName)) Then
                            With curList.Find(Function(x) x.CatalogName.Equals(element.CatalogName))
                                .AddTag(element.Tag)
                                .Count = .Count + element.Count
                            End With
                        Else
                            curList.Add(element)
                        End If
                    Next
                Else
                    resultList.Add(place, locList)
                End If
            Next

            Return resultList
        End Function
    End Module
End Namespace

