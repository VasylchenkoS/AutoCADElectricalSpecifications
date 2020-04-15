Imports System.Linq
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.DatabaseConnection
Imports Autodesk.AutoCAD.DatabaseServices
Namespace com.vasilchenko.Modules
    Module ProjectItemListMaker
        Public Function MakeProjectItemList(acDatabase As Database, acTransaction As Transaction) As SortedList(Of String, List(Of DrawingSpecificationItem))
            Dim specItemList As New List(Of DrawingSpecificationItem)
            Dim tempSpecItemList As New List(Of DrawingSpecificationItem)

            Dim strSqlQuery As String = "SELECT [TAGNAME], [MFG], [CAT], [INST], [LOC], [FAMILY], [CNT], [UM], [DESC3], [PAR1_CHLD2] " &
                "FROM [COMP] " &
                "WHERE [TAGNAME] IN (" &
                    "SELECT [TAGNAME] " &
                    "FROM [COMP] " &
                    "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) AND [TAGNAME] IS NOT NULL) " &
                "AND [PAR1_CHLD2] = '1' AND [MFG] <> 'IGNORED' AND ([DESC3] IS NULL OR ([DESC3] NOT LIKE '%снуюче%' AND [DESC3] NOT LIKE '%уществую%')) " &
                "UNION ALL " &
                "SELECT DISTINCT [TAGNAME], [MFG], [CAT], [INST], [LOC], [FAMILY], [CNT], [UM], [DESC3], [PAR1_CHLD2]  " &
                "FROM [COMP]  " &
                "WHERE " &
                    "([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL)  " &
                    "AND [TAGNAME] IS NOT NULL  " &
                    "AND [PAR1_CHLD2] BETWEEN '301' AND '399'  " &
                    "AND [MFG] <> 'IGNORED' " &
              "ORDER BY [INST], [MFG], [TAGNAME], [PAR1_CHLD2]"
            DatabaseDataAccessObject.FillItemData(specItemList, strSqlQuery)

            strSqlQuery = "SELECT [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT], [UM], [DESC3], [TAGNAME], [PAR1_CHLD2] " &
              "FROM [PNLCOMP] " &
              "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " &
                "AND ([PAR1_CHLD2] = '1' OR [PAR1_CHLD2] BETWEEN '301' AND '399') " &
                "AND [MFG] <> 'IGNORED'  AND ([DESC3] IS NULL OR ([DESC3] NOT LIKE '%снуюче%' AND [DESC3] NOT LIKE '%уществую%')) " &
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
            DatabaseDataAccessObject.FillMountItemData(specItemList, strSqlQuery)

            strSqlQuery = "SELECT [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [BLOCK], [CNT], [DESC3] " &
              "FROM [PNLCOMP] " &
              "WHERE ([INST] NOT LIKE '9[0-9]%' OR [INST] IS NULL) " &
                "AND ([WDBLKNAM] = 'DIN' OR [WDBLKNAM] = 'WW' OR [WDBLKNAM] = 'DI') " &
                "AND ([CNT] IS NOT NULL)  AND ([DESC3] IS NULL OR ([DESC3] NOT LIKE '%снуюче%' AND [DESC3] NOT LIKE '%уществую%'))"
            DatabaseDataAccessObject.FillSheetDinAndDuctItemData(specItemList, strSqlQuery)

            strSqlQuery = "SELECT [TERMS].[TAGSTRIP], [TERMS].[MFG], [TERMS].[CAT], [TERMS].[INST], [TERMS].[LOC], [PNLTERM].[WDBLKNAM], [TERMS].[CNT], [TERMS].[DESC3] " &
                            "FROM [TERMS] " &
                            "LEFT JOIN [PNLTERM] ON ([TERMS].[TAGSTRIP] = [PNLTERM].[TAGSTRIP] AND [TERMS].[TERM] = [PNLTERM].[TERM]) " &
                            "WHERE [TERMS].[INST] NOT LIKE '9[0-9]%' " &
                            "AND [TERMS].[PAR1_CHLD2] = '1' AND [TERMS].[LNUMBER] = '01' " &
                            "AND [TERMS].[MFG] <> 'IGNORED' " &
                            "AND ([TERMS].[DESC3] IS NULL OR ([TERMS].[DESC3] NOT LIKE '%снуюче%' AND [TERMS].[DESC3] NOT LIKE '%уществую%')) " &
                            "UNION ALL " &
                            "SELECT [A].[TAGSTRIP], [A].[MFG], [A].[CAT], [A].[INST], [A].[LOC], [A].[WDBLKNAM], [A].[CNT], [A].[DESC3] " &
                             "FROM [PNLTERM] as A " &
                             "WHERE EXISTS(SELECT DISTINCT [TAGSTRIP], [LOC] " &
                                 "FROM [TERMS] as B " &
                                 "WHERE [B].[INST] NOT LIKE '9[0-9]%'  " &
                                 "AND [B].[PAR1_CHLD2] = '1' " &
                                 "AND [B].[MFG] <> 'IGNORED' AND ([B].[DESC3] IS NULL OR ([B].[DESC3] NOT LIKE '%снуюче%' AND [B].[DESC3] NOT LIKE '%уществую%')) " &
                             "AND [A].[TAGSTRIP] = [B].[TAGSTRIP] " &
                             "AND [A].[LOC] = [B].[LOC]) " &
                             "AND [TERM] IS NULL"

            DatabaseDataAccessObject.FillMountTerminalData(specItemList, strSqlQuery)
            'If tempSpecItemList.Count = 0 Then
            '    strSqlQuery = "SELECT [TAGSTRIP], [MFG], [CAT], [INST], [LOC], [WDBLKNAM], [CNT], [DESC3] FROM [TERMS] WHERE [INST] NOT LIKE '9[0-9]%' AND (([PAR1_CHLD2] = '1' AND [LNUMBER] = '01') OR [PAR1_CHLD2] BETWEEN '301' AND '399') AND [MFG] <> 'IGNORED'  AND ([DESC3] IS NULL OR ([DESC3] NOT LIKE '*снуюче*' AND [DESC3] NOT LIKE '*уществую*'))"
            '    DatabaseDataAccessObject.FillPageTerminalData(specItemList, strSqlQuery)
            'Else
            '    DatabaseDataAccessObject.FillPageTerminalData(specItemList, strSqlQuery)
            'End If

            Return MakeListForSheet(specItemList)
        End Function

        Private Function MakeListForSheet(inputList As List(Of DrawingSpecificationItem)) As SortedList(Of String, List(Of DrawingSpecificationItem))
            Dim resultList As New SortedList(Of String, List(Of DrawingSpecificationItem))

            For Each location In inputList.GroupBy(Function(x) x.Location)
                Dim instKey As Short = location(0).Instance
                Dim locKey As String

                Dim locList As List(Of DrawingSpecificationItem) = location.ToList

                If instKey >= 800 AndAlso instKey <= 899 Then
                    locKey = "Кабельно-провідникова продукція"
                ElseIf instKey >= 700 AndAlso instKey <= 799 Then
                    locKey = "Обладнання на ПС"
                Else
                    locKey = location(0).Location
                End If

                If resultList.ContainsKey(locKey) Then
                    'Dim curList = resultList.Item(instKey)
                    'For Each element In location
                    '    If curList.Exists(Function(x) x.CatalogName.Equals(element.CatalogName)) Then
                    '        With curList.Find(Function(x) x.CatalogName.Equals(element.CatalogName))
                    '            .AddTag(element.TAG)
                    '            '.Count += element.Count - 1
                    '        End With
                    '    Else
                    '        curList.Add(element)
                    '    End If
                    'Next
                    MsgBox("что-то пошло не так")
                Else
                    resultList.Add(locKey, locList)
                End If
            Next

            Return resultList
        End Function
    End Module
End Namespace

