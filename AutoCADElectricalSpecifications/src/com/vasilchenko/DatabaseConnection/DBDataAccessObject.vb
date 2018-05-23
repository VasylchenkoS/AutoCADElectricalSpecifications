Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.Modules
Imports Autodesk.AutoCAD.DatabaseServices

Namespace com.vasilchenko.DatabaseConnection
    Module DBDataAccessObject
        Public Function FillSheetItemData(acDatabase As Database, inputList As List(Of DrawingSpecificationItem)) As List(Of DrawingSpecificationItem)
            Dim strSQLQuery As String = "SELECT [TAGNAME], [FAMILY], [MFG], [CAT], [INST], [LOC] " &
                            "FROM [COMP] " &
                            "WHERE [TAGNAME] IN (SELECT [TAGNAME] FROM [COMP] WHERE [DWGIX] = " &
                            "(SELECT [DWGIDX] FROM [PDS_DWGINDEX] WHERE [DWGNAME] = '" & acDatabase.OriginalFileName & "')) " &
                            "AND [PAR1_CHLD2] = '1' AND [LOC] NOT LIKE '%IGNOR%'"

            Dim objDataTable As Data.DataTable = DBConnections.GetOleDataTable(strSQLQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        If Not IsDBNull(.item("TAGNAME")) Then currentItem.AddTag(.item("TAGNAME"))
                        If IsDBNull(.item("FAMILY")) Then currentItem.Family = "PLCIO" Else currentItem.Family = .item("FAMILY")
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetNumericFromString(.item("INST"))
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC")
                    End With
                    FillItemDescription(currentItem)
                    If inputList.Exists(Function(x) x.CatalogName.Equals(currentItem.CatalogName) And x.Location.Equals(currentItem.Location)) Then
                        inputList.Find(Function(x) x.CatalogName.Equals(currentItem.CatalogName) And x.Location.Equals(currentItem.Location)).AddTag(currentItem.TAG)
                    Else
                        inputList.Add(currentItem)
                    End If
                Next objRow
            End If

            Return inputList
        End Function
        Public Function FillSheetTerminalData(acDatabase As Database, inputList As List(Of DrawingSpecificationItem)) As List(Of DrawingSpecificationItem)
            Dim strSQLQuery As String = "SELECT [TAGSTRIP], [MFG], [CAT], [INST], [LOC] " &
                            "FROM [TERMS] " &
                            "WHERE [DWGIX] = " &
                            "(SELECT [DWGIDX] FROM [PDS_DWGINDEX] WHERE [DWGNAME] = '" & acDatabase.OriginalFileName & "') " &
                            "AND [PAR1_CHLD2] = '1'"

            Dim objDataTable As Data.DataTable = DBConnections.GetOleDataTable(strSQLQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim currentItem As New DrawingSpecificationItem
                    With objRow
                        If Not IsDBNull(.item("TAGSTRIP")) Then currentItem.AddTag(.item("TAGSTRIP"))
                        currentItem.Family = "TRMS"
                        If Not IsDBNull(.item("MFG")) Then currentItem.Manufacture = .item("MFG")
                        If Not IsDBNull(.item("CAT")) Then currentItem.CatalogName = .item("CAT")
                        If Not IsDBNull(.item("INST")) Then currentItem.Instance = AdditionalFunctions.GetNumericFromString(.item("INST"))
                        If Not IsDBNull(.item("LOC")) Then currentItem.Location = .item("LOC")
                    End With
                    FillItemDescription(currentItem)
                    If inputList.Exists(Function(x) x.CatalogName.Equals(currentItem.CatalogName) And x.Location.Equals(currentItem.Location)) Then
                        inputList.Find(Function(x) x.CatalogName.Equals(currentItem.CatalogName) And x.Location.Equals(currentItem.Location)).AddTag(currentItem.TAG)
                    Else
                        inputList.Add(currentItem)
                    End If
                Next objRow
            End If

            Return inputList
        End Function
        Private Sub FillItemDescription(ByRef objInput As DrawingSpecificationItem)
            Dim objDataTable As Data.DataTable
            Dim strSQLQuery As String

            If objInput.Family = "" Then
                objInput.Family = "PLCIO"
            ElseIf objInput.Family = "CBL" Then
                objInput.Family = "W0"
            End If

            strSQLQuery = "SELECT [DESCRIPTION], [USER1], [USER2] " &
                    "FROM [" & objInput.Family & "] " &
                    "WHERE [CATALOG] = '" & objInput.CatalogName & "'"
            objDataTable = DBConnections.GetSQLDataTable(strSQLQuery, My.Settings.default_catSQLConnectionString)
            If Not IsNothing(objDataTable) Then
                objInput.Description = objDataTable.Rows(0).Item("DESCRIPTION")
                If Not IsDBNull(objDataTable.Rows(0).Item("USER1")) Then objInput.Article = objDataTable.Rows(0).Item("USER1")
                If Not IsDBNull(objDataTable.Rows(0).Item("USER2")) Then objInput.Unit = objDataTable.Rows(0).Item("USER2")
            End If
        End Sub

    End Module
End Namespace
