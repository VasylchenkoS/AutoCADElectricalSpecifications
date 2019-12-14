Imports System.Windows.Forms
Imports AutoCADElectricalSpecifications.com.vasilchenko.Classes
Imports AutoCADElectricalSpecifications.com.vasilchenko.DatabaseConnection

Namespace com.vasilchenko.Forms

    Public Class PageSelector
        Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
            Hide()
        End Sub

        Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
            Close()
        End Sub

        Private Sub PageSelector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Const strSqlQuery As String = "SELECT [DWGIDX], [DWGNAME] FROM [PDS_DWGINDEX]"
            Dim objDataTable As System.Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
            Dim acDrawingList As New SortedList(Of Integer, String)
            Dim i As Short = 0
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    If Not IsDBNull(objRow.item("DWGIDX")) AndAlso Not IsDBNull(objRow.item("DWGNAME")) Then
                        acDrawingList.Add(objRow.item("DWGIDX"), Strings.Right( objRow.item("DWGNAME"), Len(objRow.item("DWGNAME"))- InStrRev( objRow.item("DWGNAME"),"\")))
                    End If
                Next objRow
            End If

            'lvSheets.Columns.Add("Checked", Windows.HorizontalAlignment.Left)

            lvSheets.MultiSelect = True
            lvSheets.FullRowSelect = True
            lvSheets.Sorting = SortOrder.Ascending
            lvSheets.GridLines = True
            lvSheets.View = View.Details
            lvSheets.AutoSize = True
            AutoSize = true 

            lvSheets.Sorting = SortOrder.Ascending
            'lvSheets.Sort()
            
            lvSheets.Columns.Add("##",50, HorizontalAlignment.Left)
            lvSheets.Columns.Add("Full Name", 450, HorizontalAlignment.Left)
            'lvSheets.CheckBoxes = True
            Dim arrLvItem(acDrawingList.Count - 1) As Windows.Forms.ListViewItem

            For Each pair As KeyValuePair(Of Integer, String) In acDrawingList
                arrLvItem(i) = New Windows.Forms.ListViewItem(pair.key)
                arrLvItem(i).SubItems.Add(pair.Value)
                i += 1
            Next pair
            lvSheets.Items.AddRange(arrLvItem)
        End Sub
        
        Private _mSortingColumn As ColumnHeader

        Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvSheets.ColumnClick
            ' Get the new sorting column. 
            Dim newSortingColumn As ColumnHeader = lvSheets.Columns(e.Column)
            ' Figure out the new sorting order. 
            Dim sortOrder As SortOrder
            If _mSortingColumn Is Nothing Then
                ' New column. Sort ascending. 
                sortOrder = SortOrder.Ascending
            Else ' See if this is the same column. 
                If newSortingColumn.Equals(_mSortingColumn) Then
                    ' Same column. Switch the sort order. 
                    If _mSortingColumn.Text.StartsWith("> ") Then
                        sortOrder = SortOrder.Descending
                    Else
                        sortOrder = SortOrder.Ascending
                    End If
                Else
                    ' New column. Sort ascending. 
                    sortOrder = SortOrder.Ascending
                End If
                ' Remove the old sort indicator. 
                _mSortingColumn.Text = _mSortingColumn.Text.Substring(2)
            End If
            ' Display the new sort order. 
            _mSortingColumn = newSortingColumn
            If sortOrder = SortOrder.Ascending Then
                _mSortingColumn.Text = "> " & _mSortingColumn.Text
            Else
                _mSortingColumn.Text = "< " & _mSortingColumn.Text
            End If
            ' Create a comparer. 
            lvSheets.ListViewItemSorter = New ListViewColumnSorter(e.Column, sortOrder)
            ' Sort. 
            lvSheets.Sort()
        End Sub

    End Class
End Namespace