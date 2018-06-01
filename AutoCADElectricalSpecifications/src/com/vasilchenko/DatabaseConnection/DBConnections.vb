Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.Electrical.Project


Namespace com.vasilchenko.DatabaseConnection
    Module DBConnections
        Friend Function GetSQLDataTable(strSQLQuery As String, strSQLConnectionParameters As String) As Data.DataTable
            Dim objDataTable As New Data.DataTable
            Using objSQLDbConnection As New SqlConnection(strSQLConnectionParameters)
                Try
                    objSQLDbConnection.Open()
                    Dim objSQLDbCommand = New SqlCommand(strSQLQuery, objSQLDbConnection)
                    objDataTable.Load(objSQLDbCommand.ExecuteReader)


                    Return objDataTable
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Throw New ArgumentNullException
                End Try
            End Using
        End Function
        Friend Function GetOleDataTable(strSQLQuery As String) As Data.DataTable
            Dim strConstProjectDatabasePath As String = ProjectManager.GetInstance().GetActiveProject().GetDbFullPath()

            If Not IO.File.Exists(strConstProjectDatabasePath) Then
                MsgBox("Source file not found. File way: " & strConstProjectDatabasePath & " Please open some project file and repeat.", vbCritical, "File Not Found")
                strConstProjectDatabasePath = ""
                Throw New ArgumentNullException
            End If

            Dim strSQLConnectionParameters As String
            Dim objOleDbCommand As OleDbCommand
            Dim objDataTable As New Data.DataTable

            strSQLConnectionParameters = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strConstProjectDatabasePath & ";"
            Using objOleDbConnection As New OleDbConnection(strSQLConnectionParameters)
                Try
                    objOleDbConnection.Open()
                    objOleDbCommand = New OleDbCommand(strSQLQuery, objOleDbConnection)
                    objDataTable.Load(objOleDbCommand.ExecuteReader)

                    Return objDataTable
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Throw New ArgumentNullException
                End Try
            End Using

        End Function

    End Module

End Namespace
