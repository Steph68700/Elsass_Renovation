Imports System.Globalization
Imports Microsoft.Office.Interop
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports OleDbConnection = System.Data.OleDb.OleDbConnection

Public Class Form1


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) _
Handles MyBase.Load
        TextBox1.Text = DateTime.Now.ToString.Substring(6, 4)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fonctionnement()
    End Sub

    Private Sub fonctionnement()
        Label3.Text = "Chargement en cours"
        ' Variable Récupération Data API
        Dim Jsonobj As Object
        Dim JsonArr As JArray

        ' Connexion a l'API
        Dim client As RestClient
        Dim response As IRestResponse
        Dim request = New RestRequest(Method.[GET])
        request.AddHeader("Authorization", "APIKey f6...06")

        ' Calendrier Francais
        Dim dfi = DateTimeFormatInfo.CurrentInfo
        Dim calendar = dfi.Calendar



        ' **********************************
        'Indiquer le nombre de Colonnes
        ' **********************************
        Dim NumSemaine As Integer = 1
        DGW.ColumnCount = 56
        DGW.Columns(0).Name = "idUser"
        DGW.Columns(1).Name = "Noms"
        DGW.Columns(2).Name = "Zone"
        For i = 3 To 55
            DGW.Columns(i).Name = NumSemaine
            NumSemaine = NumSemaine + 1
        Next



        ' **************************************
        'Récuperer tous les utilisateurs
        ' **************************************
        client = New RestClient("https://api.alobees.com/api/user?limit=500")
        client.Timeout = -1
        response = client.Execute(request)

        Jsonobj = JsonConvert.DeserializeObject(Of Object)(response.Content)
        JsonArr = Jsonobj("data")
        For i = 0 To JsonArr.Count - 1
            Dim data As JObject = JsonArr(i)

            ' Ajouter 5 zones pour chaque utilisateurs
            For U = 1 To 5
                DGW.Rows.Add(data("id"), data("lastname"), U)
            Next
        Next




        ' ********************************
        ' Recupérer toutes les données : 500/500
        '**********************************
        client = New RestClient("https://api.alobees.com/api/assignment?limit=1")
        client.Timeout = -1
        response = client.Execute(request)
        Jsonobj = JsonConvert.DeserializeObject(Of Object)(response.Content)
        Dim totalResult As Integer = Jsonobj("total")
        Dim Skip As Integer = 0

        ' Variable 2 dimension ligne, colonne
        Dim NbrJour(DGW.Rows.Count - 1, DGW.Columns.Count - 1) As Integer

        Dim LoadCount = 0
        Dim LoadCount2 = 0
        While totalResult > 0

            ' *************************************************************************
            ' Récuperation et calculs des données nbr jours travaillé/semaine/distance/personne
            ' *************************************************************************
            client = New RestClient("https://api.alobees.com/api/assignment?limit=500&skip=" + Skip.ToString)
            client.Timeout = -1
            response = client.Execute(request)

            Dim obj As Object = JsonConvert.DeserializeObject(Of Object)(response.Content)
            Dim JsonArray As JArray = obj("data")


            For i = 0 To JsonArray.Count - 1

                Dim data As JObject = JsonArray(i)
                Dim weekOfyear As String = calendar.GetWeekOfYear(data("date"), dfi.CalendarWeekRule, DayOfWeek.Thursday)

                Dim cmd As OleDbCommand
                'Dim cn As OleDbConnection
                Dim zone As String

                Dim ExcelFile As String = AppDomain.CurrentDomain.BaseDirectory + "Zone.xlsx"

                'cn = New OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelFile + "; Extended Properties = 'Excel 12.0;HDR=Yes;IMEX=1;'")
                'cmd = New OleDbCommand("select [Zone] from [Feuil1$] where [Lieux]='" + data("site")("city").ToString + "'", cn)

                Using cn As OleDbConnection = New OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelFile + "; Extended Properties = 'Excel 12.0;HDR=Yes;IMEX=1;'")
                    cn.Open()
                    cmd = New OleDbCommand("select [Zone] from [Feuil1$] where [Lieux]='" + data("site")("city").ToString + "'", cn)

                    If cmd.ExecuteScalar <> Nothing Then
                        zone = cmd.ExecuteScalar.ToString
                    End If
                    cn.Close()
                End Using


                For a = 0 To DGW.Rows.Count - 1

                    For b = 3 To DGW.Columns.Count - 1

                        If zone = DGW.Rows(a).Cells(2).Value And data("date").ToString.Substring(0, 4) = TextBox1.Text And weekOfyear = DGW.Columns(b).HeaderText.ToString And data("user")("id").ToString = DGW.Rows(a).Cells(0).Value Then
                            NbrJour(a, b) += 1
                        End If

                        DGW.Rows(a).Cells(b).Value = NbrJour(a, b)
                    Next
                Next

                LoadCount2 += 1
                ProgressBar2.Value = LoadCount2
            Next

            totalResult = totalResult - 500
            Skip = Skip + 500
            'MessageBox.Show(totalResult)
            LoadCount2 = 0
            LoadCount += 1
            ProgressBar1.Value = LoadCount

            If totalResult < 0 Then
                Label3.Text = "Chargement Terminé"
                ProgressBar1.Value = 0
                ProgressBar1.Value = 0
            End If
        End While



        DGW.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
        DGW.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
        DGW.Columns(1).DefaultCellStyle.BackColor = Color.LightCyan
        DGW.Columns(2).DefaultCellStyle.BackColor = Color.LightCyan

        'DGW.Rows(0).Frozen = True 'Block row 1
        DGW.ColumnHeadersDefaultCellStyle.BackColor = Color.LightCyan

        DGW.Columns(0).Visible = False 'Column 1 not visible 
        DGW.Columns(1).Frozen = True ' Block column 2
        DGW.Columns(2).Frozen = True ' Block Column 3
    End Sub

End Class
