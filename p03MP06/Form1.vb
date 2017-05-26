Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class Form1
    Private con As SqlConnection
    Private dts As DataSet
    Private Adaptador As SqlDataAdapter

    Private bmb As BindingManagerBase



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        botonsVisivilitatNoCanvi()
        textBoxDesconectats()

        con = New SqlConnection
        con.ConnectionString = "Data Source=.\SQLEXPRESS;Initial Catalog=COMPASSTRAVEL; Trusted_Connection=True;"
        con.Open()

        Adaptador = New SqlDataAdapter("select * from MEMBERS", con)

        Dim cmBase As SqlCommandBuilder = New SqlCommandBuilder(Adaptador)

        dts = New DataSet
        Adaptador.Fill(dts, "Members")

        ConstruirDataBinding()

    End Sub


    Private Sub ConstruirDataBinding()
        Dim oBind As Binding

        oBind = New Binding("Text", dts, "Members.MEMBERID")
        txbID.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.FIRSTNAME")
        txbNom.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.LASTNAME")
        txbCognom.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.ADDRESS1")
        txbAdr1.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.ADDRESS2")
        txbAdr2.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.STATE")
        txbEstat.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.POSTALCODE")
        txbCP.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.COUNTRY")
        txbPais.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.USERNAME")
        txbNomUser.DataBindings.Add(oBind)
        oBind = Nothing

        oBind = New Binding("Text", dts, "Members.PASSWORD")
        txbPass.DataBindings.Add(oBind)
        txbPass.PasswordChar = "*"
        oBind = Nothing

        bmb = BindingContext(dts, "Members")

    End Sub

    Private Sub btnAfegir_Click(sender As Object, e As EventArgs) Handles btnAfegir.Click
        textBoxConectats()
        botonsVisivilitatCanvi()
        bmb.AddNew()
        Dim idPertocada = 0
        Dim idMesGran = 0
        For index As Integer = 0 To dts.Tables("Members").Rows.Count - 1
            idMesGran = dts.Tables("Members").Rows(index)("MEMBERID").ToString()
            If idPertocada < idMesGran Then
                idPertocada = idMesGran
            End If
        Next
        txbID.Text = idPertocada + 1

    End Sub

    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        If controlPosicio() Then
            If MessageBox.Show("Segur que vols eliminar al membre", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                bmb.RemoveAt(bmb.Position)
                generarPersistencia()
                MsgBox("Eliminado")
            Else
                MsgBox("No eliminat")
            End If
        End If
    End Sub

    Private Sub btnPrimer_Click(sender As Object, e As EventArgs) Handles btnPrimer.Click
        bmb.Position = 0
    End Sub

    Private Sub btnAtras_Click(sender As Object, e As EventArgs) Handles btnAtras.Click
        bmb.Position = bmb.Position - 1
    End Sub

    Private Sub btnAlante_Click(sender As Object, e As EventArgs) Handles btnAlante.Click
        bmb.Position = bmb.Position + 1
    End Sub

    Private Sub btnUltim_Click(sender As Object, e As EventArgs) Handles btnUltim.Click
        bmb.Position = bmb.Count - 1
    End Sub

    Private Function ControlContrasenya()
        Dim bandera = False
        If txbPass.Text.ToString.Count < 4 Then
            lblPass.Text = "Aquesta contrasenya es masa curta"
            txbPass.BackColor = Color.Red
            bandera = True
        Else
            If txbPass.Text.ToString.Count > 8 Then
                txbPass.BackColor = Color.White
                lblPass.Text = "Password aceptada"
            Else
                lblPass.Text = "Contrasenya aceptada, " + vbCr + "pero s'aconsella de 9 diguits o mes"
                txbPass.BackColor = Color.Orange
            End If
        End If
        Return bandera
    End Function

    'Private Sub txbPass_Leave(sender As Object, e As EventArgs) Handles txbPass.Leave
    '    If txbPass.ReadOnly = False Then
    '        ControlContrasenya()
    '    End If

    'End Sub

    Private Sub btnModificar_Click(sender As Object, e As EventArgs) Handles btnModificar.Click
        If ControlPosicio() Then
            botonsVisivilitatCanvi()
            textBoxConectats()
        End If

    End Sub

    Private Sub btnAcCanvi_Click(sender As Object, e As EventArgs) Handles btnAcCanvi.Click
        Dim validat = True

        If controlNomUser() Then
            validat = False
        End If
        If ControlContrasenya() Then
            validat = False
        End If

        If validat Then
            lblPass.Text = "-----"
            txbNomUser.BackColor = SystemColors.Control
            txbNom.ReadOnly = False
            lblUser.Text = "-----"
            txbPass.BackColor = SystemColors.Control
            txbPais.ReadOnly = False
            bmb.EndCurrentEdit()
            textBoxDesconectats()
            botonsVisivilitatNoCanvi()
            generarPersistencia()
        End If

        
    End Sub

    Private Sub btnDenegarCanvi_Click(sender As Object, e As EventArgs) Handles btnDenegarCanvi.Click
        bmb.CancelCurrentEdit()
        textBoxDesconectats()
        botonsVisivilitatNoCanvi()
        txbNomUser.BackColor = SystemColors.Control
        txbPass.BackColor = SystemColors.Control
    End Sub
    Public Sub textBoxDesconectats()
        'txbID.Enable = False
        'txbNom.Enable = False
        'txbCognom.Enable = False
        'txbAdr1.Enable = False
        'txbAdr2.Enable = False
        'txbEstat.Enable = False
        'txbCP.Enable = False
        'txbPais.Enable = False
        'txbNomUser.Enable = False
        'txbPass.Enable = False
        txbID.ReadOnly = True
        txbNom.ReadOnly = True
        txbCognom.ReadOnly = True
        txbAdr1.ReadOnly = True
        txbAdr2.ReadOnly = True
        txbEstat.ReadOnly = True
        txbCP.ReadOnly = True
        txbPais.ReadOnly = True
        txbNomUser.ReadOnly = True
        txbPass.ReadOnly = True
    End Sub
    Public Sub textBoxConectats()
        'txbNom.Enable = True
        'txbCognom.Enable = True
        'txbAdr1.Enable = True
        'txbAdr2.Enable = True
        'txbEstat.Enable = True
        'txbCP.Enable = True
        'txbPais.Enable = True
        'txbNomUser.Enable = True
        'txbPass.Enable = True

        txbNom.ReadOnly = False
        txbCognom.ReadOnly = False
        txbAdr1.ReadOnly = False
        txbAdr2.ReadOnly = False
        txbEstat.ReadOnly = False
        txbCP.ReadOnly = False
        txbPais.ReadOnly = False
        txbNomUser.ReadOnly = False
        txbPass.ReadOnly = False
    End Sub
    Private Sub botonsVisivilitatNoCanvi()


        btnAcCanvi.Visible = False
        btnDenegarCanvi.Visible = False
        btnSugueriments.Visible = False

        btnAfegir.Visible = True
        btnEliminar.Visible = True
        btnModificar.Visible = True

        btnAlante.Visible = True
        btnAtras.Visible = True
        btnPrimer.Visible = True
        btnUltim.Visible = True

        GBBuscarID.Visible = True

    End Sub
    Private Sub botonsVisivilitatCanvi()

        btnAcCanvi.Visible = True
        btnDenegarCanvi.Visible = True
        btnSugueriments.Visible = True

        btnAfegir.Visible = False
        btnEliminar.Visible = False
        btnModificar.Visible = False

        btnAlante.Visible = False
        btnAtras.Visible = False
        btnPrimer.Visible = False
        btnUltim.Visible = False

        GBBuscarID.Visible = False

    End Sub
    Public Function ControlPosicio()
        Dim bandera = True
        If bmb.Position = -1 Then
            bandera = False
        End If
        Return bandera

    End Function

    'Private Sub txbNomUser_Leave(sender As Object, e As EventArgs) Handles txbNomUser.Leave
    '    If txbNomUser.ReadOnly = False Then
    '        controlNomUser()
    '    End If

    'End Sub

    Private Function controlNomUser()
        Dim nomUserCapturat = txbNomUser.Text.ToString
        Dim NomUserComprovar
        Dim bandera = False
        If txbNomUser.Text.Trim = String.Empty Then
            bandera = True
            lblUser.Text = "El nom del user no pot quedar en blanc"
            txbNomUser.BackColor = Color.Red
        Else
            For index As Integer = 0 To dts.Tables("Members").Rows.Count - 1
                NomUserComprovar = dts.Tables("Members").Rows(index)("USERNAME").ToString()
                If txbNomUser.Text.Equals(NomUserComprovar) And nomUserCapturat <> NomUserComprovar Then
                    lblUser.Text = "Aquest nom no es pot utilitzar ja el tenim en us"
                    txbNomUser.Text = ""
                    txbNomUser.BackColor = Color.Red
                    bandera = True
                End If
            Next
        End If
        If bandera = False Then
            lblUser.Text = "Nom d'usuari correcte"
        End If
        Return bandera
    End Function

    Private Sub generarPersistencia()
        Dim dt As DataTable
        Try
            dt = dts.Tables("Members").GetChanges()
            Adaptador.Update(dt)
            dts.Tables("Members").AcceptChanges()
        Catch ex As Exception
            MsgBox("La persistencia a patit un error")
        End Try
    End Sub

    Private Sub btnSugueriments_Click(sender As Object, e As EventArgs) Handles btnSugueriments.Click
        Dim nom, cognom, username As String
        Try
            nom = txbNom.Text.ToString
            cognom = txbCognom.Text.ToString

            username = Mid(nom, 1, 3)
            username = username + Mid(cognom, cognom.Count - 2, cognom.Count)

            txbNomUser.Text = username
        Catch ex As Exception
            MsgBox("Per poder fer el sugueriment necesitarem nom i cognom de mes de 3 caracters")
        End Try
        
    End Sub
    'Events per tal de poder mourens per el nostres ids fen click alt am el teclat mes les fletxes
    'CODIGO COJIDO DE INTERNET
    Private Sub txbID_KeyDown(sender As Object, e As KeyEventArgs) Handles txbID.KeyDown

        If e.Alt And e.KeyCode = Keys.Up Then
            bmb.Position = 0
        End If

        If e.Alt And e.KeyCode = Keys.Left Then
            bmb.Position = bmb.Position - 1
        End If

        If e.Alt And e.KeyCode = Keys.Right Then
            bmb.Position = bmb.Position + 1
        End If

        If e.Alt And e.KeyCode = Keys.Down Then
            bmb.Position = bmb.Count - 1
        End If
    End Sub
    
    Private Sub btnBuscarID_Click(sender As Object, e As EventArgs) Handles btnBuscarID.Click
        Dim registreBuscat = 0
        Dim idContacte
        Dim idTractament
        Dim trobat = False
        'control si no esta buid
        If txbCercaID.Text.Trim = String.Empty Then
            MsgBox("EL camp id no pot estar buid per buscar per id")
        Else
            idContacte = txbCercaID.Text 'agafem la info del textbox
            'si estem ven posicionat farem un bucle buscan la id en la taula en cas de no trobar ens quedaremm alla
            'si el trobem pasarem agafarem el registre numero de voltes que em fet i ligualmenti cridarem al metode
            'que en mostre la informacio
            For index As Integer = 0 To dts.Tables("Members").Rows.Count - 1
                idTractament = dts.Tables("Members").Rows(index)("MEMBERID").ToString()
                If idContacte.Equals(idTractament) Then
                    trobat = True
                    bmb.Position = index
                End If
            Next
        End If

        
        If trobat = False Then
            MsgBox("Aquesta id no esta en el nostre sistema")
        End If
    End Sub
End Class
