Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class Form1
    'variables utilitzades per el tractament de la bd
    Private con As SqlConnection
    Private dts As DataSet
    Private Adaptador As SqlDataAdapter
    'variable que utilitzem per fer els bindings! metodes
    Private bmb As BindingManagerBase
    'event que carrega el form, l'utilitzem per establir la conexio amb la bd y la crida a la creacio dels bindigns, 
    'a mes fem crides a metodes per fer una interficie de usuari mes "maca"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'crides per metodes millores interficies
        botonsVisivilitatNoCanvi()
        textBoxDesconectats()
        'creacio de la comunicacio de la bd
        con = New SqlConnection
        con.ConnectionString = "Data Source=.\SQLEXPRESS;Initial Catalog=COMPASSTRAVEL; Trusted_Connection=True;"
        con.Open()
        'generem un adapter per fer la nostra feina
        Adaptador = New SqlDataAdapter("select * from MEMBERS", con)

        Dim cmBase As SqlCommandBuilder = New SqlCommandBuilder(Adaptador)

        dts = New DataSet
        Adaptador.Fill(dts, "Members")
        'crida a la creacio de els datebindings
        ConstruirDataBinding()

        'aqui tindriem que fer metodes de tractament de les columnes, no fet ja que no e conseguit fer-ho funcional
        'em pasat a capturar el event del click dels botons

    End Sub
    'metode que construieix els bindings que utilitzarem per fer la faina
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
        txbPass.PasswordChar = "*" 'perqe el camb ass sigui **** al escriure
        oBind = Nothing

        bmb = BindingContext(dts, "Members")

    End Sub
    '//EVENTS PER EL TRACTAMENT DE OPCIONS INICIALS\\
    'captura del event click afegir un nou usuari
    Private Sub btnAfegir_Click(sender As Object, e As EventArgs) Handles btnAfegir.Click
        'crides per generar una interficie mes entenedora
        lblPass.Text = "-----"
        lblUser.Text = "-----"
        textBoxConectats()
        botonsVisivilitatCanvi()
        'metode de els binding que genera una nova row
        bmb.AddNew()
        'petita idea per tal que la id autoincremental aparegui sense necesitat de fer un refresh de la bd
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
    'event que capturem del boto eliminar
    Private Sub btnEliminar_Click(sender As Object, e As EventArgs) Handles btnEliminar.Click
        'interficie millorada per se mes entendible
        lblPass.Text = "-----"
        lblUser.Text = "-----"
        'control de que existeix la row, no interesa fer click y que peti degut a que no troba una row
        If ControlPosicio() Then
            'demanem conformitat si es si, eliminem la row, i cridem al metode de persistenia que pasa la informacio del adapter a la bd
            If MessageBox.Show("Segur que vols eliminar al membre", "Atención", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                bmb.RemoveAt(bmb.Position)
                generarPersistencia()
                MsgBox("Eliminado")
            Else
                MsgBox("No eliminat")
            End If
        End If
    End Sub
    'event capturem el que volem modificar el usuari en que es trobem, problemes amb el nom d'usuari es troba a si mateix y no
    'deica inserirse
    Private Sub btnModificar_Click(sender As Object, e As EventArgs) Handles btnModificar.Click
        lblPass.Text = "-----"
        lblUser.Text = "-----"
        'control per modificar si existeix una row sino cascara
        If ControlPosicio() Then
            botonsVisivilitatCanvi()
            textBoxConectats()
        End If
    End Sub
    '//EVENTS QUE CAPTUREM AL AFEGIR MODIFICAR AQUEST BOTONS APAREIXEN QUAN ESTEM TREBALLAN SINO NO ES MOSTREN\\
    'boto capturem el click en afegir, conformitat que el que em fet ens interesa
    'event que crida metodes per tal de retornar banderas i decidir si afegim o no els canvis
    Private Sub btnAcCanvi_Click(sender As Object, e As EventArgs) Handles btnAcCanvi.Click
        Dim validat = True
        'si el nom user no es valid
        If controlNomUser() Then
            validat = False
        End If
        'si la contrasenya no es valida
        If ControlContrasenya() Then
            validat = False
        End If
        'si tot esta correcte fem algunes modificacion a la interfaç d'usuari y afegim els canvis a mes de generar la persistencia a bd
        If validat Then
            txbNomUser.BackColor = SystemColors.Control
            txbNom.ReadOnly = False
            txbPass.BackColor = SystemColors.Control
            txbPais.ReadOnly = False
            bmb.EndCurrentEdit()
            textBoxDesconectats()
            botonsVisivilitatNoCanvi()
            generarPersistencia()
        End If
    End Sub
    'boto per tal de que si les modificacions no es agraden es puguin cancelar, no volem afegir o altres
    Private Sub btnDenegarCanvi_Click(sender As Object, e As EventArgs) Handles btnDenegarCanvi.Click
        bmb.CancelCurrentEdit()
        textBoxDesconectats()
        botonsVisivilitatNoCanvi()
        txbNomUser.BackColor = SystemColors.Control
        txbPass.BackColor = SystemColors.Control
    End Sub
    'event capturem click boto generar, es un event per tal de generar una petita ajuda de nom de usuari fet per la maquina
    'codi cedit per en pau, que ha ajudat a la seba creacio
    Private Sub btnSugueriments_Click(sender As Object, e As EventArgs) Handles btnSugueriments.Click
        Dim nom, cognom, username As String
        Try
            'capturem noms
            nom = txbNom.Text.ToString
            cognom = txbCognom.Text.ToString
            'generem el nom d'usuari
            username = Mid(nom, 1, 3)
            username = username + Mid(cognom, cognom.Count - 2, cognom.Count)
            'el mostrem
            txbNomUser.Text = username
        Catch ex As Exception
            'si peta infromem
            MsgBox("Per poder fer el sugueriment necesitarem nom i cognom de almenys de 3 caracters")
        End Try

    End Sub

    '//EVENTS CAPTURATS PER FER EL MOVIMENT DE LES ROW ATRAVES DE BOTONS O CONTROL DEL TECLAT\\
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
    'event que captura el click agafan una row i buscan la seba id i la busca i es colocara en aquella posicio y la mostrar
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
        'sino esta la id linformem
        If trobat = False Then
            MsgBox("Aquesta id no esta en el nostre sistema")
        End If
    End Sub
    '//METODES QUE CRIDEN ELS EVENTS PER TAL DE TENIR UNA MICA MES DESGLOSAT EL CODI\\
    'metode que controla la contrasenya tingui un minim de caracters, mes de 4 , que sino es almeny de mes de 9
    'informem que la aceptem pero no es del tot segura i si son mes de 9 re
    'a mes tenim conrol de interficie per fer mes visual els casos
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
    'metode per la interficie que fa que els textbox no siguin modificables
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
    'metode complementari per que els text box es puguin modificar
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
    'aquest 4 metodes s'utlitizen per fer un control de la interfici y nomes mostrar el que es necesari en cada moment
    'i que nomes modifiqui quan nosaltres volem deixarlo

    'metode que mira que tinguem row per tal de no petar
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
    '!!!!!!!!!!!!!!!!!!!!!! METODE NO FUNCIONAL AL 100% te problemes quan modifiquem
    'metode que controla que el nom de usuari no estiqui repetit en cas que modifiquem o afegim el usuari
    'al modificar no funciona be
    'tambe tenim control sobre part de la interficie per fer mes entenedor el que es la comunicacio
    Private Function controlNomUser()
        Dim nomUserCapturat = txbNomUser.Text.ToString
        MsgBox(nomUserCapturat)
        Dim NomUserComprovar
        Dim bandera = False
        If txbNomUser.Text.Trim = String.Empty Then
            bandera = True
            lblUser.Text = "El nom del user no pot quedar en blanc"
            txbNomUser.BackColor = Color.Red
        Else
            For index As Integer = 0 To dts.Tables("Members").Rows.Count - 1
                'idea para el modificar es no cojer la row que estoy pero casca al añadir
                NomUserComprovar = dts.Tables("Members").Rows(index)("USERNAME").ToString()
                MsgBox(NomUserComprovar)
                If txbNomUser.Text.Equals(NomUserComprovar) Then
                    lblUser.Text = "Nom usuari ja en us"
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
    'metode que genera la persistencia, pasa la info del adapter a la bd
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
End Class
