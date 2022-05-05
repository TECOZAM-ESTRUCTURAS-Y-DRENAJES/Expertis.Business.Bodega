Imports Solmicro.Expertis.Business.Financiero

Public Class BdgCartillista

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCartillista"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Talon") = 1
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidateDuplicateKey)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_C.IDCartillista)) = 0 Then ApplicationService.GenerateError("El Cartillista es obligatorio.")
        If Length(data(_C.DescCartillista)) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
        If Length(data(_C.IDProveedor)) = 0 Then ApplicationService.GenerateError("El Proveedor es obligatorio.")
        If (New BdgParametro().GestionExplotadorObligatoria) Then
            If Length(data(_C.IDExplotadorFincas)) = 0 Then ApplicationService.GenerateError("El Explotador es obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal Data As DataRow, ByVal services As ServiceProvider)
        If Data.RowState = DataRowState.Added Then
            Dim dtAux As DataTable = New BdgCartillista().SelOnPrimaryKey(Data(_C.IDCartillista))
            If dtAux.Rows.Count > 0 Then ApplicationService.GenerateError("No se permite insertar una clave duplicada en la tabla |", cnEntidad)
        End If
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoCartillista)(AddressOf ProcesoCartillista.CrearDocumento)
        updateProcess.AddTask(Of DocumentoCartillista)(AddressOf ProcesoCartillista.ActualizarProveedor)
        updateProcess.AddTask(Of DocumentoCartillista)(AddressOf ProcesoCartillista.AsignarDatosFincas)
        updateProcess.AddTask(Of DocumentoCartillista)(AddressOf ProcesoCartillista.AsignarMaximosRdtoMunicipios)
        updateProcess.AddTask(Of DocumentoCartillista)(AddressOf ProcesoCartillista.AsignarVendimiaVariedad)
        updateProcess.AddTask(Of DocumentoCartillista)(AddressOf ProcesoCartillista.ActualizarCupos)
        updateProcess.AddTask(Of DocumentoCartillista)(AddressOf Comunes.UpdateDocument)
    End Sub

    Public Overloads Function GetItemRow(ByVal IDCartillista As String) As DataRow
        Dim dt As DataTable = New BdgCartillista().SelOnPrimaryKey(IDCartillista)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el cartillista |", IDCartillista)
        Else : Return dt.Rows(0)
        End If
    End Function

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StRecalcularVendimia
        Public IDCartillista As String
        Public IDVendimia As Integer
        Public CalcularPor As BdgCupoCartillaPor

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCartillista As String, ByVal IDVendimia As Integer, ByVal CalcularPor As BdgCupoCartillaPor)
            Me.IDCartillista = IDCartillista
            Me.IDVendimia = IDVendimia
            Me.CalcularPor = CalcularPor
        End Sub
    End Class

    <Task()> Public Shared Sub RecalcularVendimia(ByVal data As StRecalcularVendimia, ByVal services As ServiceProvider)
        Dim dblMaxT, dblMaxB, dblHaT, dblHaB, dblRdtoxHaT, dblRdtoxHaB As Double
        Dim udtRcs As New UpdatePackage

        If Length(data.IDCartillista) > 0 And data.IDVendimia <> 0 Then
            Dim FilMunicipio As New Filter
            FilMunicipio.Add("IDCartillista", FilterOperator.Equal, data.IDCartillista)
            FilMunicipio.Add("Vendimia", FilterOperator.Equal, data.IDVendimia)
            Dim dtCM As DataTable = New BdgCartillistaMunicipio().Filter(FilMunicipio)

            If data.CalcularPor = BdgCupoCartillaPor.Finca Then
                For Each rw As DataRow In dtCM.Rows
                    rw.Delete()
                Next

                Dim rwVendimia As DataRow = New BdgVendimia().GetItemRow(data.IDVendimia)

                dblRdtoxHaT = rwVendimia("RdtoxHaT")
                dblRdtoxHaB = rwVendimia("RdtoxHaB")

                Dim oFltr As New Filter
                oFltr.Add(New StringFilterItem("IDCartillista", data.IDCartillista))
                oFltr.Add(New NumberFilterItem("Vendimia", data.IDVendimia))
                Dim dtCF As DataTable = New BE.DataEngine().Filter("FrmBdgCartillistaFinca", oFltr, "IDFinca,Superficie,Rdto,Maximo,IDMunicipio,TipoVariedad")
                For Each rwCF As DataRow In dtCF.Select(Nothing, "IDMunicipio")
                    Dim rws() As DataRow = dtCM.Select(AdoFilterComposer.ComposeStringFilter("IDMunicipio", FilterOperator.Equal, rwCF("IDMunicipio")))
                    Dim rwCM As DataRow
                    If rws.Length = 0 Then
                        rwCM = dtCM.NewRow
                        dtCM.Rows.Add(rwCM)
                        rwCM("IDCartillista") = data.IDCartillista
                        rwCM("Vendimia") = data.IDVendimia
                        rwCM("IDMunicipio") = rwCF("IDMunicipio")
                        rwCM("RdtoT") = 0
                        rwCM("RdtoB") = 0
                        rwCM("HaT") = 0
                        rwCM("MaxT") = 0
                        rwCM("HaB") = 0
                        rwCM("MaxB") = 0
                    Else
                        rwCM = rws(0)
                    End If
                    Select Case rwCF("TipoVariedad")
                        Case BdgTipoVariedad.Tinta
                            rwCM("HaT") = rwCM("HaT") + rwCF("Superficie")
                            rwCM("MaxT") = rwCM("MaxT") + rwCF("Maximo")
                        Case BdgTipoVariedad.Blanca
                            rwCM("HaB") = rwCM("HaB") + rwCF("Superficie")
                            rwCM("MaxB") = rwCM("MaxB") + rwCF("Maximo")
                    End Select

                    If rwCM("HaT") > 0 And dblRdtoxHaT > 0 Then
                        rwCM("RdtoT") = rwCM("MaxT") * 100 / (rwCM("HaT") * dblRdtoxHaT)
                    End If
                    If rwCM("HaB") > 0 And dblRdtoxHaB > 0 Then
                        rwCM("RdtoB") = rwCM("MaxB") * 100 / (rwCM("HaB") * dblRdtoxHaB)
                    End If
                Next
                udtRcs.Add(dtCM)
            End If

            For Each rwCM As DataRow In dtCM.Select
                dblHaT = dblHaT + rwCM("HaT")
                dblHaB = dblHaB + rwCM("HaB")
                dblMaxT = dblMaxT + rwCM("MaxT")
                dblMaxB = dblMaxB + rwCM("MaxB")
            Next

            Dim rwCta As DataRow = New BdgCartillista().GetItemRow(data.IDCartillista)
            Dim dtCV As DataTable = New BdgCartillistaVendimia().SelOnPrimaryKey(data.IDCartillista, data.IDVendimia)
            Dim rwCV As DataRow
            If dtCV.Rows.Count = 0 Then
                rwCV = dtCV.NewRow
                dtCV.Rows.Add(rwCV)
                rwCV("IDCartillista") = data.IDCartillista
                rwCV("Vendimia") = data.IDVendimia
            Else : rwCV = dtCV.Rows(0)
            End If
            rwCV("HaT") = dblHaT
            rwCV("HaB") = dblHaB
            rwCV("MaxT") = dblMaxT
            rwCV("MaxB") = dblMaxB
            udtRcs.Add(dtCV)
            BusinessHelper.UpdatePackage(udtRcs)
        End If

    End Sub

    <Serializable()> _
    Public Class StIncrementarVendimiaTx
        Public IDCartillista As String
        Public Vendimia As Integer
        Public IncHaT As Double
        Public IncHaB As Double
        Public IncMaxT As Double
        Public IncMaxB As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCartillista As String, ByVal Vendimia As Integer, ByVal IncHaT As Double, ByVal IncHaB As Double, ByVal IncMaxT As Double, ByVal IncMaxB As Double)
            Me.IDCartillista = IDCartillista
            Me.Vendimia = Vendimia
            Me.IncHaT = IncHaT
            Me.IncHaB = IncHaB
            Me.IncMaxT = IncMaxT
            Me.IncMaxB = IncMaxB
        End Sub
    End Class

    <Task()> Public Shared Function IncrementarVendimiaTx(ByVal data As StIncrementarVendimiaTx, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDCartillista) > 0 And data.Vendimia <> 0 Then
            Dim dtCV As DataTable = New BdgCartillistaVendimia().SelOnPrimaryKey(data.IDCartillista, data.Vendimia)

            Dim rwCV As DataRow
            If dtCV.Rows.Count = 0 Then
                rwCV = dtCV.NewRow
                dtCV.Rows.Add(rwCV)
                rwCV("IDCartillista") = data.IDCartillista
                rwCV("Vendimia") = data.Vendimia
                rwCV("HaT") = 0
                rwCV("HaB") = 0
                rwCV("MaxT") = 0
                rwCV("MaxB") = 0
            Else : rwCV = dtCV.Rows(0)
            End If

            rwCV("HaT") = rwCV("HaT") + data.IncHaT
            rwCV("HaB") = rwCV("HaB") + data.IncHaB
            rwCV("MaxT") = rwCV("MaxT") + data.IncMaxT
            rwCV("MaxB") = rwCV("MaxB") + data.IncMaxB

            Return dtCV
        End If
    End Function

    <Serializable()> _
    Public Class StPrepararMunicipio
        Public IDCartillista As String
        Public Vendimia As Integer
        Public IDMunicipio As String
        Public IncHaT As Double
        Public IncHaB As Double
        Public IncMaxT As Double
        Public IncMaxB As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCartillista As String, ByVal Vendimia As Integer, ByVal IDMunicipio As String, ByVal IncHaT As Double, ByVal IncHaB As Double, ByVal IncMaxT As Double, ByVal IncMaxB As Double)
            Me.IDCartillista = IDCartillista
            Me.Vendimia = Vendimia
            Me.IDMunicipio = IDMunicipio
            Me.IncHaT = IncHaT
            Me.IncHaB = IncHaB
            Me.IncMaxT = IncMaxT
            Me.IncMaxB = IncMaxB
        End Sub
    End Class

    <Task()> Public Shared Function PrepararMunicipio(ByVal data As StPrepararMunicipio, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDCartillista) > 0 And Length(data.IDMunicipio) > 0 And data.Vendimia <> 0 Then
            Dim rwVendimia As DataRow = New BdgVendimia().GetItemRow(data.Vendimia)
            Dim dblRdtoxHaT As Double = rwVendimia("RdtoxHaT")
            Dim dblRdtoxHaB As Double = rwVendimia("RdtoxHaB")

            Dim dtCM As DataTable = New BdgCartillistaMunicipio().SelOnPrimaryKey(data.IDCartillista, data.Vendimia, data.IDMunicipio)
            Dim rwCM As DataRow
            If dtCM.Rows.Count = 0 Then
                rwCM = dtCM.NewRow
                dtCM.Rows.Add(rwCM)
                rwCM("IDCartillista") = data.IDCartillista
                rwCM("Vendimia") = data.Vendimia
                rwCM("IDMunicipio") = data.IDMunicipio
                rwCM("RdtoT") = 0
                rwCM("RdtoB") = 0
            Else
                rwCM = dtCM.Rows(0)
            End If
            rwCM("HaT") = 0
            rwCM("MaxT") = 0
            rwCM("HaB") = 0
            rwCM("MaxB") = 0

            Dim oFltr As New Filter
            oFltr.Add("IDCartillista", FilterOperator.Equal, data.IDCartillista)
            oFltr.Add("IDMunicipio", FilterOperator.Equal, data.IDMunicipio)
            oFltr.Add("Vendimia", FilterOperator.Equal, data.Vendimia)
            Dim DtFincas As DataTable = New BE.DataEngine().Filter("FrmBdgCartillistaFinca", oFltr, "IDFinca,Superficie,Rdto,Maximo,TipoVariedad")
            For Each rwFinca As DataRow In DtFincas.Select
                Select Case rwFinca("TipoVariedad")
                    Case BdgTipoVariedad.Tinta
                        rwCM("HaT") = rwCM("HaT") + rwFinca("Superficie")
                        rwCM("MaxT") = rwCM("MaxT") + rwFinca("Maximo")
                    Case BdgTipoVariedad.Blanca
                        rwCM("HaB") = rwCM("HaB") + rwFinca("Superficie")
                        rwCM("MaxB") = rwCM("MaxB") + rwFinca("Maximo")
                End Select
            Next
            rwCM("HaT") = rwCM("HaT") + data.IncHaT
            rwCM("MaxT") = rwCM("MaxT") + data.IncMaxT
            rwCM("HaB") = rwCM("HaB") + data.IncHaB
            rwCM("MaxB") = rwCM("MaxB") + data.IncMaxB

            If rwCM("HaT") > 0 And dblRdtoxHaT > 0 Then
                rwCM("RdtoT") = rwCM("MaxT") * 100 / (rwCM("HaT") * dblRdtoxHaT)
            End If
            If rwCM("HaB") > 0 And dblRdtoxHaB > 0 Then
                rwCM("RdtoB") = rwCM("MaxB") * 100 / (rwCM("HaB") * dblRdtoxHaB)
            End If
            Return dtCM
        End If
    End Function

    <Serializable()> _
    Public Class StMunicipiosSobrantes
        Public IDCartillista As String
        Public Vendimia As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCartillista As String, ByVal Vendimia As Integer)
            Me.IDCartillista = IDCartillista
            Me.Vendimia = Vendimia
        End Sub
    End Class

    <Task()> Public Shared Function MunicipiosSobrantes(ByVal data As StMunicipiosSobrantes, ByVal services As ServiceProvider) As DataTable
        Dim StrIN() As String
        Dim DtFincas, DtMunicipios As DataTable

        If Length(data.IDCartillista) > 0 And data.Vendimia <> 0 Then
            DtFincas = New BdgCartillistaFinca().Filter("DISTINCT IDFinca", "IDCartillista='" & data.IDCartillista & "' AND Vendimia=" & data.Vendimia)
            If Not DtFincas Is Nothing AndAlso DtFincas.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each oRw As DataRow In DtFincas.Select
                    ReDim Preserve StrIN(i)
                    StrIN(i) = CType(oRw("IDFinca"), Guid).ToString
                    i += 1
                Next

                If StrIN.Length > 0 Then
                    DtFincas = New BdgFinca().Filter(New InListFilterItem("IDFinca", StrIN, FilterType.String))

                    If Not DtFincas Is Nothing Then
                        If DtFincas.Rows.Count > 0 Then
                            i = 0
                            ReDim StrIN(i)
                            For Each oRw As DataRow In DtFincas.Select
                                ReDim Preserve StrIN(i)
                                StrIN(i) = Quoted(oRw("IDMunicipio"))
                            Next
                            If StrIN.Length > 0 Then
                                Dim FilMunicipio As New Filter
                                FilMunicipio.Add("IDCartillista", FilterOperator.Equal, data.IDCartillista)
                                FilMunicipio.Add("Vendimia", FilterOperator.Equal, data.Vendimia)
                                FilMunicipio.Add(New InListFilterItem("IDMunicipio", StrIN, FilterType.String, False))
                                DtMunicipios = New BdgCartillistaMunicipio().Filter(FilMunicipio)
                                If Not DtMunicipios Is Nothing Then
                                    For Each oRw As DataRow In DtMunicipios.Select
                                        oRw.Delete()
                                    Next
                                    Return DtMunicipios
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            DtMunicipios = New BdgCartillistaMunicipio().Filter(, "IDCartillista='" & data.IDCartillista & "' AND Vendimia=" & data.Vendimia)
            If Not DtMunicipios Is Nothing Then
                For Each oRw As DataRow In DtMunicipios.Select
                    oRw.Delete()
                Next
                Return DtMunicipios
            End If
        End If
    End Function

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDCartillista As String, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgCartillista().SelOnPrimaryKey(IDCartillista)
        If DtAux Is Nothing OrElse DtAux.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Cartillista: | ya existe en la tabla.", IDCartillista)
        End If
    End Sub

    <Serializable()> _
    Public Class stCrearProveedor
        Public IDProveedor As String
        Public IDContador As String
        Public DescProveedor As String
        Public Cif As String
        Public IDGrupoProveedor As String
        Public IDCNAE As String
        Public FechaAlta As Date
        Public IDCentroGestion As String
        Public RazonSocial As String
        Public Direccion As String
        Public CodPostal As String
        Public Poblacion As String
        Public Provincia As String
        Public IDPais As String
        Public IdZona As String
        Public Telefono1 As String
        Public Email As String
        Public IDMercado As String

        Public guardarProveedor As Boolean = False

        Public Sub New(ByVal IDProveedor As String, ByVal IDContador As String, ByVal DescProveedor As String, ByVal Cif As String, ByVal IDGrupoProveedor As String, ByVal IDCNAE As String, _
                       ByVal FechaAlta As Date, ByVal IDCentroGestion As String, ByVal RazonSocial As String, ByVal Direccion As String, ByVal CodPostal As String, _
                       ByVal Poblacion As String, ByVal Provincia As String, ByVal IDPais As String, ByVal IdZona As String, ByVal Telefono1 As String, _
                       ByVal Email As String, Optional ByVal guardarProveedor As Boolean = False)
            Me.IDProveedor = IDProveedor
            Me.IDContador = IDContador
            Me.DescProveedor = DescProveedor
            Me.Cif = Cif
            Me.IDGrupoProveedor = IDGrupoProveedor
            Me.IDCNAE = IDCNAE
            Me.FechaAlta = FechaAlta
            Me.IDCentroGestion = IDCentroGestion
            Me.RazonSocial = RazonSocial
            Me.Direccion = Direccion
            Me.CodPostal = CodPostal
            Me.Poblacion = Poblacion
            Me.Provincia = Provincia
            Me.IDPais = IDPais
            Me.IdZona = IdZona
            Me.Telefono1 = Telefono1
            Me.Email = Email
            Me.IDMercado = New BdgParametro().Mercado
            Me.guardarProveedor = guardarProveedor
        End Sub
    End Class

    <Task()> Public Shared Function CrearProveedor(ByVal info As stCrearProveedor, ByVal services As ServiceProvider) As DataTable
        Dim p As New Parametro
        Dim bsnProveedor As New Proveedor()
        Dim dtProveedor As DataTable = bsnProveedor.AddNewForm
        If Len(info.IDProveedor) > 0 Then
            dtProveedor.Rows(0)("IDProveedor") = info.IDProveedor
            dtProveedor.Rows(0)("IDContador") = String.Empty
        ElseIf Len(info.IDContador) > 0 Then
            dtProveedor.Rows(0)("IDContador") = info.IDContador
        End If
        dtProveedor.Rows(0)("DescProveedor") = info.DescProveedor
        dtProveedor.Rows(0)("CifProveedor") = info.Cif
        If Len(info.IDGrupoProveedor) > 0 Then
            dtProveedor.Rows(0)("IDGrupoProveedor") = info.IDGrupoProveedor
        End If
        dtProveedor.Rows(0)("IDTipoClasif") = p.TipoClasificacionProveedor
        If dtProveedor.Columns.Contains("IDCNAE") Then dtProveedor.Rows(0)("IDCNAE") = info.IDCNAE

        Dim intNumDigitos As Integer = DigitosCContable()
        If Length(dtProveedor.Rows(0)("CCProveedor")) = 0 Then
            dtProveedor.Rows(0)("CCProveedor") = PuntoPorCero(p.CCProveedor & ".", intNumDigitos)
        End If

        dtProveedor.Rows(0)("FechaAlta") = info.FechaAlta
        dtProveedor.Rows(0)("IDCentroGestion") = info.IDCentroGestion
        dtProveedor.Rows(0)("RazonSocial") = info.RazonSocial
        dtProveedor.Rows(0)("Direccion") = info.Direccion
        dtProveedor.Rows(0)("CodPostal") = info.CodPostal
        dtProveedor.Rows(0)("Poblacion") = info.Poblacion
        dtProveedor.Rows(0)("Provincia") = info.Provincia
        If Length(info.IDPais) > 0 Then dtProveedor.Rows(0)("IDPais") = info.IDPais
        dtProveedor.Rows(0)("IDZona") = info.IdZona
        dtProveedor.Rows(0)("Telefono") = info.Telefono1
        dtProveedor.Rows(0)("Provincia") = info.Provincia
        dtProveedor.Rows(0)("Email") = info.Email
        dtProveedor.Rows(0)("IDMercado") = info.IDMercado

        If Length(dtProveedor.Rows(0)("IDMoneda")) = 0 Then dtProveedor.Rows(0)("IDMoneda") = p.MonedaPred
        If Length(dtProveedor.Rows(0)("IDDiaPago")) = 0 Then dtProveedor.Rows(0)("IDDiaPago") = p.DiaPago
        If Length(dtProveedor.Rows(0)("IDFormaEnvio")) = 0 Then dtProveedor.Rows(0)("IDFormaEnvio") = p.FormaEnvio
        If Length(dtProveedor.Rows(0)("IDCondicionEnvio")) = 0 Then dtProveedor.Rows(0)("IDCondicionEnvio") = p.CondicionEnvio

        If (info.guardarProveedor) Then dtProveedor = bsnProveedor.Update(dtProveedor)

        Return dtProveedor
    End Function

    Private Shared Function DigitosCContable() As Integer
        Dim DigitosCCont As Integer
        Dim StrEjercicio As String = ProcessServer.ExecuteTask(Of Date, String)(AddressOf EjercicioContable.Predeterminado, Today.Date, New ServiceProvider)
        If Length(StrEjercicio) > 0 Then
            DigitosCCont = New EjercicioContable().GetNDigitosAuxiliar(StrEjercicio)
        End If

        Return DigitosCCont
    End Function

    Private Shared Function PuntoPorCero(ByVal pCuenta As String, ByVal pNDigitos As Integer) As String
        Dim strCeros As String
        Dim strC As String

        If InStr(pCuenta, ".") Then
            strC = "."
        ElseIf InStr(pCuenta, ",") Then
            strC = ","
        End If

        If Length(strC) Then
            strCeros = New String("0", pNDigitos - Len(pCuenta) + 1)
            pCuenta = Replace(pCuenta, strC, strCeros, , 1)
        End If
        If InStr(pCuenta, ".") Then
            pCuenta = Replace(pCuenta, ".", "0")
        End If
        If InStr(pCuenta, ",") Then
            pCuenta = Replace(pCuenta, ",", "0")
        End If

        Return (pCuenta)

    End Function

    <Serializable()> _
    Public Class stCrearExplotador
        Public IDExplotador As String
        Public DescExplotador As String

        'Public guardarExplotador As Boolean = False

        Public Sub New(ByVal IDExplotador As String, ByVal DescExplotador As String)
            Me.IDExplotador = IDExplotador
            Me.DescExplotador = DescExplotador
            'Me.guardarExplotador = guardarExplotador
        End Sub
    End Class

    <Task()> Public Shared Function CrearExplotador(ByVal info As stCrearExplotador, ByVal services As ServiceProvider) As DataTable
        Dim exp As New BdgExplotadorFincas
        Dim dt As DataTable = exp.AddNewForm

        dt.Rows(0)("IDExplotadorFincas") = info.IDExplotador
        dt.Rows(0)("DescExplotadorFincas") = info.DescExplotador

        dt = exp.Update(dt)
        Return dt
    End Function

#End Region

End Class

Public Enum BdgCupoCartillaPor
    Manual = 0
    Municipio = 1
    Finca = 2
    Variedad = 3
End Enum

Public Enum BdgTipoCartillista
    Interno = 0
    Externo = 1
End Enum

<Serializable()> _
Public Class _C
    Public Const IDCartillista As String = "IDCartillista"
    Public Const DescCartillista As String = "DescCartillista"
    Public Const IDProveedor As String = "IDProveedor"
    Public Const Talon As String = "Talon"
    Public Const CupoCartillaPor As String = "CupoCartillaPor"
    Public Const IDExplotadorFincas As String = "IDExplotadorFincas"
End Class