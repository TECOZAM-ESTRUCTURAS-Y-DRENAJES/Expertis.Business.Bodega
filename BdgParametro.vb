Public Class BdgParametro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbParametro"

#End Region

#Region " Reemplazos "

    Public Overloads Function SelOnPrimaryKey(ByVal IDParametro As String) As DataTable
        Dim dt As DataTable = MyBase.SelOnPrimaryKey(IDParametro)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe el registro del Parametro |", IDParametro)
        End If
        Return dt
    End Function

    Friend Shared Function SelOnPrimaryKeySH(ByVal IDParametro As String) As DataTable
        Return New BdgParametro().SelOnPrimaryKey(IDParametro)
    End Function

#End Region

    '#Region " DAA "

    '    Public Function TipoClienteAduana() As String
    '        Const cIDParam As String = "BDGADUCLI"
    '        Dim dt As DataTable = SelOnPrimaryKey(cIDParam)
    '        If dt.Rows.Count > 0 Then
    '            If dt.Rows(0).IsNull("Valor") Then
    '                ApplicationService.GenerateError("El parámetro |: | no tiene valor", cIDParam, dt.Rows(0)("DescParametro"))
    '            Else
    '                Return dt.Rows(0)("Valor")
    '            End If
    '        End If
    '    End Function

    '    Public Function PredeterminarDireccionesDAA() As Boolean
    '        Const cIDParam As String = "BDGDIRDAA"
    '        Dim dt As DataTable = SelOnPrimaryKey(cIDParam)
    '        If dt.Rows.Count > 0 Then
    '            If dt.Rows(0).IsNull("Valor") Then
    '                PredeterminarDireccionesDAA = CBool(0)
    '            ElseIf IsNumeric(dt.Rows(0)("Valor")) Then
    '                PredeterminarDireccionesDAA = CBool(dt.Rows(0)("Valor"))
    '            End If
    '        End If
    '    End Function

    '    Public Function ContadorAnexoDAA() As String
    '        Const cIDParam As String = "BDGCAnexDA"
    '        Dim dt As DataTable = SelOnPrimaryKey(cIDParam)
    '        If dt.Rows.Count > 0 Then
    '            If dt.Rows(0).IsNull("Valor") Then
    '                ApplicationService.GenerateError("El parámetro |: | no tiene valor", cIDParam, dt.Rows(0)("DescParametro"))
    '            Else
    '                Return dt.Rows(0)("Valor")
    '            End If
    '        End If

    '    End Function

    '    Public Function BdgDAAGarantia() As String

    '        Dim dt As DataTable = SelOnPrimaryKey("BDGDAAGAR")
    '        If dt.Rows.Count > 0 Then
    '            BdgDAAGarantia = dt.Rows(0)("Valor") & String.Empty
    '        End If

    '    End Function

    '    Public Function AgrupSubFamilia() As String
    '        Dim dt As DataTable = SelOnPrimaryKey("BDGDAASUBF")
    '        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
    '            Return dt.Rows(0)("Valor") & String.Empty
    '        End If
    '    End Function

    '    Public Function BdgDAAFirmante() As String
    '        Dim dt As DataTable = SelOnPrimaryKey("BDGDAAFIR")
    '        If dt.Rows.Count > 0 Then
    '            BdgDAAFirmante = dt.Rows(0)("Valor") & String.Empty
    '        End If
    '    End Function

    '    Public Function DatosTransporteAutomatico() As Boolean

    '        Dim dt As DataTable = SelOnPrimaryKey("BDGTRANAUT")
    '        If dt.Rows.Count > 0 Then
    '            If dt.Rows(0).IsNull("Valor") Then
    '                DatosTransporteAutomatico = CBool(0)
    '            ElseIf IsNumeric(dt.Rows(0)("Valor")) Then
    '                DatosTransporteAutomatico = CBool(dt.Rows(0)("Valor"))
    '            End If
    '        End If

    '    End Function

    '#End Region

#Region " Entradas Uva "

    Public Function BdgDecimalesEntradas() As Integer

        Dim dt As DataTable = SelOnPrimaryKey("BDGDECENT")
        If dt.Rows.Count > 0 Then
            BdgDecimalesEntradas = dt.Rows(0)("Valor")
        End If

    End Function

    Public Function BdgArticuloCajaEU() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGARTCAJA")
        If dt.Rows.Count > 0 Then
            BdgArticuloCajaEU = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function BdgArticuloPaletEU() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGARTPALE")
        If dt.Rows.Count > 0 Then
            BdgArticuloPaletEU = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function BodegaPredetEntradaUva() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGBODPRE")
        If dt.Rows.Count > 0 Then
            BodegaPredetEntradaUva = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function EnvasePredetEntrada() As Integer

        Dim dt As DataTable = SelOnPrimaryKey("BDGENVENT")
        If dt.Rows.Count > 0 Then
            EnvasePredetEntrada = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function Mercado() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGMERCADO")
        If dt.Rows.Count > 0 Then
            Mercado = dt.Rows(0)("Valor") & String.Empty
        End If
    End Function

    Public Function CalculoFraEnDetalle() As Boolean

        Dim dt As DataTable = SelOnPrimaryKey("BDGFRADTLL")
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).IsNull("Valor") Then
            ElseIf IsNumeric(dt.Rows(0)("Valor")) Then
                CalculoFraEnDetalle = CBool(dt.Rows(0)("Valor"))
            End If
        End If

    End Function

    Public Function DesgloseEnFacturacion() As Integer
        Dim dt As DataTable = SelOnPrimaryKey("BDGFRADSGL")
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).IsNull("Valor") Then
            ElseIf IsNumeric(dt.Rows(0)("Valor")) Then
                DesgloseEnFacturacion = Nz(dt.Rows(0)("Valor"), 0)
            End If
        End If

    End Function

    Public Function PropietarioFincasPropias() As String
        Dim dt As DataTable = SelOnPrimaryKey("PROPFINCA")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor") & String.Empty
        End If
    End Function

    Public Function TipoRecogidaObligatorioEntradas() As Boolean
        Dim dt As DataTable = SelOnPrimaryKey("BDGTRCGOBL")
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).IsNull("Valor") Then
            ElseIf IsNumeric(dt.Rows(0)("Valor")) Then
                TipoRecogidaObligatorioEntradas = CBool(dt.Rows(0)("Valor"))
            End If
        End If

    End Function

    Public Function TipoRecogidaPredefinidoEntradas() As String
        Dim dt As DataTable = SelOnPrimaryKey("BDGTRCGPRE")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor") & String.Empty
        End If
    End Function

    Public Function DetectarConexionEntUva() As Boolean
        Dim dt As DataTable = SelOnPrimaryKey("BDGENTCON")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        End If
    End Function

    Public Function DireccionIPConexionEntUva() As String
        Dim dt As DataTable = SelOnPrimaryKey("BDGENTIPCN")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor") & String.Empty
        End If
    End Function

    Public Function SegundosComprobacionConexionEntUva() As Integer
        Dim dt As DataTable = SelOnPrimaryKey("BDGENTSECN")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        End If
    End Function

    Public Function GestionExplotadorObligatoria() As Boolean
        Dim dt As DataTable = New General.Parametro().SelOnPrimaryKey("BDGGESTEXP")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        Else
            Return False
        End If
    End Function

    Public Function ValidarCupoFincas() As Boolean
        Dim Dt As DataTable = SelOnPrimaryKey("BDGENTCUPO")
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            Return Dt.Rows(0)("Valor")
        Else : Return False
        End If
    End Function

    Public Function ControlCuposEntradasUva() As Boolean
        Dim Dt As DataTable = SelOnPrimaryKey("BDGCUPO")
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            Return Dt.Rows(0)("Valor")
        Else : Return False
        End If
    End Function

    Public Function ComprobarDesgloseEntradaUva() As Boolean
        Dim Dt As DataTable = SelOnPrimaryKey("BDGDSGEUVA")
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            Return Dt.Rows(0)("Valor")
        Else : Return False
        End If
    End Function

#End Region

#Region " Lote "
    Public Function LotePorDefecto() As String
        Dim rw As DataRow = Me.GetItemRow("BDGLOTEDEF")
        If rw.IsNull("Valor") Then
            '//el lote es obligatorio. Si no está parametrizado hay que proporcinarlo
            'Return "Granel"
        Else
            Return rw("Valor")
        End If
    End Function

    'Friend Shared Function LotePorDefectoSH() As String
    '    Dim oo As BdgParametro = New BdgParametro
    '    Return oo.LotePorDefecto()
    'End Function

    Public Function LoteExplicitoEnBotellero() As Boolean
        Dim rw As DataRow = Me.GetItemRow("BDGLOTEBOT")
        If rw.IsNull("Valor") Then
            '//el lote es obligatorio. Si no está parametrizado hay que proporcinarlo
            'Return "Granel"
        Else
            Return rw("Valor")
        End If
    End Function

    Public Function LotePorEntrada() As String
        Dim dt As DataTable = MyBase.SelOnPrimaryKey("BDGLOTEENT")
        If dt Is Nothing OrElse dt.Rows.Count = 0 OrElse dt.Rows(0).IsNull("Valor") Then
            '//el lote es obligatorio. Si no está parametrizado hay que proporcinarlo
            'Return "Granel"
        Else
            Return dt.Rows(0)("Valor")
        End If
    End Function

#End Region

#Region " Análisis "
    Public Function AnalisisEntrada() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGANAENT")
        If dt.Rows.Count > 0 Then
            AnalisisEntrada = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function AnalisisEntradaVino() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGANAVIN")
        If dt.Rows.Count > 0 Then
            AnalisisEntradaVino = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function VariableGrado() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGVARGRAD")
        If dt.Rows.Count > 0 Then
            VariableGrado = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function VariableCalidad() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGVARCAL")
        If dt.Rows.Count > 0 Then
            VariableCalidad = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function VariableGradoVino() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGVARGRVN")
        If dt.Rows.Count > 0 Then
            VariableGradoVino = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function GradoArticulo() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGCARGRAD")
        If dt.Rows.Count > 0 Then
            GradoArticulo = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function MesInicialAnalisisFinca() As Integer
        Dim valor As Integer
        Dim dt As DataTable = SelOnPrimaryKey("BDGMESINI")
        If dt.Rows.Count > 0 Then
            valor = Nz(dt.Rows(0)("Valor"), 0)
        End If
        Return valor
    End Function

    Public Function MesFinalAnalisisFinca() As Integer
        Dim valor As Integer
        Dim dt As DataTable = SelOnPrimaryKey("BDGMESFIN")
        If dt.Rows.Count > 0 Then
            valor = Nz(dt.Rows(0)("Valor"), 0)
        End If
        Return valor
    End Function

    Public Function PathExcelAnalisisOrigen() As String
        Return New Parametro().LeerParametro("PHXLSFNOR")
    End Function

    Public Function PathExcelAnaliticaDestino() As String
        Return New Parametro().LeerParametro("PHXLSFNDES")
    End Function

#End Region

#Region " Operaciones "

    Public Function FechaNuevaOperacionIgualUltimaOperacion() As Boolean
        Dim dt As DataTable = SelOnPrimaryKey("BDGFECHOPE")
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).IsNull("Valor") Then
            ElseIf IsNumeric(dt.Rows(0)("Valor")) Then
                FechaNuevaOperacionIgualUltimaOperacion = CBool(dt.Rows(0)("Valor"))
            End If
        End If
    End Function

    Public Function UnidadesCampoLitros() As String

        Dim dt As DataTable = SelOnPrimaryKey("BDGUDMED")
        If dt.Rows.Count > 0 Then
            UnidadesCampoLitros = dt.Rows(0)("Valor") & String.Empty
        End If

    End Function

    Public Function AvisoMermas() As Boolean

        Dim dt As DataTable = SelOnPrimaryKey("BDGAVIMERM")
        If dt.Rows.Count > 0 Then
            If dt.Rows(0).IsNull("Valor") Then
            ElseIf IsNumeric(dt.Rows(0)("Valor")) Then
                AvisoMermas = CBool(dt.Rows(0)("Valor"))
            End If
        End If

    End Function
    Public Function BdgComprobarFechaOperacion() As Boolean

        Dim stRdo As Boolean
        Dim dt As DataTable = SelOnPrimaryKey("BDGCTLFECH")
        If dt.Rows.Count > 0 Then
            stRdo = CBool(dt.Rows(0)("Valor"))
        Else
            ApplicationService.GenerateError("No existe el Parámetro BDGCTLFECH.")
        End If

        BdgComprobarFechaOperacion = stRdo

    End Function

#End Region

    '#Region " EMCS "
    '    Public Function EmcsCodigoZonaVitivinicola() As String

    '        Dim stRdo As String
    '        Dim dt As DataTable = SelOnPrimaryKey("EmcsWGZ")
    '        If dt.Rows.Count > 0 Then
    '            stRdo = dt.Rows(0)("Valor") & String.Empty
    '        Else
    '            ApplicationService.GenerateError("No existe el Parámetro EmcsWGZ.")
    '        End If

    '        If Length(stRdo & String.Empty) = 0 Then
    '            ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '        End If

    '        EmcsCodigoZonaVitivinicola = stRdo

    '    End Function

    '    Public Function EmcsCodigoManipulacionVitivinicola() As String

    '        Dim stRdo As String
    '        Dim dt As DataTable = SelOnPrimaryKey("EmcsWOC")
    '        If dt.Rows.Count > 0 Then
    '            stRdo = dt.Rows(0)("Valor") & String.Empty
    '        Else
    '            ApplicationService.GenerateError("No existe el Parámetro EmcsWOC.")
    '        End If

    '        If Length(stRdo & String.Empty) = 0 Then
    '            ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '        End If

    '        EmcsCodigoManipulacionVitivinicola = stRdo

    '    End Function

    '    'Public Function EmcsTipoInformacion() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsSMT")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsSMT.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsTipoInformacion = stRdo

    '    'End Function

    '    'Public Function EmcsLenguaGruposDatos() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsLNG")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsLNG.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsLenguaGruposDatos = stRdo

    '    'End Function

    '    'Public Function EmcsCodigoAduanaImportacion() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsORN")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsORN.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsCodigoAduanaImportacion = stRdo

    '    'End Function

    '    'Public Function EmcsCodigoMedioTransporte() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsTUC")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsTUC.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsCodigoMedioTransporte = stRdo

    '    'End Function

    '    'Public Function EmcsTipoExpedidorDocumento() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsOTC")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsOTC.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsTipoExpedidorDocumento = stRdo

    '    'End Function

    '    'Public Function EmcsOrganizacionTransporte() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsTA")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsTA.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsOrganizacionTransporte = stRdo

    '    'End Function

    '    'Public Function EmcsTipoGarante() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsGTC")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsGTC.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsTipoGarante = stRdo

    '    'End Function

    '    Public Function EmcsRutaTemporal() As String

    '        Dim stRdo As String
    '        Dim dt As DataTable = SelOnPrimaryKey("EmcsTMP")
    '        If dt.Rows.Count > 0 Then
    '            stRdo = dt.Rows(0)("Valor") & String.Empty
    '        Else
    '            ApplicationService.GenerateError("No existe el Parámetro EmcsTMP.")
    '        End If

    '        If Length(stRdo & String.Empty) = 0 Then
    '            ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '        End If

    '        EmcsRutaTemporal = stRdo

    '    End Function

    '    Public Function EmcsRutaProcesados() As String

    '        Dim stRdo As String
    '        Dim dt As DataTable = SelOnPrimaryKey("EmcsFRP")
    '        If dt.Rows.Count > 0 Then
    '            stRdo = dt.Rows(0)("Valor") & String.Empty
    '        Else
    '            ApplicationService.GenerateError("No existe el Parámetro EmcsFRP.")
    '        End If

    '        If Length(stRdo & String.Empty) = 0 Then
    '            ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '        End If

    '        EmcsRutaProcesados = stRdo

    '    End Function

    '    Public Function EmcsRutaFicherosRespuesta() As String

    '        Dim stRdo As String
    '        Dim dt As DataTable = SelOnPrimaryKey("EmcsRLF")
    '        If dt.Rows.Count > 0 Then
    '            stRdo = dt.Rows(0)("Valor") & String.Empty
    '        Else
    '            ApplicationService.GenerateError("No existe el Parámetro EmcsRLF.")
    '        End If

    '        If Length(stRdo & String.Empty) = 0 Then
    '            ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '        End If

    '        EmcsRutaFicherosRespuesta = stRdo

    '    End Function

    '    'Public Function EmcsCodigoCancelacion() As String

    '    '    Dim stRdo As String
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsCRC")
    '    '    If dt.Rows.Count > 0 Then
    '    '        stRdo = dt.Rows(0)("Valor") & String.Empty
    '    '    Else
    '    '        ApplicationService.GenerateError("No existe el Parámetro EmcsCRC.")
    '    '    End If

    '    '    If Length(stRdo & String.Empty) = 0 Then
    '    '        ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
    '    '    End If

    '    '    EmcsCodigoCancelacion = stRdo

    '    'End Function

    '    'Public Function EmcsAgruparLineas() As Boolean

    '    '    'Por defecto agrupamos
    '    '    Dim blnRdo As Boolean = True
    '    '    Dim dt As DataTable = SelOnPrimaryKey("EmcsAgrupL")
    '    '    If dt.Rows.Count > 0 Then
    '    '        blnRdo = Nz(dt.Rows(0)("Valor"), blnRdo)
    '    '    End If

    '    '    EmcsAgruparLineas = blnRdo

    '    'End Function

    '#End Region

#Region " Libros Bodega"
    Public Function BdgZonaVitivinicolaLibros() As String

        Dim stRdo As String
        Dim dt As DataTable = SelOnPrimaryKey("BDGZVL")
        If dt.Rows.Count > 0 Then
            stRdo = dt.Rows(0)("Valor") & String.Empty
        Else
            ApplicationService.GenerateError("No existe el Parámetro BDGZVL.")
        End If

        If Length(stRdo & String.Empty) = 0 Then
            ApplicationService.GenerateError("Introduzca un valor en el Parámetro | | .", dt.Rows(0)("IDParametro"), dt.Rows(0)("DescParametro"))
        End If

        BdgZonaVitivinicolaLibros = stRdo

    End Function
#End Region

#Region "Gestión"

    Public Function GestionCartillistas() As Boolean
        Dim dt As DataTable = New General.Parametro().SelOnPrimaryKey("BDG_GECART")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        Else
            Return True
        End If
    End Function

    Public Function GestionPorFincaPadre() As Boolean
        Dim dt As DataTable = New BE.DataEngine().Filter("tbParametro", New StringFilterItem("IDParametro", "CTRFINCAPA"))
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        Else
            ApplicationService.GenerateError("No existe el Parámetro CTRFINCAPA.")
        End If

        Return False
    End Function

#End Region

#Region "Partes"

    Public Function VolcarCantidadPartesAgrupados() As Boolean
        Dim dt As DataTable = New General.Parametro().SelOnPrimaryKey("BDGCANTPAR")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        Else
            Return True
        End If
    End Function

    Public Function AutoactualizacionStocksPartesAgrupados() As Boolean
        Dim dt As DataTable = New General.Parametro().SelOnPrimaryKey("BDGACSTKPA")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        Else
            Return False
        End If
    End Function

    Public Function CopiarImputacionesDuplicarParte() As Boolean
        Dim dt As DataTable = New General.Parametro().SelOnPrimaryKey("BDGCOPDUP")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Valor")
        Else
            Return False
        End If
    End Function

#End Region

#Region "Trazabilidad"
    Public Function BdgNumeroMaximoNodos() As Long
        BdgNumeroMaximoNodos = 100
        Dim dt As DataTable = SelOnPrimaryKey("BDGNODOMX")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            BdgNumeroMaximoNodos = dt.Rows(0)("Valor")
        End If
    End Function
#End Region

#Region " Fuera Inventario "

    Public Function BdgGestionFueraInventario() As Boolean
        BdgGestionFueraInventario = False
        Dim dt As DataTable = MyBase.SelOnPrimaryKey("BDGFUEINV")
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            BdgGestionFueraInventario = dt.Rows(0)("Valor")
        End If
    End Function

#End Region

End Class