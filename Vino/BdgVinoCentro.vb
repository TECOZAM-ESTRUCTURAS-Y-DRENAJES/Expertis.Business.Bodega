Public Class BdgVinoCentro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoCentro"

#End Region

#Region "Eventos Entidad"

#Region " GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDCentro", AddressOf BdgGeneral.CambioCentro)
        Obrl.Add("IDIncidencia", AddressOf BdgGeneral.CambioIncidencia)
        Return Obrl
    End Function

#End Region


#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentro")) = 0 Then ApplicationService.GenerateError("El Centro es un dato obligatorio.")
    End Sub

#End Region

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearVinoCentro
        Public Fecha As Date
        Public NOperacion As String
        Public Data As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal NOperacion As String, ByVal Data As DataTable)
            Me.Fecha = Fecha
            Me.NOperacion = NOperacion
            Me.Data = Data
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoCentro(ByVal data As StCrearVinoCentro, ByVal services As ServiceProvider)
        If Not data.Data Is Nothing AndAlso data.Data.Rows.Count > 0 Then
            'Este proceso si es comun para un nuevo centro y para una modificacion debera de ir en el Update.
            'Pero no se puede poner comun xq se necesita un IDVino en la creacion y no se necesita en la modificacion
            'aunq el proceso de tasas si q es comun
            Dim ClsCentro As New BdgVinoCentro
            Dim dtVC As DataTable = ClsCentro.AddNew
            For Each oRw As DataRow In data.Data.Rows
                Dim rwVC As DataRow = dtVC.NewRow
                rwVC(_VC.IdVinoCentro) = Guid.NewGuid
                rwVC(_VC.IDVino) = oRw(_VC.IDVino)
                rwVC(_VC.IDCentro) = oRw(_VC.IDCentro)
                rwVC(_VC.Fecha) = data.Fecha
                If Length(data.NOperacion) <> 0 Then rwVC(_VC.NOperacion) = data.NOperacion
                rwVC(_VC.Tiempo) = oRw(_VC.Tiempo)

                Dim StGet As New StGetTasa(oRw(_VC.IDCentro), data.Fecha)
                Dim oTasa As TasaInfo = ProcessServer.ExecuteTask(Of StGetTasa, TasaInfo)(AddressOf GetTasa, StGet, services)
                If Not oTasa Is Nothing Then
                    rwVC(_VC.UDTiempo) = oTasa.UdTiempo
                    rwVC(_VC.TasaD) = oTasa.TasaD
                    rwVC(_VC.TasaI) = oTasa.TasaI
                    rwVC(_VC.TasaF) = oTasa.TasaF
                    rwVC(_VC.TasaV) = oTasa.TasaV
                    rwVC(_VC.Tasa) = oTasa.Tasa
                    rwVC(_VC.TasaFiscal) = oTasa.TasaFscl
                End If

                rwVC(_VC.Cantidad) = oRw(_VC.Cantidad)
                rwVC(_VC.PorCantidad) = oRw(_VC.PorCantidad)
                rwVC(_VC.IdUdMedidaCentro) = oRw(_VC.IdUdMedidaCentro)
                rwVC(_VC.IdUdMedidaArticulo) = oRw(_VC.IdUdMedidaArticulo)
                rwVC(_VC.IDOperacionCentro) = oRw(_VC.IDOperacionCentro)
                dtVC.Rows.Add(rwVC)
            Next
            ClsCentro.Update(dtVC)
        End If
    End Function

    <Serializable()> _
    Public Class StGetTasa
        Public IDCentro As String
        Public Fecha As Date
        Public IDIncidencia As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDCentro As String, ByVal Fecha As Date, Optional ByVal IDIncidencia As String = "")
            Me.IDCentro = IDCentro
            Me.Fecha = Fecha
            Me.IDIncidencia = IDIncidencia
        End Sub
    End Class

    <Task()> Public Shared Function GetTasa(ByVal data As StGetTasa, ByVal services As ServiceProvider) As TasaInfo
        If Length(data.IDCentro) > 0 Then
            Dim oTs As New TasaInfo

            oTs.IDCentro = data.IDCentro
            Dim dtCentro As DataTable = New Centro().SelOnPrimaryKey(data.IDCentro)
            If dtCentro.Rows.Count > 0 Then
                oTs.UdTiempo = dtCentro.Rows(0)("UdTiempo") & String.Empty
                oTs.IDUdMedida = dtCentro.Rows(0)("IDUDMedida") & String.Empty
            End If

            Dim f As New Filter
            f.Add(New StringFilterItem("IDCentro", data.IDCentro))
            f.Add(New DateFilterItem("FechaDesde", FilterOperator.LessThanOrEqual, data.Fecha))
            f.Add(New DateFilterItem("FechaHasta", FilterOperator.GreaterThanOrEqual, data.Fecha))

            Dim dtCntrTs As DataTable = New CentroTasa().Filter(f)
            For Each oRw As DataRow In dtCntrTs.Rows
                Dim dblTasa As Double
                If Length(data.IDIncidencia) > 0 Then
                    dblTasa = Nz(oRw("PreparacionValorA"), 0)
                Else
                    dblTasa = Nz(oRw("EjecucionValorA"), 0)
                End If

                If CType(oRw("TipoCosteDI"), BusinessEnum.enumtcdiTipoCoste) = enumtcdiTipoCoste.tcdiDirecto Then
                    oTs.TasaD += dblTasa
                Else
                    oTs.TasaI += dblTasa
                End If
                If CType(oRw("TipoCosteFV"), BusinessEnum.enumtcfvTipoCoste) = enumtcfvTipoCoste.tcfvFijo Then
                    oTs.TasaF += dblTasa
                Else
                    oTs.TasaV += dblTasa
                End If
                If Nz(oRw("Fiscal"), False) Then
                    oTs.TasaFscl += dblTasa
                End If
            Next
            Return oTs
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _VC
    Public Const IdVinoCentro As String = "IdVinoCentro"
    Public Const IDVino As String = "IDVino"
    Public Const IDCentro As String = "IDCentro"
    Public Const Fecha As String = "Fecha"
    Public Const NOperacion As String = "NOperacion"
    Public Const Tiempo As String = "Tiempo"
    Public Const TiempoPorcen As String = "TiempoPorcen"
    Public Const UDTiempo As String = "UDTiempo"
    Public Const Tasa As String = "Tasa"
    Public Const Cantidad As String = "Cantidad"
    Public Const PorCantidad As String = "PorCantidad"
    Public Const TasaF As String = "TasaF"
    Public Const TasaV As String = "TasaV"
    Public Const TasaD As String = "TasaD"
    Public Const TasaI As String = "TasaI"
    Public Const TasaFiscal As String = "TasaFiscal"
    Public Const IDHist As String = "IDHist"
    Public Const IdUdMedidaCentro As String = "IdUdMedidaCentro"
    Public Const IdUdMedidaArticulo As String = "IdUdMedidaArticulo"
    Public Const IDOperacionCentro As String = "IDOperacionCentro"
End Class

<Serializable()> _
Public Class TasaInfo
    Public IDCentro As String
    Public TasaF As Double
    Public TasaV As Double
    Public TasaD As Double
    Public TasaI As Double
    Public TasaFscl As Double
    Public UdTiempo As enumstdUdTiempo
    Public IDUdMedida As String
    Public Function Tasa() As Double
        Return TasaF + TasaV
    End Function
End Class