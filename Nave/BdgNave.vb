Public Class BdgNave

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgNave"

#End Region

#Region "Eventos Entidad"

    Public Overloads Function GetItemRow(ByVal IDNave As String) As DataRow
        Dim dt As DataTable = New BdgNave().SelOnPrimaryKey(IDNave)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("No existe la nave |", IDNave)
        Else : Return dt.Rows(0)
        End If
    End Function

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentroGestion")) = 0 Then ApplicationService.GenerateError("El Centro de Gestión es obligatorio.")
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StValidarFechasEstanciaNave
        Public IDNave As String
        Public FechaDesde As DateTime
        Public FechaHasta As DateTime

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDNave As String, ByVal FechaDesde As DateTime, ByVal FecahHasta As DateTime)
            Me.IDNave = IDNave
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub
    End Class

    <Task()> Public Shared Function ValidarFechasEstanciaNave(ByVal data As StValidarFechasEstanciaNave, ByVal services As ServiceProvider) As Boolean
        Dim oNH As New BdgNaveHist
        Dim f As New Filter
        f.Add(New StringFilterItem(_NH.IDNave, FilterOperator.Equal, data.IDNave))
        Dim dtNH As DataTable = oNH.Filter(f)

        If Not IsNothing(dtNH) AndAlso dtNH.Rows.Count > 0 Then
            For Each dr As DataRow In dtNH.Select(Nothing, _NH.FechaDesde)
                If data.FechaDesde >= dr(_NH.FechaDesde) And data.FechaDesde <= dr(_NH.FechaHasta) Then
                    Return False
                End If
                If data.FechaHasta >= dr(_NH.FechaDesde) And data.FechaHasta <= dr(_NH.FechaHasta) Then
                    Return False
                End If
                If dr(_NH.FechaDesde) >= data.FechaDesde And dr(_NH.FechaDesde) <= data.FechaHasta Then
                    Return False
                End If
                If dr(_NH.FechaHasta) >= data.FechaDesde And dr(_NH.FechaHasta) <= data.FechaHasta Then
                    Return False
                End If
            Next
        End If
        Return True
    End Function

    <Serializable()> _
    Public Class StCrearVinosCentro
        Public IDNave As String
        Public FechaDesde As DateTime
        Public FechaHasta As DateTime

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDNave As String, ByVal FechaDesde As DateTime, ByVal FechaHasta As DateTime)
            Me.IDNave = IDNave
            Me.FechaDesde = FechaDesde
            Me.FechaHasta = FechaHasta
        End Sub
    End Class

    
    <Task()> Public Shared Sub CrearVinosCentro(ByVal data As StCrearVinosCentro, ByVal services As ServiceProvider)
        If Length(data.IDNave) = 0 Then Exit Sub

        '//Validamos que tengas un centro asociado
        Dim IDCentro As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf ValidarCentroNave, data.IDNave, services)

        '//Comprobar periodos imputados
        ProcessServer.ExecuteTask(Of StCrearVinosCentro)(AddressOf ValidarFechasImputacion, data, services)


        '//Historico de Naves
        Dim ClsNH As New BdgNaveHist
        Dim dtNH As DataTable = ClsNH.AddNew()

        Dim IDHist As Guid = Guid.NewGuid
        Dim rwNH As DataRow = dtNH.NewRow
        rwNH(_NH.IDNave) = data.IDNave
        rwNH(_NH.IDHist) = IDHist
        rwNH(_NH.FechaDesde) = data.FechaDesde
        rwNH(_NH.FechaHasta) = data.FechaHasta
        rwNH(_NH.Fecha) = Date.Today
        dtNH.Rows.Add(rwNH)

        '//ajuste para no perder un dia de imputación por cada proceso
        data.FechaHasta = data.FechaHasta.AddDays(1)

        '//Traemos los movimientos de los vinos
        Dim dttMovimientosVino As DataTable = AdminData.Execute("spBdgCosteNaveMovimientosVino", False, data.IDNave, data.FechaDesde, data.FechaHasta)

        '//Traemos las tasas del centro asociado
        Dim StVCentro As New BdgVinoCentro.StGetTasa(IDCentro, data.FechaHasta)
        Dim oTasa As TasaInfo = ProcessServer.ExecuteTask(Of BdgVinoCentro.StGetTasa, TasaInfo)(AddressOf BdgVinoCentro.GetTasa, StVCentro, services)

        '//Preparamos los datos de BdgVinoCentro
        Dim oVC As New BdgVinoCentro
        Dim dtVC As DataTable = oVC.AddNew
        If Not dttMovimientosVino Is Nothing AndAlso dttMovimientosVino.Rows.Count > 0 Then
            Dim MovimientosVino As List(Of DataRow) = (From c In dttMovimientosVino Order By c("IDVino"), c("FechaInicio") Select c).ToList
            If Not MovimientosVino Is Nothing AndAlso MovimientosVino.Count > 0 Then
                For Each oRw As DataRow In MovimientosVino
                    Dim nwRw As DataRow = dtVC.NewRow
                    nwRw(_VC.IdVinoCentro) = Guid.NewGuid
                    nwRw(_VC.IDHist) = IDHist
                    nwRw(_VC.IDVino) = oRw(_V.IDVino)
                    nwRw(_VC.Fecha) = data.FechaHasta

                    'La columna Tiempo siempre está en Horas.
                    nwRw(_VC.Tiempo) = oRw(_V.FechaFin).Subtract(oRw(_V.FechaInicio)).Days * 24

                    'Datos del centro
                    nwRw(_VC.IDCentro) = IDCentro
                    nwRw(_VC.UDTiempo) = oTasa.UdTiempo 'UDTiempo del centro
                    nwRw(_VC.Tasa) = oTasa.Tasa
                    nwRw(_VC.PorCantidad) = True
                    nwRw(_VC.TasaF) = oTasa.TasaF
                    nwRw(_VC.TasaV) = oTasa.TasaV
                    nwRw(_VC.TasaD) = oTasa.TasaD
                    nwRw(_VC.TasaI) = oTasa.TasaI
                    nwRw(_VC.TasaFiscal) = oTasa.TasaFscl

                    nwRw(_VC.Cantidad) = oRw("CantidadEnFecha") 'En la unidad del artículo
                    nwRw(_VC.IdUdMedidaCentro) = oTasa.IDUdMedida
                    nwRw(_VC.IdUdMedidaArticulo) = oRw("IdUdMedidaArticulo") & String.Empty

                    nwRw("FechaInicioCosteNave") = oRw(_V.FechaInicio)
                    nwRw("FechaFinCosteNave") = oRw(_V.FechaFin)
                    dtVC.Rows.Add(nwRw)
                Next
            End If
        End If


        AdminData.BeginTx()
        ClsNH.Update(dtNH) 'Historico de Naves
        If Not dtVC Is Nothing AndAlso dtVC.Rows.Count > 0 Then
            oVC.Update(dtVC)
            Dim StCrear As New BdgVinoCentroTasa.StCrearVinoCentroTasa(data.FechaHasta, dtVC)
            ProcessServer.ExecuteTask(Of BdgVinoCentroTasa.StCrearVinoCentroTasa)(AddressOf BdgVinoCentroTasa.CrearVinoCentroTasa, StCrear, services)
        End If
        AdminData.CommitTx(True)
    End Sub

    <Task()> Public Shared Function ValidarCentroNave(ByVal IDNave As String, ByVal services As ServiceProvider) As String
        If Length(IDNave) > 0 Then
            Dim rwNv As DataRow = New BdgNave().GetItemRow(IDNave)
            '//Comprobar centro de trabajo.
            Dim IDCentro As String = rwNv(_N.IDCentro) & String.Empty
            If Length(IDCentro) = 0 Then
                ApplicationService.GenerateError("La Nave no tiene asignado un Centro de Trabajo del que obtener las tasas.")
            End If
            Return IDCentro
        End If
    End Function

    <Task()> Public Shared Sub ValidarFechasImputacion(ByVal data As StCrearVinosCentro, ByVal services As ServiceProvider)
        If Length(data.IDNave) > 0 Then
            If data.FechaDesde > data.FechaHasta Then
                ApplicationService.GenerateError("La Fecha Desde no puede ser mayor que la Fecha Hasta.")
            End If
            Dim StVal As New StValidarFechasEstanciaNave(data.IDNave, data.FechaDesde, data.FechaHasta)
            If Not ProcessServer.ExecuteTask(Of StValidarFechasEstanciaNave, Boolean)(AddressOf ValidarFechasEstanciaNave, StVal, services) Then
                ApplicationService.GenerateError("El periodo indicado o parte de él ya tiene coste de estancia en nave.")
            End If
        End If
    End Sub


    <Serializable()> _
    Public Class StFromDaysTo
        Public Value As Double
        Public UDTiempo As enumstdUdTiempo

        Public Sub New()
        End Sub

        Public Sub New(ByVal Value As Double, ByVal UDTiempo As enumstdUdTiempo)
            Me.Value = Value
            Me.UDTiempo = UDTiempo
        End Sub
    End Class

    <Task()> Public Shared Function FromDaysTo(ByVal data As StFromDaysTo, ByVal services As ServiceProvider) As Double
        Select Case data.UDTiempo
            Case enumstdUdTiempo.Dias
                Return data.Value
            Case enumstdUdTiempo.Horas
                Return data.Value * 24
            Case enumstdUdTiempo.Minutos
                Return data.Value * 24 * 60
            Case enumstdUdTiempo.Segundos
                Return data.Value * 24 * 3600
        End Select
    End Function

    <Task()> Public Shared Function Almacen(ByVal IDNave As String, ByVal services As ServiceProvider) As String
        If Length(IDNave) > 0 Then
            Dim rslt As Object = New BdgNave().GetItemRow(IDNave)(_N.IDAlmacen)
            If TypeOf rslt Is String Then Return rslt
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class _N
    Public Const IDNave As String = "IDNave"
    Public Const DescNave As String = "DescNave"
    Public Const IDCentro As String = "IDCentro"
    Public Const Imagen As String = "Imagen"
    Public Const IDAlmacen As String = "IDAlmacen"
End Class