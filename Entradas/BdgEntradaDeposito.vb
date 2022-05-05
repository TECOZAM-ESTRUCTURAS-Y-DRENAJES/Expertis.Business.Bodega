Public Class BdgEntradaDeposito
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaDeposito"


#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDArticulo", AddressOf CambioArticulo)
        Obrl.Add("Cantidad", AddressOf CambioCantidad)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) Then
            ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf BdgVino.ValidarArticuloVino, data.Value, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioCantidad(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) = 0 Then ApplicationService.GenerateError("El campo Cantidad debe ser numérico.")
    End Sub

#End Region

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf SacarUvaDeposito)
    End Sub

    <Task()> Public Shared Sub SacarUvaDeposito(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim IDVino As Guid
        Dim DblQ As Double
        If Length(data(_ED.IDVino)) Then
            IDVino = data(_ED.IDVino)
            DblQ = data(_ED.Cantidad)
        End If
        If Not IDVino.Equals(Guid.Empty) Then
            Dim rwE As DataRow = New BdgEntrada().GetItemRow(data(_ED.IDEntrada))
            Dim ClsWork As New BdgWorkClass
            Dim StInc As New BdgWorkClass.StIncrementarIDVino(IDVino, -DblQ)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StIncrementarIDVino)(AddressOf BdgWorkClass.IncrementarCantidadIDVino, StInc, services)
            If Length(rwE(_E.IDMovimiento)) > 0 Then
                Dim StEj As New BdgWorkClass.StEjecutarMovimientos(rwE(_E.NEntrada), rwE(_E.Fecha))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientos, Integer)(AddressOf BdgWorkClass.EjecutarMovimientos, StEj, services)
            Else
                Dim StEjNum As New BdgWorkClass.StEjecutarMovimientosNumero(rwE(_E.IDMovimiento), rwE(_E.NEntrada), rwE(_E.Fecha))
                ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEjNum, services)
            End If
        End If
    End Sub

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_ED.IDEntrada)) = 0 Then ApplicationService.GenerateError("No se ha especificado una entrada")
        If data(_ED.IDEntrada) = 0 Then ApplicationService.GenerateError("No se ha especificado una entrada")
        If Length(data(_ED.IDDeposito)) = 0 Then ApplicationService.GenerateError("No se ha especificado el depósito")
        If Length(data(_ED.Lote)) = 0 Then ApplicationService.GenerateError("No se ha especificado el lote")
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearEntradaDeposito
        Public IDEntrada As Integer
        Public IDDeposito As String
        Public Neto As Double
        Public IDVariedad As String
        Public Vendimia As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDEntrada As Integer, ByVal IDDeposito As String, ByVal Neto As Double, ByVal IDVariedad As String, ByVal Vendimia As String)
            Me.IDEntrada = IDEntrada
            Me.IDDeposito = IDDeposito
            Me.Neto = Neto
            Me.IDVariedad = IDVariedad
            Me.Vendimia = Vendimia
        End Sub
    End Class

    <Task()> Public Shared Function CrearEntradaDeposito(ByVal data As StCrearEntradaDeposito, ByVal services As ServiceProvider) As DataTable
        Dim dtED As DataTable = New BdgEntradaDeposito().AddNew
        If Length(data.IDDeposito) > 0 And data.IDEntrada > 0 Then
            Dim rwED As DataRow = dtED.NewRow
            rwED(_ED.IDEntrada) = data.IDEntrada
            rwED(_ED.IDDeposito) = data.IDDeposito
            rwED(_ED.Cantidad) = data.Neto
            Dim rwVariedad As DataRow = New BdgVariedad().GetItemRow(data.IDVariedad)
            Dim rwVendimia As DataRow = New BdgVendimia().GetItemRow(data.Vendimia)
            If rwVariedad("TipoVariedad") = BdgTipoVariedad.Tinta Then
                rwED(_ED.IDArticulo) = rwVendimia("IDArticuloT")
            Else
                rwED(_ED.IDArticulo) = rwVendimia("IDArticuloB")
            End If
            Dim oParam As New BdgParametro
            rwED(_ED.Lote) = oParam.LotePorDefecto
            dtED.Rows.Add(rwED)
        End If
        Return dtED
    End Function

#End Region

End Class

<Serializable()> _
Public Class _ED
    Public Const IDEntrada As String = "IDEntrada"
    Public Const IDDeposito As String = "IDDeposito"
    Public Const Cantidad As String = "Cantidad"
    Public Const IDVino As String = "IDVino"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const Lote As String = "Lote"
End Class