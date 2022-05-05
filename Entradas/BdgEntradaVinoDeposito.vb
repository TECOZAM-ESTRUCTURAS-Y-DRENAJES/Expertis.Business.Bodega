Public Class BdgEntradaVinoDeposito

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaVinoDeposito"

#End Region

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf SacarVinoDeposito)
    End Sub

    <Task()> Public Shared Sub SacarVinoDeposito(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim VinoEnt As Guid
        If Length(data(_EVD.IDVino)) > 0 Then VinoEnt = data(_EVD.IDVino)
        Dim dblQOld As Double = data(_EVD.Cantidad)
        Dim StInc As New BdgWorkClass.StIncrementarIDVino(VinoEnt, -dblQOld)
        If dblQOld > 0 Then ProcessServer.ExecuteTask(Of BdgWorkClass.StIncrementarIDVino)(AddressOf BdgWorkClass.IncrementarCantidadIDVino, StInc, services)
        Dim rwE As DataRow = New BdgEntradaVino().GetItemRow(data("NEntrada"))
        Dim StEj As New BdgWorkClass.StEjecutarMovimientosNumero(rwE(_EVn.IDMovimiento), data("NEntrada"), rwE(_EVn.Fecha))
        ProcessServer.ExecuteTask(Of BdgWorkClass.StEjecutarMovimientosNumero)(AddressOf BdgWorkClass.EjecutarMovimientosNumero, StEj, services)

        ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgVino.ActualizarUltimoEstadoVino, VinoEnt, services)
    End Sub


#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDepositoObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarLoteObligatorio)
    End Sub

    <Task()> Public Shared Sub ValidarDepositoObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_EVD.IDDeposito)) = 0 Then ApplicationService.GenerateError("El Depósito es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarLoteObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Lote")) = 0 Then ApplicationService.GenerateError("El Lote es un dato obligatorio.")
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDDeposito", AddressOf CambioDeposito)
        OBrl.Add("IDArticulo", AddressOf CambioIDArticulo)
        OBrl.Add("IDBarrica", AddressOf CambioBarrica)
        Return OBrl
    End Function

    <Task()> Public Shared Sub CambioDeposito(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GestionaDatosDeposito, data, services)
    End Sub

    <Task()> Public Shared Sub CambioBarrica(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GestionaDatosDeposito, data, services)
    End Sub

    <Task()> Public Shared Sub GestionaDatosDeposito(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDDeposito")) > 0 Then
            Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
            Dim DtpoInfo As BdgDepositoInfo = Depositos.GetEntity(data.Current("IDDeposito"))

            If Not DtpoInfo Is Nothing AndAlso Length(DtpoInfo.IDDeposito) > 0 Then
                data.Current("TipoDeposito") = DtpoInfo.TipoDeposito
                If DtpoInfo.TipoDeposito = TipoDeposito.Barricas AndAlso DtpoInfo.UsarBarricaComoLote Then
                    data.Current("Lote") = data.Current("IDBarrica")
                ElseIf DtpoInfo.TipoDeposito <> TipoDeposito.Barricas Then
                    data.Current("IDBarrica") = DBNull.Value
                End If
            End If
        Else
            data.Current("IDBarrica") = DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDArticulo")) > 0 Then
            ProcessServer.ExecuteTask(Of String)(AddressOf BdgVino.ValidarArticuloVino, data.Current("IDArticulo"), services)
            Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(data.Current("IDArticulo"))
            If dtArticulo.Rows.Count > 0 Then
                data.Current("DescArticulo") = dtArticulo.Rows(0)("DescArticulo")
            End If
        End If
    End Sub

#End Region

End Class

<Serializable()> _
Public Class _EVD
    Public Const NEntrada As String = "NEntrada"
    Public Const IDDeposito As String = "IDDeposito"
    Public Const Cantidad As String = "Cantidad"
    Public Const IDVino As String = "IDVino"
End Class