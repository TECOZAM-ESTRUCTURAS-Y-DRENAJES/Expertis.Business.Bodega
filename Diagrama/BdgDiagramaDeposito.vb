Imports Solmicro.Expertis.Business.Bodega.BdgEntradaDeposito

Public Enum enumBdgVisualizacionDiagrama
    Interior
    Exterior
End Enum

Public Class BdgDiagramaDeposito

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgDiagramaDeposito"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDDiagrama")) = 0 Then ApplicationService.GenerateError("El identificador de diagrama no es válido.")
        If Length(data("IDDeposito")) = 0 Then ApplicationService.GenerateError("El identificador de depósito no es válido.")
        If Length(data("X")) = 0 OrElse data("X") = 0 Then ApplicationService.GenerateError("La coordenada X no es válida.")
        If Length(data("Y")) = 0 OrElse data("Y") = 0 Then ApplicationService.GenerateError("La coordenada Y no es válida.")
        If Length(data("Ancho")) = 0 OrElse data("Ancho") = 0 Then ApplicationService.GenerateError("El ancho especificado no es válido.")
        If Length(data("Alto")) = 0 OrElse data("Alto") = 0 Then ApplicationService.GenerateError("La altura especificada no es válida.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New BdgDiagramaDeposito().SelOnPrimaryKey(data("IDDiagrama"), data("IDDeposito"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError("El depósito | ya está en el diagrama |.", Quoted(data("IDDeposito")), Quoted(data("IDDiagrama")))
        End If
    End Sub

#End Region

#Region "Public Task"
    <Serializable()> _
    Public Class stAsignarEntradaADeposito
        Public IDEntrada As String
        Public IDDeposito As String
    End Class

    <Serializable()> _
    Public Class stAsignarEntradaADepositoResult
        Public Realizado As Boolean
        Public Mensaje As String
    End Class

    <Task()> _
    Public Shared Function AsignarEntradaADeposito(ByVal data As stAsignarEntradaADeposito, ByVal services As ServiceProvider) As stAsignarEntradaADepositoResult
        Dim result As New stAsignarEntradaADepositoResult()

        Try
            Dim bsnEntrada As New Business.Bodega.BdgEntrada
            Dim dtrEntrada As DataRow = bsnEntrada.GetItemRow(data.IDEntrada)

            Dim bsnDeposito As New Business.Bodega.BdgDeposito
            Dim dtrDeposito As DataRow = bsnDeposito.GetItemRow(data.IDDeposito)

            Dim ocupacion As Double = ProcessServer.ExecuteTask(Of String, Double)(AddressOf Business.Bodega.BdgOperacion.DevolverOcupacion, data.IDDeposito, services)

            Dim dblTotalAsignacion As Double = dtrDeposito("Capacidad") - ocupacion
            If dtrEntrada("Neto") < dblTotalAsignacion Then
                dblTotalAsignacion = dtrEntrada("Neto")
            End If

            If (Length(dtrEntrada("IDVariedad")) = 0) Then
                'ApplicationService.GenerateError("No se ha podido completar la asignación de la entrada. La variedad indicada no es válida.")
                Throw New Exception("No se ha podido completar la asignación de la entrada. La variedad indicada no es válida.")
            End If

            Dim stData As New StCrearEntradaDeposito(dtrEntrada("IDEntrada"), data.IDDeposito, dtrEntrada("Neto"), dtrEntrada("IDVariedad"), dtrEntrada("Vendimia"))
            Dim dttResult As DataTable = ProcessServer.ExecuteTask(Of StCrearEntradaDeposito, DataTable)(AddressOf BdgEntradaDeposito.CrearEntradaDeposito, stData, services)
            Dim ClsEntrada As New BdgEntrada
            Dim UpdtPckg As New UpdatePackage
            Dim DtEntrada As DataTable = ClsEntrada.SelOnPrimaryKey(dtrEntrada("IDEntrada"))
            Dim DtEntDepositos As DataTable = New BdgEntradaDeposito().AddNew
            For Each DrResult As DataRow In dttResult.Select
                DtEntDepositos.Rows.Add(DrResult.ItemArray)
            Next
            UpdtPckg.Add(DtEntrada)
            UpdtPckg.Add(DtEntDepositos)
            ClsEntrada.Update(UpdtPckg)
            'Dim bsnEntradaDeposito As New Business.Bodega.BdgEntradaDeposito
            'bsnEntradaDeposito.Update(dttResult)

            result.Realizado = True
            Return result
        Catch ex As Exception
            result.Realizado = False
            result.Mensaje = ex.Message
            Return result
        End Try

    End Function

#End Region

End Class