Public Class ObraParteAgrupadoMod
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupadoMod"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarObraModControl)
    End Sub

    <Task()> Public Shared Sub BorrarObraModControl(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim OMC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraModControl")
        Dim dt As DataTable = OMC.Filter(New GuidFilterItem("IDParteAgrupadoMod", data("IDParteAgrupadoMod")))
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dr("IDParteAgrupadoMod") = DBNull.Value
            Next
            OMC.Delete(dt)
        End If
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDOperario", AddressOf CambioIDOperario)
        Obrl.Add("IDHora", AddressOf CambioHora)
        Obrl.Add("QHoras", AddressOf CalcularImporte)
        Obrl.Add("TasaA", AddressOf CalcularImporte)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioIDOperario(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
            Dim Operario As OperarioInfo = Operarios.GetEntity(data.Value)

            data.Current("DescOperario") = Operario.DescOperario
            data.Current("IDHora") = System.DBNull.Value
            If Length(Operario.IDCategoria) > 0 Then
                data.Current("IDCategoria") = Operario.IDCategoria
                Dim IDHora As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf HoraCategoria.GetHoraPredeterminada, Operario.IDCategoria, services)
                If Length(IDHora) > 0 Then
                    data.Current("IDHora") = IDHora
                    data.Current("DescHora") = New Hora().GetItemRow(IDHora)("DescHora")
                    ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioHora, data, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioHora(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            data.Current("TasaRealModA") = 0
            If Length(data.Current("IDCategoria")) = 0 AndAlso Length(data.Current("IDOperario")) > 0 Then
                Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
                Dim Operario As OperarioInfo = Operarios.GetEntity(data.Current("IDOperario"))
                data.Current("IDCategoria") = Operario.IDCategoria
            End If
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf HoraCategoria.ObtenerPrecioHoraCategoria, data.Current, services)
            data.Current("TasaA") = data.Current("TasaRealModA")
            data.Current("ImporteA") = Nz(data.Current("TasaA"), 0) * Nz(data.Current("QHoras"), 0)
        End If
    End Sub

    <Task()> Public Shared Sub CalcularImporte(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            data.Current("ImporteA") = Nz(data.Current("TasaA"), 0) * Nz(data.Current("QHoras"), 0)
        End If
    End Sub

#End Region

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If (Length(data("IDCategoria")) = 0) Then ApplicationService.GenerateError("La categoría es obligatoria.")
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDParteAgrupadoMod")) = 0 Then data("IDParteAgrupadoMod") = Guid.NewGuid 'AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class dataGrupoOperario
        Public IDParteAgrupado As Guid
        Public IDGrupo As String
        Public IDHora As String
        Public DescHora As String
        Public Cantidad As Double
        Public Factor As enumFactor

        Public Sub New(ByVal IDParteAgrupado As Guid, ByVal IDGrupo As String, ByVal IDHora As String, ByVal DescHora As String, ByVal Cantidad As Double, ByVal Factor As Integer)
            Me.IDParteAgrupado = IDParteAgrupado
            Me.IDGrupo = IDGrupo
            Me.IDHora = IDHora
            Me.DescHora = DescHora
            Me.Cantidad = Cantidad
            Me.Factor = Factor
        End Sub
    End Class
    <Task()> Public Shared Function GeneracionOperariosPartesAgrupados(ByVal data As dataGrupoOperario, ByVal services As ServiceProvider) As DataTable
        Dim dtGrupoOp As DataTable = New GrupoOperario().Filter(New StringFilterItem("IDGrupo", data.IDGrupo))
        If dtGrupoOp.Rows.Count > 0 Then
            Dim m As New ObraParteAgrupadoMod
            Dim dtOperarios As DataTable = New BE.DataEngine().Filter("vBdgMntoObraParteAgrupadoMod", New NoRowsFilterItem)
            Dim Horas As Double = data.Cantidad
            If data.Factor = enumFactor.Divide Then
                Horas = data.Cantidad / dtGrupoOp.Rows.Count
            End If
            For Each drGrupoOp As DataRow In dtGrupoOp.Rows
                Dim drOperario As DataRow = dtOperarios.NewRow
                drOperario("IDParteAgrupado") = data.IDParteAgrupado
                drOperario("IDOperario") = drGrupoOp("IDOperario")
                ' If Length(data.IDHora) = 0 Then
                drOperario = m.ApplyBusinessRule("IDOperario", drOperario("IDOperario"), drOperario, New BusinessData)
                'Else
                Dim Operarios As EntityInfoCache(Of OperarioInfo) = services.GetService(Of EntityInfoCache(Of OperarioInfo))()
                Dim Operario As OperarioInfo = Operarios.GetEntity(drOperario("IDOperario"))

                drOperario("DescOperario") = Operario.DescOperario
                drOperario("DescHora") = data.DescHora
                drOperario("IDHora") = data.IDHora
                If Length(Operario.IDCategoria) > 0 Then
                    drOperario("IDCategoria") = Operario.IDCategoria
                    drOperario("DescCategoria") = Operario.DescCategoria
                    Dim IDHora As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf HoraCategoria.GetHoraPredeterminada, Operario.IDCategoria, services)
                    If Length(IDHora) > 0 Then
                        drOperario("IDHora") = IDHora
                        drOperario("DescHora") = New Hora().GetItemRow(IDHora)("DescHora")
                    End If
                    ' End If
                    drOperario = m.ApplyBusinessRule("IDHora", drOperario("IDHora"), drOperario, New BusinessData)

                End If
                drOperario("QHoras") = Horas
                drOperario = m.ApplyBusinessRule("QHoras", drOperario("QHoras"), drOperario, New BusinessData)

                dtOperarios.Rows.Add(drOperario.ItemArray)
            Next

            Return dtOperarios
        End If

        Return Nothing
    End Function

End Class