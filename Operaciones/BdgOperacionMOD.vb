Public Class BdgOperacionMOD

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionMOD"

#End Region

#Region "Eventos Entidad"

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDOperario", AddressOf BdgGeneral.CambioOperario)
        Obrl.Add("IDHora", AddressOf BdgGeneral.CalculoTasa)
        Obrl.Add("IDCategoria", AddressOf BdgGeneral.CalculoTasa)
        Return Obrl
    End Function

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El Operario es un dato obligatorio.")
        If Length(data("IDHora")) = 0 Then ApplicationService.GenerateError("La Hora es obligatoria.")
        If Length(data("IDCategoria")) = 0 Then ApplicationService.GenerateError("La Categoría es obligatoria.")
    End Sub

#End Region


#End Region

#Region "Tareas Públicas"

    <Serializable()> _
    Public Class dataGrupoOperario
        Public IDGrupo As String
        Public Cantidad As Double
        Public IDHora As String
        Public Factor As enumFactor

        Public Sub New(ByVal IDGrupo As String, ByVal IDHora As String, ByVal Cantidad As Double, ByVal Factor As Integer)
            Me.IDGrupo = IDGrupo
            Me.IDHora = IDHora
            Me.Cantidad = Cantidad
            Me.Factor = Factor
        End Sub
    End Class

    <Task()> Public Shared Function GeneracionOperariosPartesAgrupados(ByVal data As dataGrupoOperario, ByVal services As ServiceProvider) As DataTable
        Dim dtGrupoOp As DataTable = New GrupoOperario().Filter(New StringFilterItem("IDGrupo", data.IDGrupo))
        If dtGrupoOp.Rows.Count > 0 Then
            Dim ClsPlanMOD As New BdgOperacionMOD
            Dim dtOperarios As DataTable = New BE.DataEngine().Filter("frmBdgOperacionMODRealGlobal", New NoRowsFilterItem)
            Dim Horas As Double = data.Cantidad
            If data.Factor = enumFactor.Divide Then
                Horas = data.Cantidad / dtGrupoOp.Rows.Count
            End If
            For Each drGrupoOp As DataRow In dtGrupoOp.Rows
                Dim drOperario As DataRow = dtOperarios.NewRow
                drOperario("IDOperario") = drGrupoOp("IDOperario")
                drOperario("IDHora") = data.IDHora
                drOperario("Fecha") = Today.Date
                drOperario = ClsPlanMOD.ApplyBusinessRule("IDOperario", drOperario("IDOperario"), drOperario, New BusinessData)
                drOperario("Tiempo") = Horas
                dtOperarios.Rows.Add(drOperario.ItemArray)
            Next
            Return dtOperarios
        End If
        Return Nothing
    End Function

#End Region

End Class