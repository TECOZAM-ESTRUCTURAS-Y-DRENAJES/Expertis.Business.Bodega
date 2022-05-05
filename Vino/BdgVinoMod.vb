Public Class BdgVinoMod

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgVinoMOD"

#End Region

#Region "Eventos Entidad"

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim OBrl As New BusinessRules
        OBrl.Add("IDOperario", AddressOf BdgGeneral.CambioOperario)
        OBrl.Add("IDHora", AddressOf BdgGeneral.CalculoTasa)
        OBrl.Add("IDCategoria", AddressOf BdgGeneral.CalculoTasa)
        Return OBrl
    End Function

#End Region

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDOperario")) = 0 Then ApplicationService.GenerateError("El Operario es obligatorio.")
        If Length(data("IDHora")) = 0 Then ApplicationService.GenerateError("La Hora es obligatoria.")
        If Length(data("IDCategoria")) = 0 Then ApplicationService.GenerateError("La Categoría es obligatoria.")
    End Sub

#End Region

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StCrearVinoMOD
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

    <Task()> Public Shared Function CrearVinoMOD(ByVal data As StCrearVinoMOD, ByVal services As ServiceProvider)
        If Not data.Data Is Nothing AndAlso data.Data.Rows.Count > 0 Then
            Dim dtVMD As DataTable = New BdgVinoMod().AddNew
            For Each oRw As DataRow In data.Data.Rows
                Dim rwVMD As DataRow = dtVMD.NewRow
                rwVMD(_VMD.IDVino) = oRw(_VMD.IDVino)
                rwVMD(_VMD.IDOperario) = oRw(_VMD.IDOperario)
                rwVMD(_VMD.DescOperacion) = oRw(_VMD.DescOperacion)
                rwVMD(_VMD.Fecha) = data.Fecha
                rwVMD(_VMD.NOperacion) = data.NOperacion
                rwVMD(_VMD.Tiempo) = oRw(_VMD.Tiempo)
                rwVMD(_VMD.Tasa) = oRw(_VMD.Tasa)
                '''''''''''''''''''''''''''''''''''


                If Length(oRw("IDHora")) = 0 Then
                    Dim HoraCategoria As New HoraCategoria
                    Dim mstrHoraPred As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf HoraCategoria.GetHoraPredeterminada, oRw("IDCategoria"), services)

                    rwVMD("IDHora") = mstrHoraPred
                    Dim dataHC As New General.HoraCategoria.DatosPrecioHoraCatOper(oRw("IDCategoria"), rwVMD("IDHora"), data.Fecha, oRw("IDOperario"))
                    oRw("Tasa") = ProcessServer.ExecuteTask(Of General.HoraCategoria.DatosPrecioHoraCatOper, Double)(AddressOf General.HoraCategoria.ObtenerPrecioHoraCategoriaOperario, dataHC, services)
                Else
                    rwVMD("IDHora") = oRw("IDHora")
                End If

                rwVMD("IDCategoria") = oRw("IDCategoria")
                rwVMD(_VMD.IDOperacionMOD) = oRw(_VMD.IDOperacionMOD)
                '''''''''''''''''''''''''''''''''''
                dtVMD.Rows.Add(rwVMD)
            Next
            BusinessHelper.UpdateTable(dtVMD)
        End If
    End Function

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
            Dim ClsPlanMOD As New BdgVinoMod
            Dim dtOperarios As DataTable = New BE.DataEngine().Filter("frmBdgOperacionMod", New NoRowsFilterItem)
            Dim Horas As Double = data.Cantidad
            If data.Factor = enumFactor.Divide Then
                Horas = data.Cantidad / dtGrupoOp.Rows.Count
            End If
            For Each drGrupoOp As DataRow In dtGrupoOp.Rows
                Dim drOperario As DataRow = dtOperarios.NewRow
                drOperario("IDOperario") = drGrupoOp("IDOperario")
                drOperario("IDHora") = data.IDHora
                Dim operarioBusiness As BusinessHelper = BusinessHelper.CreateBusinessObject("Operario")
                Dim dttOperario As DataTable = operarioBusiness.SelOnPrimaryKey(drOperario("IDOperario"))
                If (Not dttOperario Is Nothing AndAlso dttOperario.Rows.Count > 0) Then
                    Dim operario As OperarioInfo = New OperarioInfo(dttOperario.Rows(0))
                    drOperario("IDCategoria") = operario.IDCategoria
                    drOperario("DescOperario") = operario.DescOperario
                End If
                Dim dataHora As New General.HoraCategoria.DatosPrecioHoraCatOper(drOperario("IDCategoria"), drOperario("IDHora"), Today.Date, drOperario("IDOperario"))
                drOperario("Tasa") = ProcessServer.ExecuteTask(Of General.HoraCategoria.DatosPrecioHoraCatOper, Double)(AddressOf General.HoraCategoria.ObtenerPrecioHoraCategoriaOperario, dataHora, services)
                drOperario("Tiempo") = Horas
                dtOperarios.Rows.Add(drOperario.ItemArray)
            Next
            Return dtOperarios
        End If
        Return Nothing
    End Function

#End Region

End Class

<Serializable()> _
Public Class _VMD
    Public Const IDVino As String = "IDVino"
    Public Const IDOperario As String = "IDOperario"
    Public Const Fecha As String = "Fecha"
    Public Const DescOperacion As String = "DescOperacion" '¿Se necesita?
    Public Const NOperacion As String = "NOperacion"
    Public Const Tiempo As String = "Tiempo"
    Public Const Tasa As String = "Tasa"
    Public Const IDCategoria As String = "IDCategoria"
    Public Const IDOperacionMOD As String = "IDOperacionMOD"
End Class