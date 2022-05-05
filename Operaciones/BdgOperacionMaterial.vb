Public Class BdgOperacionMaterial

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgOperacionMaterial"

#End Region

#Region "Eventos Entidad"

#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarLotes)
    End Sub

    <Task()> Public Shared Sub EliminarLotes(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsMatLotes As New BdgOperacionMaterialLote
        Dim DtLotes As DataTable = ClsMatLotes.Filter(New FilterItem("IDOperacionMaterial", data("IDOperacionMaterial")))
        ClsMatLotes.Delete(DtLotes)
    End Sub

#End Region


#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDArticulo", AddressOf BdgGeneral.CambioMaterialGlobal)
        Obrl.Add("Cantidad", AddressOf BdgGeneral.CambioCantidadGlobal)
        Obrl.Add("Merma", AddressOf BdgGeneral.CambioMermaGlobal)
        Obrl.Add("IDAlmacen", AddressOf BdgGeneral.CambioAlmacenGlobal)
        Return Obrl
    End Function

#End Region


#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

#End Region


#End Region

End Class