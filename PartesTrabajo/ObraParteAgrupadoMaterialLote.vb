Public Class ObraParteAgrupadoMaterialLote
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupadoMaterialLote"


#Region " RegisterDeleteTasks "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarObraMaterialLote)
    End Sub

    <Task()> Public Shared Sub BorrarObraMaterialLote(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add("IDParteAgrupadoMatLote", data("IDParteAgrupadoMatLote"))
        Dim dttResult As DataTable = New ObraMaterialControlLote().Filter(f)
        If Not dttResult Is Nothing AndAlso dttResult.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede eliminar el desglose porque ya existen movimientos asociados." & vbCrLf & "Borre la imputación del material para deshacer los movimientos.")
        End If
    End Sub

#End Region

End Class
