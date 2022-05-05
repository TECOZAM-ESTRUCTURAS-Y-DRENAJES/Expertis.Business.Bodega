Public Class BdgEntradaFacturacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgEntradaFacturacion"

#End Region

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("Neto", "Cantidad")
        oBRL.Add("Porcentaje", AddressOf BdgProcesoEntrada.CambioPorcentaje)
        oBRL.Add("Cantidad", AddressOf BdgProcesoEntrada.CambioCantidad)
        Return oBRL
    End Function

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarFactura)
    End Sub

    <Task()> Public Shared Sub ComprobarFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("IDLineaFactura") <> 0 Then
            ApplicationService.GenerateError("No se puede borrar la Entrada-Facturación, existen facturas relacionadas.")
        End If
    End Sub

End Class

Public Enum enumTipoFacturacion
    Declarado
    Excedente
    Papel
    UvaSinPapel
End Enum
