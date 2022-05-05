Public Class BdgAnalisisLineaValor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgAnalisisLineaValor"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatos)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_ALV.IDAnalisisCabecera)) = 0 Then
            ApplicationService.GenerateError("No se ha indicado la cabecera del Análisis.")
        Else
            Dim dt As DataTable = New BdgAnalisisCabecera().SelOnPrimaryKey(data(_ALV.IDAnalisisCabecera))
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("No se ha encontrado la cabecera del Análisis indicado en la Base de Datos.")
            End If
        End If

        If Length(data(_ALV.IDVariable)) = 0 Then
            ApplicationService.GenerateError("No se ha indicado la Variable.")
        Else
            Dim dt As DataTable = New BdgVariable().SelOnPrimaryKey(data(_ALV.IDVariable))
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("No se ha encontrado la Variable indicada en la Base de Datos.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data(_ALV.ValorNumerico)) <> 0 AndAlso Not IsNumeric(data(_ALV.ValorNumerico)) Then
            ApplicationService.GenerateError("El valor del campo Valor Numérico no es un número.")
        ElseIf Length(data(_ALV.Orden)) <> 0 AndAlso Not IsNumeric(data(_ALV.Orden)) Then
            ApplicationService.GenerateError("El valor del campo Orden no es un número.")
        End If
    End Sub


    'Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
    '    MyBase.RegisterUpdateTasks(updateProcess)
    '    'updateProcess.AddTask(Of DataRow)(AddressOf Comunes.UpdateEntityRow)
    '    'updateProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoActualizado)
    '    'updateProcess.AddTask(Of DataRow)(AddressOf ActualizarOrden)
    'End Sub

    '<Task()> Public Shared Sub ActualizarOrden(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    '    Dim dtALV As DataTable = New BdgAnalisisLineaValor().Filter(New GuidFilterItem(_ALV.IDAnalisisCabecera, FilterOperator.Equal, data(_ALV.IDAnalisisCabecera)), _ALV.Orden)
    '    '    If Not dtALV Is Nothing AndAlso dtALV.Rows.Count > 0 Then
    '    '        Dim i As Integer = 1
    '    '        For Each DrVar As DataRow In dtALV.Select
    '    '            DrVar("Orden") = i
    '    '            i += 1
    '    '        Next
    '    '        BusinessHelper.UpdateTable(dtALV)
    '    '    End If
    'End Sub

#End Region

#Region "Funciones Públicas"


#End Region

End Class

<Serializable()> _
Public Class _ALV
    Public Const Entidad As String = "BdgAnalisisLineaValor"
    Public Const IDAnalisisCabecera As String = "IDAnalisisCabecera"
    Public Const IDVariable As String = "IDVariable"
    Public Const Valor As String = "Valor"
    Public Const ValorNumerico As String = "ValorNumerico"
    Public Const Orden As String = "Orden"
End Class