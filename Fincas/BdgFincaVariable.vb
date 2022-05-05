Public Class BdgFincaVariable

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFincaVariable"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf BorradoVariable)
    End Sub

    <Task()> Public Shared Sub BorradoVariable(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New BdgFincaVariable().SelOnPrimaryKey(data(_FV.IDFinca), data(_FV.Fecha), data(_FV.IDVariable))
            If dt.Rows.Count > 0 Then
                dt.Rows(0).Delete()
                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StValoresEnFecha
        Public IDVariable As String
        Public Fecha As Date

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVariable As String, ByVal Fecha As Date)
            Me.IDVariable = IDVariable
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function ValoresEnFecha(ByVal data As StValoresEnFecha, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Execute("spBdgValorVariableFinca", False, data.IDVariable, data.Fecha)
    End Function

#End Region

End Class

<Serializable()> _
Public Class _FV
    Public Const IDFinca As String = "IDFinca"
    Public Const Fecha As String = "Fecha"
    Public Const IDVariable As String = "IDVariable"
    Public Const IDAnalisis As String = "IDAnalisis"
    Public Const Valor As String = "Valor"
    Public Const ValorNumerico As String = "ValorNumerico"
End Class