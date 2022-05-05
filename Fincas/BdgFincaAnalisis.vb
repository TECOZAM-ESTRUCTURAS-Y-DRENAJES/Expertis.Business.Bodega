Public Class BdgFincaAnalisis

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFincaAnalisis"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DataCont As New Contador.DatosDefaultCounterValue(data, "BdgFincaAnalisis", _VA.NVinoAnalisis)
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, DataCont, services)
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdContador")) > 0 Then
                data("NFincaAnalisis") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StGetGAnalisis
        Public IDAnalisis As String
        Public Filtro As Filter

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDAnalisis As String, ByVal Filtro As Filter)
            Me.IDAnalisis = IDAnalisis
            Me.Filtro = Filtro
        End Sub
    End Class

    <Task()> Public Shared Function getAnalisis(ByVal data As StGetGAnalisis, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New DataEngine().Filter("frmBdgCIAnalisisFinca", data.Filtro)
        Dim htColumas As Hashtable
        Dim strSelect As String = String.Empty

        'ahora traemos todas las columnas correspondientes
        Dim dtVariables As DataTable = New BdgAnalisisVariable().Filter(New StringFilterItem("IDAnalisis", data.IDAnalisis))
        Dim fVariables As New Filter(FilterUnionOperator.Or)
        For Each dcV As DataRow In dtVariables.Select(Nothing, "Orden")
            If Not dt.Columns.Contains(dcV("IDVariable")) Then
                dt.Columns.Add(dcV("IDVariable"), GetType(String))
                dt.Columns.Add(dcV("IDVariable") & "_N", GetType(Double))
                fVariables.Add("IDVariable", dcV("IDVariable"))
                strSelect = strSelect + dcV("IDVariable") + ","
            End If
        Next
        strSelect += "0"

        'ahora las actualizamos
        Dim bsnBFA As New BdgFincaVariable
        For Each dr As DataRow In dt.Select
            Dim f As New Filter
            f.Add("IDFinca", dr("IDFinca"))
            f.Add(fVariables)
            Dim dtResult As DataTable = bsnBFA.Filter(f)
            For Each drResult As DataRow In dtResult.Select("Fecha = '" & DateValue(dr("Fecha")) & "'")
                dr(drResult("IDVariable")) = drResult("Valor")
                dr(drResult("IDVariable") & "_N") = drResult("ValorNumerico")
            Next
        Next
        Return dt
    End Function

#End Region

End Class