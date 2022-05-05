Public Class ObraParteAgrupadoLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupadoLinea"

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDParteAgrupadoLinea")) = 0 Then data("IDParteAgrupadoLinea") = Guid.NewGuid 'AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region " GetBusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDFinca", AddressOf CambioFinca)
        Obrl.Add("CFinca", AddressOf CambioFinca)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioFinca(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("CFinca") = DBNull.Value
        data.Current("DescFinca") = DBNull.Value
        data.Current("Superficie") = DBNull.Value
        data.Current("Cantidad") = DBNull.Value
        data.Current("IDFincaPadre") = DBNull.Value
        data.Current("CFincaPadre") = DBNull.Value
        data.Current("DescFincaPadre") = DBNull.Value

        If Length(data.Value) > 0 Then
            data.Current(data.ColumnName) = data.Value
            Dim dt As DataTable = New BE.DataEngine().Filter("vBdgNegDatosFinca", New GuidFilterItem("IDFinca", data.Current("IDFinca")))
            If dt.Rows.Count > 0 Then
                data.Current("CFinca") = dt.Rows(0)("CFinca")
                data.Current("DescFinca") = dt.Rows(0)("DescFinca")
                data.Current("Superficie") = dt.Rows(0)("SuperficieViñedo")
                data.Current("Cantidad") = dt.Rows(0)("SuperficieViñedo")
                data.Current("IDFincaPadre") = dt.Rows(0)("IDFincaPadre")
                data.Current("CFincaPadre") = dt.Rows(0)("CFincaPadre")
                data.Current("DescFincaPadre") = dt.Rows(0)("DescFincaPadre")
            End If
        End If
    End Sub

#End Region

End Class