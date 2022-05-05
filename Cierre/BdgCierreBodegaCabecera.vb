Public Class BdgCierreBodegaCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgCierreBodegaCabecera"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub DeleteCierre(ByVal Fecha As Date, ByVal services As ServiceProvider)
        Dim ClsCierre As New BdgCierreBodegaCabecera
        Dim DtCabecera As DataTable = ClsCierre.Filter(New FilterItem("FechaCierreBodega", FilterOperator.Equal, Fecha, FilterType.DateTime))
        If Not DtCabecera Is Nothing AndAlso DtCabecera.Rows.Count > 0 Then ClsCierre.Delete(DtCabecera)
    End Sub

    <Serializable()> _
    Public Class StCierreBodega
        Public Fecha As Date
        Public DescCierre As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Fecha As Date, ByVal DescCierre As String)
            Me.Fecha = Fecha
            Me.DescCierre = DescCierre
        End Sub
    End Class

    <Task()> Public Shared Sub CierreBodega(ByVal data As StCierreBodega, ByVal services As ServiceProvider)
        If Length(data.Fecha) > 0 AndAlso Length(data.DescCierre) > 0 Then
            Dim ClsCierre As New BdgCierreBodegaCabecera
            Dim dtCabecera As DataTable = ClsCierre.Filter(New FilterItem("FechaCierreBodega", FilterOperator.Equal, data.Fecha, FilterType.DateTime))
            If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 Then
                ClsCierre.Delete(dtCabecera)
            End If
            dtCabecera = ProcessServer.ExecuteTask(Of StCierreBodega, DataTable)(AddressOf NuevaCabeceraCierre, data, services)
            If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 Then
                Dim dtLineas As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf NuevaLineaCierre, dtCabecera.Rows(0)("IDCierreBodega"), services)
                BusinessHelper.UpdateTable(dtCabecera)
                BusinessHelper.UpdateTable(dtLineas)
            End If
        Else : ApplicationService.GenerateError("La Fecha Cierre y su descripción son datos necesarios para el Cierre.")
        End If
    End Sub

    <Task()> Public Shared Function NuevaCabeceraCierre(ByVal data As StCierreBodega, ByVal services As ServiceProvider) As DataTable
        Dim dtCabecera As DataTable = New BdgCierreBodegaCabecera().AddNewForm
        dtCabecera.Rows(0)("IDCierreBodega") = Guid.NewGuid
        dtCabecera.Rows(0)("FechaCierreBodega") = data.Fecha
        dtCabecera.Rows(0)("DescCierreBodega") = data.DescCierre
        Return dtCabecera
    End Function

    <Task()> Public Shared Function NuevaLineaCierre(ByVal IDCierreBodega As Guid, ByVal services As ServiceProvider) As DataTable
        Dim Dt As DataTable = New BE.DataEngine().Filter("vNegBdgDepositosConVino", "*", "")
        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
            Dim dtLineas As DataTable = New BdgCierreBodegaLinea().AddNew
            For Each dr As DataRow In Dt.Select
                Dim drLinea As DataRow = dtLineas.NewRow
                For Each dc As DataColumn In dtLineas.Columns
                    If dc.ColumnName <> "IDCierreBodegaLinea" And dc.ColumnName <> "IDCierreBodega" Then
                        drLinea(dc.ColumnName) = dr(dc.ColumnName)
                    End If
                Next
                drLinea("IDCierreBodegaLinea") = Guid.NewGuid
                drLinea("IDCierreBodega") = IDCierreBodega
                dtLineas.Rows.Add(drLinea.ItemArray)
            Next
            Return dtLineas
        End If
    End Function

#End Region

End Class