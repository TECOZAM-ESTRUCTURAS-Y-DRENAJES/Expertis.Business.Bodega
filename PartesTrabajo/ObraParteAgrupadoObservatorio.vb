Public Class ObraParteAgrupadoObservatorio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbObraParteAgrupadoObservatorio"

#End Region

#Region "Tareas Públicas"

    <Serializable()> _
    Public Class DataInsFincaObs
        Public IDObservatorio As String
        Public IDParteAgrupado As Guid
        Public DtPartesAgrup As DataTable

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDObservatorio As String, ByVal IDParteAgrupado As Guid, ByVal DtPartesAgrup As DataTable)
            Me.IDObservatorio = IDObservatorio
            Me.IDParteAgrupado = IDParteAgrupado
            Me.DtPartesAgrup = DtPartesAgrup
        End Sub
    End Class

    <Task()> Public Shared Function InsertarFincasObservatorio(ByVal data As DataInsFincaObs, ByVal services As ServiceProvider) As DataTable
        Dim DtBdgFincaObs As DataTable = New BdgObservatorioFinca().Filter(New FilterItem("IDObservatorio", FilterOperator.Equal, data.IDObservatorio))
        If Not DtBdgFincaObs Is Nothing AndAlso DtBdgFincaObs.Rows.Count > 0 Then
            Dim ClsPartesAgrup As New ObraParteAgrupadoLinea
            For Each DrFinca As DataRow In DtBdgFincaObs.Select
                If Not data.DtPartesAgrup Is Nothing AndAlso data.DtPartesAgrup.Rows.Count > 0 Then
                    Dim DrFind() As DataRow = data.DtPartesAgrup.Select("IDFinca = '" & CType(DrFinca("IDFinca"), Guid).ToString & "'")
                    If DrFind.Length <= 0 Then
                        Dim DrNew As DataRow = data.DtPartesAgrup.NewRow
                        DrNew = ClsPartesAgrup.ApplyBusinessRule("IDFinca", DrFinca("IDFinca"), DrNew)
                        DrNew("IDParteAgrupado") = data.IDParteAgrupado
                        DrNew("Marca") = False
                        data.DtPartesAgrup.Rows.Add(DrNew)
                    End If
                Else
                    Dim DrNew As DataRow = data.DtPartesAgrup.NewRow
                    DrNew = ClsPartesAgrup.ApplyBusinessRule("IDFinca", DrFinca("IDFinca"), DrNew)
                    DrNew("IDParteAgrupado") = data.IDParteAgrupado
                    DrNew("Marca") = False
                    data.DtPartesAgrup.Rows.Add(DrNew)
                End If
            Next
        End If
        Return data.DtPartesAgrup
    End Function

#End Region

End Class