Public Class BdgFiltroCIGeneral
    Inherits BusinessHelper

    <Serializable()> _
    Public Class FiltroGeneralInfo
        Public IDFiltro As Integer
        Public Campo As String
        Public Operador As String
        Public Valor As String
        Public Tipo As String
        Public PK As Integer
    End Class

    Private Const cnEntidad As String = "tbFiltroCIGeneral"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Sub DeleteFiltro(ByVal IDFiltro As Integer)
        Dim f As New Filter
        f.Add("IDFiltro", IDFiltro)
        Dim dtF As DataTable = Me.Filter(f)
        If Not dtF Is Nothing AndAlso dtF.Rows.Count > 0 Then
            Me.Delete(dtF)
        End If
    End Sub

    Public Sub GuardarFiltro(ByVal FiltrosInfo() As FiltroGeneralInfo, ByVal IDFiltro As Integer)
        Dim dt As DataTable = Me.AddNew
        DeleteFiltro(IDFiltro)
        For Each filtroinfo As BdgFiltroCIGeneral.FiltroGeneralInfo In FiltrosInfo
            Dim dr As DataRow = dt.NewRow
            dr("IDFiltroCampo") = AdminData.GetAutoNumeric
            dr("IDFiltro") = IDFiltro
            dr("Campo") = filtroinfo.Campo
            dr("Valor") = filtroinfo.Valor
            dt.Rows.Add(dr)
        Next

        MyBase.Update(dt)
    End Sub

    Public Function CargarFiltros(ByVal IDFiltro As Integer) As FiltroGeneralInfo()
        Dim dtFiltros As DataTable = Me.Filter(New NumberFilterItem("IDFiltro", IDFiltro))
        Dim FiltrosInfo(-1) As FiltroGeneralInfo
        Dim i As Integer = 0
        For Each dr As DataRow In dtFiltros.Rows
            Dim filtroInfo As New FiltroGeneralInfo
            filtroInfo.IDFiltro = dr("IDFiltro")
            filtroInfo.Campo = dr("Campo")
            filtroInfo.Valor = dr("Valor")
            filtroInfo.PK = dr("IDFiltroCampo")
            ReDim Preserve FiltrosInfo(UBound(FiltrosInfo) + 1)
            FiltrosInfo(UBound(FiltrosInfo)) = filtroInfo
            i += 1
        Next
        Return FiltrosInfo
    End Function

End Class
