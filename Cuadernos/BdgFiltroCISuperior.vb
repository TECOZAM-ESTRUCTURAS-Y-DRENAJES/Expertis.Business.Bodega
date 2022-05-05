Public Class BdgFiltroCISuperior
    Inherits BusinessHelper

    <Serializable()> _
    Public Class FiltroSuperiorInfo
        Public IDFiltro As Integer
        Public IDPrograma As Guid
        Public DescFiltro As String
        Public Orden As Integer
        Public WhereAvanzada As String
        Public XMLWhereGeneral As String
        Public Filtro1 As ConfigFiltroSuperior
        Public Filtro2 As ConfigFiltroSuperior
        Public Filtro3 As ConfigFiltroSuperior
        Public Filtro4 As ConfigFiltroSuperior
    End Class

    <Serializable()> _
    Public Class ConfigFiltroSuperior
        Public strWhere As String
        Public blnIN As Boolean
        Public strOperador As Integer
    End Class

    Private Const cnEntidad As String = "tbFiltroCISuperior"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Public Sub DeleteFiltro(ByVal IDFiltro As Integer)
        Dim dtF As DataTable = Me.SelOnPrimaryKey(IDFiltro)
        If Not dtF Is Nothing AndAlso dtF.Rows.Count > 0 Then
            Me.Delete(dtF)
        End If
    End Sub

    Public Function GuardarFiltro(ByVal FiltroInfo As FiltroSuperiorInfo) As String
        Dim dt As DataTable
        If FiltroInfo.IDFiltro = Nothing OrElse FiltroInfo.IDFiltro = -1 Then
            dt = Me.AddNewForm
            dt.Rows(0)("IDFiltro") = AdminData.GetAutoNumeric
            dt.Rows(0)("IDPrograma") = FiltroInfo.IDPrograma
            dt.Rows(0)("DescFiltro") = FiltroInfo.DescFiltro
            Dim dtOrden As DataTable = Me.Filter(New GuidFilterItem("IDPrograma", FiltroInfo.IDPrograma))
            If dtOrden.Rows.Count > 0 Then dt.Rows(0)("Orden") = dtOrden.Rows(dtOrden.Rows.Count - 1)("Orden") + 10 Else dt.Rows(0)("Orden") = 10
        Else
            dt = Me.SelOnPrimaryKey(FiltroInfo.IDFiltro)
        End If
        dt.Rows(0)("WhereAvanzada") = FiltroInfo.WhereAvanzada

        If Not FiltroInfo.Filtro1 Is Nothing Then
            dt.Rows(0)("Where1") = FiltroInfo.Filtro1.strWhere
            dt.Rows(0)("IN1") = FiltroInfo.Filtro1.blnIN
            dt.Rows(0)("Operador1") = FiltroInfo.Filtro1.strOperador
        End If
        If Not FiltroInfo.Filtro2 Is Nothing Then
            dt.Rows(0)("Where2") = FiltroInfo.Filtro2.strWhere
            dt.Rows(0)("IN2") = FiltroInfo.Filtro2.blnIN
            dt.Rows(0)("Operador2") = FiltroInfo.Filtro2.strOperador
        End If
        If Not FiltroInfo.Filtro3 Is Nothing Then
            dt.Rows(0)("Where3") = FiltroInfo.Filtro3.strWhere
            dt.Rows(0)("Operador3") = FiltroInfo.Filtro3.strOperador
            dt.Rows(0)("IN3") = FiltroInfo.Filtro3.blnIN
        End If
        If Not FiltroInfo.Filtro4 Is Nothing Then
            dt.Rows(0)("Where4") = FiltroInfo.Filtro4.strWhere
            dt.Rows(0)("Operador4") = FiltroInfo.Filtro4.strOperador
            dt.Rows(0)("IN4") = FiltroInfo.Filtro4.blnIN
        End If

        MyBase.Update(dt)
        Return dt.Rows(0)("IDFiltro")
    End Function

    Public Function CargarFiltros(ByVal strPrograma As Guid) As FiltroSuperiorInfo()
        Dim dtFiltros As DataTable = Me.Filter(New GuidFilterItem("IDPrograma", strPrograma))
        Dim FiltrosInfo(-1) As FiltroSuperiorInfo
        Dim i As Integer = 0
        For Each dr As DataRow In dtFiltros.Select(Nothing, "orden")
            Dim filtroInfo As New FiltroSuperiorInfo
            filtroInfo.IDFiltro = dr("IDFiltro")
            filtroInfo.IDPrograma = dr("IDPrograma")
            filtroInfo.DescFiltro = dr("DescFiltro")
            filtroInfo.Orden = dr("Orden")
            filtroInfo.WhereAvanzada = dr("WhereAvanzada") & String.Empty

            If Length(dr("Where1")) > 0 Then
                Dim c1 As New ConfigFiltroSuperior
                c1.strWhere = dr("Where1")
                c1.blnIN = dr("IN1")
                c1.strOperador = dr("Operador1")
                filtroInfo.Filtro1 = c1
            End If

            If Length(dr("Where2")) > 0 Then
                Dim c2 As New ConfigFiltroSuperior
                c2.strWhere = dr("Where2")
                c2.blnIN = dr("IN2")
                c2.strOperador = dr("Operador2")
                filtroInfo.Filtro2 = c2
            End If

            If Length(dr("Where3")) > 0 Then
                Dim c3 As New ConfigFiltroSuperior
                c3.strWhere = dr("Where3")
                c3.blnIN = dr("IN3")
                c3.strOperador = dr("Operador3")
                filtroInfo.Filtro3 = c3
            End If

            If Length(dr("Where4")) > 0 Then
                Dim c4 As New ConfigFiltroSuperior
                c4.strWhere = dr("Where4")
                c4.blnIN = dr("IN4")
                c4.strOperador = dr("Operador4")
                filtroInfo.Filtro4 = c4
            End If

            ReDim Preserve FiltrosInfo(UBound(FiltrosInfo) + 1)
            FiltrosInfo(UBound(FiltrosInfo)) = filtroInfo
            i += 1
        Next

        Return FiltrosInfo
    End Function

End Class
