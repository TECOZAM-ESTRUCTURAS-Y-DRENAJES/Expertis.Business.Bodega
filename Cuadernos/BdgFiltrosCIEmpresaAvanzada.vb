<Transactional()> _
Public Class BdgFiltrosCIEmpresaAvanzada
    Inherits ContextBoundObject
    <Serializable()> _
    Public Class SuperAdvInfo
        Public strInNotIn As String
        Public strEntidad As String
        Public strCampo As String
        Public strCriterio As String
        Public strAndOr As String
    End Class

    Public Function ConstruirFiltroEspecial(ByVal WhereGeneral As String, ByVal WhereAvanzada As String) As String
        Dim selectSQL As New Text.StringBuilder
        If Length(WhereGeneral) > 0 Or Length(WhereAvanzada) > 0 Then
            selectSQL.Append(" WHERE ")
            If Length(WhereGeneral) > 0 Then
                selectSQL.Append(WhereGeneral)
                If Length(WhereAvanzada) > 0 Then
                    selectSQL.Append(" AND ")
                    selectSQL.Append("(")
                End If
            End If
            If Length(WhereAvanzada) > 0 Then
                selectSQL.Append(WhereAvanzada)
                If Length(WhereGeneral) > 0 Then selectSQL.Append(")")
            End If
        End If
        Return selectSQL.ToString
    End Function

    Public Function ConstruirWhereGeneral(ByVal f As Filter) As String
        Dim whereSQL As New Text.StringBuilder
        If f.Count > 0 Then
            whereSQL.Append(AdminData.ComposeFilter(f))
        End If
        Return whereSQL.ToString
    End Function

    Public Function ConstruirWhereAvanzada(ByVal advInfo() As SuperAdvInfo, ByVal Entidad As String) As String
        Dim whereSQL As New Text.StringBuilder
        Dim blnCerrarParentesis As Boolean = False
        If Not advInfo Is Nothing AndAlso advInfo.Length > 0 Then
            Dim i As Integer
            Dim strCampo As String
            For i = 0 To advInfo.Length - 1
                'If Entidad = "BdgFinca" Then
                '    strCampo = "CFinca"
                'Else
                strCampo = advInfo(i).strCampo
                'End If
                Dim strCriterio As String = advInfo(i).strCriterio.Replace("*", "%")
                whereSQL.Append(strCampo & " " & advInfo(i).strInNotIn & " (SELECT " & strCampo & " FROM " & advInfo(i).strEntidad & " WHERE " & strCriterio & ") " & advInfo(i).strAndOr)
            Next
        End If
        Dim strWhere As String = whereSQL.ToString
        strWhere = whereSQL.ToString.Substring(0, whereSQL.Length - 4)
        Return strWhere
    End Function

    Public Function SuperBusquedaAvanzada(ByVal strWhere As String, ByVal strVista As String) As DataTable
        Return AdminData.Execute("SELECT * FROM " & strVista & " " & strWhere, ExecuteCommand.ExecuteReader, False)
    End Function

    Public Function CrearCriterioSBA(ByVal f As Filter) As String
        Return AdminData.ComposeFilter(f)
    End Function
End Class
