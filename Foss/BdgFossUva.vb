Public Class BdgFossUva
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbBdgFossUva"


    Private Const cnDelimiter As String = ";"
    Private Const cnInitVar As Integer = 7
    Private Const cnPosDate As Integer = 35
    Private Const cnPosFinca As Integer = 1
    Private Const cnPosCodFossUva As Integer = 0
    Private Const AnalisisFoss As String = "BDGFOSUVA"
    Private Const AnalisisPredeterminado As String = "vNegAnalisisPredeterminadoFoss"
    Private Const AnalisisFinca As String = "vNegBdgFincaAnalisis"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Public Overloads Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        If dttSource Is Nothing Then Exit Function

        For Each oRw As DataRow In dttSource.Rows
            If oRw.RowState = DataRowState.Added Then
                If oRw.IsNull("IDFossUva") OrElse oRw("IDFossUva") = 0 Then
                    oRw("IDFossUva") = AdminData.GetAutoNumeric
                End If
            End If
        Next

        BusinessHelper.UpdateTable(dttSource)
        Update = dttSource
    End Function

    Public Function AccessFilterCompose(ByVal strFechaDesde As String, ByVal strFechaHasta As String) As String
        Dim e As New Filter
        If Len(strFechaDesde) > 0 Then e.Add("Date", FilterOperator.GreaterThanOrEqual, CDate(strFechaDesde), FilterType.DateTime)
        If Len(strFechaHasta) > 0 Then e.Add("Date", FilterOperator.LessThan, DateAdd(DateInterval.Day, 1, CDate(strFechaHasta)), FilterType.DateTime)
        Return e.Compose(New AdoFilterComposer)
    End Function

    <Serializable()> _
    Public Class StProcesarLineaUva
        Public strLinea As String
        Public dtPlantilla As DataTable
        Public strFechaDesde As String
        Public strFechaHasta As String
        Public dtData As DataTable
        Public strSinEntrada As String = String.Empty
        Public strFechaIncorrecta As String = String.Empty

        Public Sub New()
        End Sub

        Public Sub New(ByVal Linea As String, ByVal DtPlantilla As DataTable, ByVal FechaDesde As String, ByVal FechaHasta As String, ByVal DtData As DataTable)
            Me.strLinea = Linea
            Me.dtPlantilla = DtPlantilla
            Me.strFechaDesde = FechaDesde
            Me.strFechaHasta = FechaHasta
            Me.dtData = DtData
        End Sub
    End Class

    <Task()> Public Shared Function ProcesarLineaFossUva(ByVal data As StProcesarLineaUva, ByVal services As ServiceProvider) As StProcesarLineaUva
        Dim e As New Filter
        Dim blnFechaOK As Boolean
        Dim strNEntrada As String = String.Empty
        Dim drLine As DataRow

        If Not data.dtData Is Nothing Then drLine = data.dtData.NewRow

        data.strLinea = Replace(data.strLinea, """", String.Empty)
        If Len(data.strLinea) > 0 Then
            Dim strArrayLine() As String = Split(data.strLinea, cnDelimiter)
            Dim strCodigoFoss As String = strArrayLine(cnPosCodFossUva) & String.Empty
            If strArrayLine.Length > 0 Then
                If IsDate(strArrayLine(cnPosDate)) Then

                    blnFechaOK = True
                    Dim FechaLinea As New Date(CDate(strArrayLine(cnPosDate)).Year, CDate(strArrayLine(cnPosDate)).Month, CDate(strArrayLine(cnPosDate)).Day)
                    If Len(data.strFechaDesde) > 0 Then
                        If FechaLinea < CDate(data.strFechaDesde) Then blnFechaOK = False
                    End If
                    If Len(data.strFechaHasta) > 0 And blnFechaOK Then
                        If FechaLinea > CDate(data.strFechaHasta) Then blnFechaOK = False
                    End If

                    If blnFechaOK Then
                        If Len(Trim(strArrayLine(cnPosFinca))) > 0 Then
                            If Left(Trim(strArrayLine(cnPosFinca)), 1) = "A" Then
                                strNEntrada = Right(Trim(strArrayLine(cnPosFinca)), Len(Trim(strArrayLine(cnPosFinca))) - 1)
                            End If
                        End If
                        If Length(strNEntrada) > 0 Then
                            drLine("NEntrada") = strNEntrada
                            drLine("Vendimia") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf BdgVendimia.UltimaVendimia, New Object, services)
                            drLine("CodigoFossUva") = strCodigoFoss
                            drLine("RecuperadoEntrada") = False
                            drLine("Fecha") = FechaLinea

                            Dim strFormat As String
                            Dim i As Integer = 1
                            For Each drVar As DataRow In data.dtPlantilla.Rows
                                If Len(drVar("Orden")) > 0 AndAlso drVar("Orden") > 0 Then
                                    If Len(drVar("ColFossUva") & String.Empty) > 0 Then
                                        drLine("CodVar" & drVar("Orden")) = drVar("IDVariable")
                                        If strArrayLine.Length > (drVar("ColFossUva") + cnInitVar) Then
                                            If Length(Trim(strArrayLine(drVar("ColFossUva") + cnInitVar))) = 0 Then
                                                drLine("Var" & drVar("Orden")) = System.DBNull.Value
                                            Else
                                                If drVar("NDecimales") = 0 Then
                                                    strFormat = "0"
                                                Else
                                                    strFormat = "0."
                                                End If
                                                If IsNumeric(Trim(strArrayLine(drVar("ColFossUva") + cnInitVar))) Then
                                                    drLine("Var" & drVar("Orden")) = xRound(Replace(Trim(strArrayLine(drVar("ColFossUva") + cnInitVar)), ".", ","), drVar("NDecimales"))
                                                Else
                                                    drLine("Var" & drVar("Orden")) = Trim(strArrayLine(drVar("ColFossUva") + cnInitVar))
                                                End If

                                            End If
                                        Else
                                            drLine("Var" & drVar("Orden")) = System.DBNull.Value
                                        End If
                                        i += 1
                                    End If
                                End If
                            Next
                            If Not drLine Is Nothing Then data.dtData.Rows.Add(drLine)
                        Else
                            'Sin Entrada
                            data.strSinEntrada = strCodigoFoss
                        End If
                        strNEntrada = String.Empty
                    Else
                        'Fecha fuera de rango
                        data.strFechaIncorrecta = strCodigoFoss
                    End If
                Else
                    'Fecha vacía
                    data.strFechaIncorrecta = strCodigoFoss
                End If
            End If
        End If

        Return data
    End Function

    Public Function ValidatedImportDataEntradaFoss(ByRef dttData As DataTable) As Boolean
        Dim e As New Filter
        If Not IsNothing(dttData) AndAlso dttData.Rows.Count Then
            Dim objEntrada As New BdgEntrada
            Dim objEntradaVino As New BdgEntradaAnalisis

            For Each drdata As DataRow In dttData.Rows

                drdata("Entrada") = False ' no existe esa entrada de uva a esa fecha
                drdata("Execute") = False ' ejecuta directamente la importacion
                drdata("Replace") = 0 ' existen datos de variables (pregunta si se quiere reemplazar)

                e.Clear()
                e.Add("NEntrada", drdata("NEntrada"))
                e.Add("Vendimia", drdata("Vendimia"))
                ''e.Add("Fecha", drdata("Fecha"))
                'e.Add("Fecha", FilterOperator.GreaterThanOrEqual, drdata("Fecha"))

                Dim dtEntrada As DataTable = objEntrada.Filter(e)
                If dtEntrada.Rows.Count > 0 Then

                    drdata("IDEntrada") = dtEntrada.Rows(0)("IDEntrada")

                    e.Clear()
                    e.Add("IDEntrada", drdata("IDEntrada"))
                    'e.Add(New IsNullFilterItem("Valor", False))

                    Dim dtAnalisisVino As DataTable = objEntradaVino.Filter(e)
                    If dtAnalisisVino.Rows.Count > 0 Then
                        If dtAnalisisVino.Select("Not Valor is null").Length() = 0 Then
                            drdata("Replace") = 2
                        Else
                            drdata("Replace") = 1
                        End If

                    Else
                        drdata("Execute") = True
                    End If
                Else
                    drdata("Entrada") = True
                End If
            Next
        End If
    End Function

    Public Function ExisteCodigoFossEnEntradaAnalisis(ByVal strCodigoFossUva As String) As Boolean
        Dim objVinAna As New BdgEntradaAnalisis
        Dim e As New Filter

        e.Add("CodigoFossUva", strCodigoFossUva)
        Dim dtEntAna As DataTable = AdminData.GetData("tbBdgEntradaAnalisis", e, "top 1 CodigoFossUva")
        If Not IsNothing(dtEntAna) AndAlso dtEntAna.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function ExisteCodigoFossEnFossUva(ByVal strCodigoFossUva As String) As Boolean
        Dim objVinAna As New BdgEntradaAnalisis
        Dim e As New Filter

        e.Add("CodigoFossUva", strCodigoFossUva)
        Dim dtEntAna As DataTable = AdminData.GetData("tbBdgFossUva", e, "top 1 CodigoFossUva")
        If Not IsNothing(dtEntAna) AndAlso dtEntAna.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    <Task()> Public Shared Function MirarCodigoFossEnEntradaAnalisis(ByVal strCodigoFossUva As String, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("tbBdgEntradaAnalisis", New StringFilterItem("CodigoFossUva", strCodigoFossUva))
    End Function

    Public Function ImportDataEntradaFoss(ByVal dttData As DataTable, ByVal Parametro As String) As Boolean
        Me.BeginTx()
        'Dim objEntrada As New BdgEntrada
        Dim objAnalisis As New BdgAnalisisVariable
        Dim objEntradaVar As New BdgEntradaAnalisis
        Dim blnReturn As Boolean
        Dim e As New Filter
        Dim eDelete As New Filter(FilterUnionOperator.Or)
        Dim objParam As New Parametro
        Dim strIDFossIN(-1) As String

        Dim strParamAnalisis As String = objParam.ObtenerPredeterminado(AnalisisFoss)
        e.Add("IDAnalisis", strParamAnalisis)

        Dim dtAnalisis As DataTable = objAnalisis.Filter(e, "Orden")

        Dim dttEntradaAnalisis As DataTable = objEntradaVar.AddNew

        If Not IsNothing(dtAnalisis) And dtAnalisis.Rows.Count > 0 Then

            If Not IsNothing(dttEntradaAnalisis) Then
                For Each drData As DataRow In dttData.Select("Execute=true")
                    Dim drFoss As DataRow = GetItemRow(drData("IDFossUva"))

                    For Each drAnalisis As DataRow In dtAnalisis.Rows
                        Dim drVar As DataRow = dttEntradaAnalisis.NewRow
                        '''''''''''''''''''''
                        drVar("CodigoFossUva") = drData("CodigoFossUva")
                        '''''''''''''''''''''
                        drVar("IDEntrada") = drData("IDEntrada")

                        If Length(drFoss("CodVar" & drAnalisis("Orden")) & String.Empty) > 0 Then
                            drVar("IDVariable") = drFoss("CodVar" & drAnalisis("Orden"))
                            drVar("Valor") = drFoss("Var" & drAnalisis("Orden"))
                            If IsNumeric(drFoss("Var" & drAnalisis("Orden"))) Then
                                drVar("ValorNumerico") = Replace(drFoss("Var" & drAnalisis("Orden")), ".", ",")
                            Else
                                drVar("ValorNumerico") = 0
                            End If
                        Else
                            drVar("IDVariable") = drAnalisis("IDVariable")
                        End If

                        drVar("Orden") = drAnalisis("Orden")

                        dttEntradaAnalisis.Rows.Add(drVar)
                    Next

                    If drData("Replace") = 1 OrElse drData("Replace") = 2 Then
                        eDelete.Add("IDEntrada", drData("IDEntrada"))
                    End If

                    ReDim Preserve strIDFossIN(strIDFossIN.Length)
                    strIDFossIN(strIDFossIN.Length - 1) = drFoss("IDFossUva")

                Next
            End If
        End If

        If eDelete.Count > 0 Then
            Dim dttDelete As DataTable = objEntradaVar.Filter(eDelete)
            If Not IsNothing(dttDelete) AndAlso dttDelete.Rows.Count > 0 Then
                objEntradaVar.Delete(dttDelete)
            End If
        End If

        If Not IsNothing(dttEntradaAnalisis) AndAlso dttEntradaAnalisis.Rows.Count > 0 Then
            BusinessHelper.UpdateTable(dttEntradaAnalisis)
            blnReturn = True
        End If

        If strIDFossIN.Length > 0 Then
            e.Clear()
            Dim dttFoss As DataTable = Me.Filter(New InListFilterItem("IDFossUva", strIDFossIN, FilterType.String))
            If Not IsNothing(dttFoss) AndAlso dttFoss.Rows.Count > 0 Then
                For Each drFoss As DataRow In dttFoss.Rows
                    drFoss("RecuperadoEntrada") = True
                Next
                BusinessHelper.UpdateTable(dttFoss)
                blnReturn = True
            End If
        End If
        Return blnReturn
    End Function

    Public Function AccessDataIRTF(ByVal dttOrigen As DataTable) As String
        Dim Services As New ServiceProvider
        Dim strReimportados As String = String.Empty
        Dim lngNumImportados As Long = 0
        Dim strMensaje As String = String.Empty
        Dim strCampoOrigen As String = String.Empty

        Dim dttFoss As DataTable
        Dim strParamAnalisis As String
        Dim oPrm As New Parametro
        Dim e As New Filter
        Dim strCodigoFoss As String

        strParamAnalisis = oPrm.ObtenerPredeterminado(AnalisisFoss)
        e.Add("IDAnalisis", strParamAnalisis)
        e.Add(New IsNullFilterItem("ColFossUva", False))
        Dim dttAnalisis As DataTable = AdminData.GetData(AnalisisPredeterminado, e, , "Orden")

        If Not IsNothing(dttAnalisis) AndAlso dttAnalisis.Rows.Count > 0 Then

            dttFoss = AddNew()

            'Por cada análisis de IRTF
            For Each drOrigen As DataRow In dttOrigen.Rows

                'Comprobar si el análisis ya se ha importado para avisar.
                strCodigoFoss = drOrigen("Identite") & "_" & drOrigen("Date")
                If ExisteCodigoFossEnEntradaAnalisis(strCodigoFoss) Or ExisteCodigoFossEnFossUva(strCodigoFoss) Then
                    strReimportados = strReimportados & "Nº " & strCodigoFoss & ", Muestra: " & drOrigen("Identite") & vbNewLine
                End If

                'Insertar en tbBdgFossUva
                Dim drFoss As DataRow = dttFoss.NewRow
                drFoss("IDFossUva") = AdminData.GetAutoNumeric()
                drFoss("NEntrada") = drOrigen("Identite") & String.Empty
                drFoss("Fecha") = drOrigen("Date")
                drFoss("RecuperadoEntrada") = False
                drFoss("CodigoFossUva") = strCodigoFoss
                drFoss("Vendimia") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf BdgVendimia.UltimaVendimia, New Object, services)

                For Each drAnalisis As DataRow In dttAnalisis.Rows
                    drFoss("CodVar" & drAnalisis("Orden")) = drAnalisis("IDVariable")
                    strCampoOrigen = "Critere " & drAnalisis("ColFossUva")
                    If IsDBNull(drOrigen(strCampoOrigen)) Or drOrigen(strCampoOrigen) = -1 Then
                        drFoss("Var" & drAnalisis("Orden")) = System.DBNull.Value
                    Else
                        drFoss("Var" & drAnalisis("Orden")) = xRound(drOrigen(strCampoOrigen), drAnalisis("NDecimales"))
                    End If
                Next
                dttFoss.Rows.Add(drFoss)
                lngNumImportados = lngNumImportados + 1
            Next

            If Not IsNothing(dttFoss) AndAlso dttFoss.Rows.Count > 0 Then
                BusinessHelper.UpdateTable(dttFoss)
            End If
        End If

        If lngNumImportados > 0 Then
            strMensaje = "Se han importado " & CStr(lngNumImportados) & " análisis." & vbNewLine
        End If
        If Length(strReimportados) > 0 Then
            strMensaje = strMensaje & "Los análisis siguientes se han vuelto a importar aunque ya estaban importados:" & vbNewLine & strReimportados
        End If
        Return strMensaje

    End Function

End Class
