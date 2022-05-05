Public Class BdgFossFinca

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFossFinca"

#End Region

#Region "Constantes"

    Private Const cnDelimiter As String = ";"
    Private Const cnInitVar As Integer = 7
    Private Const cnPosDate As Integer = 35
    Private Const cnPosFinca As Integer = 1
    Private Const cnPosCodFossUva As Integer = 0
    Private Const AnalisisFoss As String = "BDGFOSFIN"
    Private Const AnalisisPredeterminado As String = "vNegAnalisisPredeterminadoFoss"
    Private Const AnalisisFinca As String = "vNegBdgFincaAnalisis"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDFossUva")) = 0 OrElse data("IDFossUva") = 0 Then
                data("IDFossUva") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class StProcesarLineaFinca
        Public strLinea As String
        Public dtPlantilla As DataTable
        Public strFechaDesde As String
        Public strFechaHasta As String
        Public dtData As DataTable
        Public strSinFinca As String = String.Empty
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

    <Task()> Public Shared Function ProcesarLineaFossFinca(ByVal data As StProcesarLineaFinca, ByVal services As ServiceProvider) As StProcesarLineaFinca
        Dim e As New Filter
        Dim blnFechaOK As Boolean
        Dim strFinca As String = String.Empty
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
                            If Left(Trim(strArrayLine(cnPosFinca)), 1) = "C" Then
                                strFinca = Right(Trim(strArrayLine(cnPosFinca)), Len(Trim(strArrayLine(cnPosFinca))) - 1)
                            End If
                        End If
                        If Length(strFinca) > 0 Then
                            drLine("CFinca") = strFinca
                            If Length(strFinca) > 0 Then
                                Dim de As New Engine.BE.DataEngine
                                Dim fil As New Filter
                                fil.Add("CFinca", FilterOperator.Equal, strFinca)
                                Dim dtFinca As DataTable = de.Filter("TbBdgFinca", fil)
                                If Not dtFinca Is Nothing AndAlso dtFinca.Rows.Count > 0 Then
                                    drLine("DescFinca") = dtFinca.Rows(0)("DescFinca")
                                Else
                                    drLine("DescFinca") = String.Empty
                                End If
                            Else
                                drLine("DescFinca") = String.Empty
                            End If

                            drLine("CodigoFossUva") = strCodigoFoss
                            drLine("RecuperadoFinca") = False
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
                            'Sin finca
                            data.strSinFinca = strCodigoFoss
                        End If
                        strFinca = String.Empty
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

    <Task()> Public Shared Function MirarCodigoFossEnFincaAnalisis(ByVal strCodigoFossUva As String, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("tbBdgFincaAnalisis", New StringFilterItem("CodigoFossUva", strCodigoFossUva))
    End Function

    <Task()> Public Shared Function ValidatedImportDataFincaFoss(ByVal dttData As DataTable, ByVal services As ServiceProvider) As DataTable
        Dim e As New Filter
        If Not IsNothing(dttData) AndAlso dttData.Rows.Count Then
            Dim objFinca As New BdgFinca
            Dim objAnalisisFinca As New BdgFincaAnalisis
            Dim objParam As New Parametro
            Dim strParamAnalisis As String = objParam.ObtenerPredeterminado(AnalisisFoss)

            For Each drdata As DataRow In dttData.Rows
                drdata("Finca") = False
                drdata("Execute") = False
                drdata("Replace") = False
                e.Clear()
                e.Add("CFinca", drdata("CFinca"))
                Dim dtFinca As DataTable = objFinca.Filter(e)
                If dtFinca.Rows.Count > 0 Then
                    drdata("IDFinca") = dtFinca.Rows(0)("IDFinca")

                    e.Clear()
                    e.Add("IDAnalisis", strParamAnalisis)
                    e.Add("IDFinca", drdata("IDFinca"))
                    e.Add("Fecha", drdata("Fecha"))

                    drdata("IDAnalisis") = strParamAnalisis

                    Dim dtAnalisisFinca As DataTable = objAnalisisFinca.Filter(e)
                    If dtAnalisisFinca.Rows.Count > 0 Then
                        drdata("Replace") = True
                    Else
                        drdata("Execute") = True
                    End If
                Else
                    drdata("Finca") = True
                End If

            Next
            Return dttData
        End If
    End Function

    <Task()> Public Shared Function ImportDataFincaFoss(ByVal dttData As DataTable, ByVal services As ServiceProvider) As Boolean
        AdminData.BeginTx()

        Dim objFincaAn As New BdgFincaAnalisis
        Dim objFincaAnVar As New BdgFincaVariable
        Dim objAnalisis As New BdgAnalisisVariable
        Dim objParam As New Parametro

        Dim blnReturn As Boolean
        Dim e As New Filter
        Dim eDelete As New Filter(FilterUnionOperator.Or)

        Dim dttFincaAnalisis As DataTable = objFincaAn.AddNew
        Dim oCont As New Contador
        Dim dtFincaAnCont As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, objFincaAn.GetType.Name, services)
        Dim strFincaAnCont As String
        If Not dtFincaAnCont Is Nothing And dtFincaAnCont.Rows.Count > 0 Then
            strFincaAnCont = dtFincaAnCont.Rows(0)("IdContador")
        End If
        Dim dttFincaAnVar As DataTable = objFincaAnVar.AddNew

        Dim strIDFossIN(-1) As String


        Dim strParamAnalisis As String = objParam.ObtenerPredeterminado(AnalisisFoss)
        e.Add("IDAnalisis", strParamAnalisis)
        Dim dtAnalisis As DataTable = objAnalisis.Filter(e, "Orden")

        For Each drData As DataRow In dttData.Select("Execute=true")
            Dim drFinca As DataRow

            If Not drData("Replace") Then

                drFinca = dttFincaAnalisis.NewRow
                drFinca("IDFinca") = drData("IDFinca")
                drFinca("IDAnalisis") = strParamAnalisis
                drFinca("Fecha") = drData("Fecha")
                drFinca("CodigoFossUva") = drData("CodigoFossUva")
                If Len(strFincaAnCont) > 0 Then
                    drFinca("NFincaAnalisis") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, strFincaAnCont, services)
                    drFinca("IdContador") = strFincaAnCont
                End If
                dttFincaAnalisis.Rows.Add(drFinca)

            End If

            Dim drFoss As DataRow = New BdgFossFinca().GetItemRow(drData("IDFossUva"))

            For Each drAnalisis As DataRow In dtAnalisis.Rows
                Dim drVar As DataRow = dttFincaAnVar.NewRow
                drVar("Fecha") = drFoss("Fecha")
                drVar("IDAnalisis") = strParamAnalisis
                drVar("IDFinca") = drData("IDFinca")
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

                dttFincaAnVar.Rows.Add(drVar)
            Next

            If drData("Replace") Then

                e.Clear()
                e.Add("IDFinca", FilterOperator.Equal, drData("IDFinca"), FilterType.Guid)
                e.Add("Fecha", FilterOperator.Equal, drData("Fecha"), FilterType.DateTime)
                e.Add("IDAnalisis", FilterOperator.Equal, strParamAnalisis, FilterType.String)

                Dim dttDelete As DataTable = objFincaAnVar.Filter(e)
                If Not IsNothing(dttDelete) AndAlso dttDelete.Rows.Count > 0 Then
                    objFincaAnVar.Delete(dttDelete)
                End If
            End If

            ReDim Preserve strIDFossIN(strIDFossIN.Length)
            strIDFossIN(strIDFossIN.Length - 1) = drFoss("IDFossUva")
        Next

        If Not IsNothing(dttFincaAnalisis) AndAlso dttFincaAnalisis.Rows.Count > 0 Then
            BusinessHelper.UpdateTable(dttFincaAnalisis)
            blnReturn = True
        End If

        If Not IsNothing(dttFincaAnVar) AndAlso dttFincaAnVar.Rows.Count > 0 Then
            BusinessHelper.UpdateTable(dttFincaAnVar)
            blnReturn = True
        End If

        If strIDFossIN.Length > 0 Then
            e.Clear()
            Dim dttFoss As DataTable = New BdgFossFinca().Filter(New InListFilterItem("IDFossUva", strIDFossIN, FilterType.String))
            If Not IsNothing(dttFoss) AndAlso dttFoss.Rows.Count > 0 Then
                For Each drFoss As DataRow In dttFoss.Rows
                    drFoss("RecuperadoFinca") = True
                Next
                BusinessHelper.UpdateTable(dttFoss)
                blnReturn = True
            End If
        End If

        Return blnReturn
    End Function

#End Region

End Class