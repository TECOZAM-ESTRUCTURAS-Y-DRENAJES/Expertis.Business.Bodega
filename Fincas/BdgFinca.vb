Public Class BdgFinca

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgFinca"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPorDefecto)
    End Sub

    <Task()> Public Shared Sub AsignarValoresPorDefecto(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("FincaTrabajo") = True
        data("FincaSigpac") = False
    End Sub

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMunicipio")) = 0 Then ApplicationService.GenerateError("El municipio es obligatorio.")
        If Length(data("IDVariedad")) = 0 Then ApplicationService.GenerateError("La variedad de vino es obligatoria.")
        If Length(data("DescFinca")) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarContador)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDContador")) Then
                data("CFinca") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDFinca") = Guid.NewGuid
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidatePrimaryKey(ByVal IDFinca As String, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgFinca().SelOnPrimaryKey(IDFinca)
        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count = 0 Then
            ApplicationService.GenerateError("Actualización en conflicto con el valor de la clave | de la tabla |.", IDFinca, cnEntidad)
        End If
    End Sub

    <Task()> Public Shared Sub ValidateDuplicateKey(ByVal IDFinca As String, ByVal services As ServiceProvider)
        Dim DtAux As DataTable = New BdgFinca().SelOnPrimaryKey(IDFinca)
        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se permite insertar una clave duplicada en la tabla |.", cnEntidad)
        End If
    End Sub

    <Serializable()> _
    Public Class StMatrizDatos
        Public Data As DataTable
        Public Dia1 As Integer
        Public Mes1 As Integer
        Public Dia2 As Integer
        Public Mes2 As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal Data As DataTable, ByVal Dia1 As Integer, ByVal Mes1 As Integer, ByVal Dia2 As Integer, ByVal Mes2 As Integer)
            Me.Data = Data
            Me.Dia1 = Dia1
            Me.Mes1 = Mes1
            Me.Dia2 = Dia2
            Me.Mes2 = Mes2
        End Sub
    End Class

    <Task()> Public Shared Function MatrizDatos(ByVal data As StMatrizDatos, ByVal services As ServiceProvider) As DataSet
        Dim ds As New DataSet
        If Not data.Data Is Nothing AndAlso data.Data.Rows.Count > 0 Then
            Dim fecha1, fecha2 As Date
            '//En este metodo se excluye el dia 29 de febrero. El año que se utiliza para construir las fechas 
            '//fecha1 y fecha2 es un año cualquiera no-bisiesto que solo sirve para construir la matriz de datos.
            '//El dia bisiesto se excluye porque existe la posibilidad de comparar intervalos de fechas para años 
            '//diferentes, si se incluye el dia 29 de febrero los datos resultantes muestran un salto para los 
            '//años que no son bisiestos.
            Dim NotLeapYear As Integer = 2007
            If data.Mes1 = 2 AndAlso data.Dia1 = 29 Then
                data.Dia1 = 28
            End If
            If data.Mes2 = 2 AndAlso data.Dia2 = 29 Then
                data.Dia2 = 28
            End If
            fecha1 = New Date(NotLeapYear, data.Mes1, data.Dia1)
            fecha2 = New Date(NotLeapYear, data.Mes2, data.Dia2)

            If (fecha1 > fecha2) Then
                ApplicationService.GenerateError("El intervalo de fechas no es válido.")
            Else
                Dim dv As New DataView(data.Data)
                dv.Sort = "IDFincaPadre,Vendimia DESC,IDFinca,IDVariable"

                '//cabeceras
                Dim bdgv As New BdgVariable
                Dim cabeceras As New DataTable("Cabecera")
                cabeceras.Columns.Add("Vendimia", GetType(Integer))
                cabeceras.Columns.Add("IDFinca", GetType(Guid))
                cabeceras.Columns.Add("DescFinca", GetType(String))
                cabeceras.Columns.Add("IDFincaPadre", GetType(Guid))
                cabeceras.Columns.Add("IDVariable", GetType(String))
                cabeceras.Columns.Add("Abreviatura", GetType(String))
                cabeceras.Columns.Add("Maximo", GetType(Double))
                cabeceras.Columns.Add("Minimo", GetType(Double))
                cabeceras.Columns.Add("UdMedida", GetType(String))
                cabeceras.Columns.Add("FactorEscala", GetType(Double))

                '//series
                Dim series As New DataTable("Series")
                For i As Integer = 1 To data.Data.Rows.Count
                    series.Columns.Add().DataType = GetType(Double)
                Next
                For i As Integer = 0 To fecha2.Subtract(fecha1).Days
                    series.Rows.Add(series.NewRow())
                Next

                Dim columnIndex, rowIndex As Integer
                Dim f1 As New Filter
                For Each drv As DataRowView In dv
                    fecha1 = New Date(drv("Vendimia"), data.Mes1, data.Dia1)
                    fecha2 = New Date(drv("Vendimia"), data.Mes2, data.Dia2)
                    f1.Clear()
                    f1.Add(New GuidFilterItem("IDFinca", drv("IDFinca")))
                    f1.Add(New StringFilterItem("IDVariable", drv("IDVariable")))
                    f1.Add(New DateFilterItem("Fecha", FilterOperator.GreaterThanOrEqual, fecha1))
                    f1.Add(New DateFilterItem("Fecha", FilterOperator.LessThanOrEqual, fecha2))
                    Dim valores As DataTable = New BdgFincaVariable().Filter(f1)
                    If valores.Rows.Count > 0 Then
                        Dim nr As DataRow = cabeceras.NewRow()
                        nr("Vendimia") = drv("Vendimia")
                        nr("IDFinca") = drv("IDFinca")
                        nr("IDVariable") = drv("IDVariable")
                        Dim finca As DataRow = New BdgFinca().GetItemRow(drv("IDFinca"))
                        nr("DescFinca") = finca("DescFinca")
                        nr("IDFincaPadre") = finca("IDFincaPadre")
                        Dim variable As DataRow = bdgv.GetItemRow(drv("IDVariable"))
                        nr("Abreviatura") = variable("Abreviatura")
                        nr("FactorEscala") = variable("FactorEscala")
                        nr("Maximo") = variable("Maximo")
                        nr("Minimo") = variable("Minimo")
                        nr("UdMedida") = variable("UdMedida")

                        cabeceras.Rows.Add(nr)

                        For Each dr As DataRow In valores.Rows
                            If (CType(dr("Fecha"), Date).Month <> 2) Or (CType(dr("Fecha"), Date).Day <> 29) Then
                                rowIndex = CType(dr("Fecha"), Date).Subtract(fecha1).Days
                                series.Rows(rowIndex)(columnIndex) = variable("FactorEscala") * dr("ValorNumerico")
                            End If
                        Next

                        columnIndex += 1
                    End If
                Next

                '//eliminar las columnas sin datos
                For i As Integer = series.Columns.Count - 1 To cabeceras.Rows.Count Step -1
                    series.Columns.RemoveAt(i)
                Next
                ds.Tables.Add(cabeceras)
                ds.Tables.Add(series)
            End If
        End If
        Return ds
    End Function

    <Serializable()> _
    Public Class StGeneracionObrasCampaña
        Public Origenes As DataTable
        Public IDObraModelo As Integer
        Public FechaDesde As DateTime
        Public FechaHasta As DateTime

        Public Sub New(ByVal dttOrigenes As DataTable, ByVal IDObraModelo As Integer, ByVal dtFechaDesde As DateTime, ByVal dtFechaHasta As DateTime)
            Me.Origenes = dttOrigenes
            Me.IDObraModelo = IDObraModelo
            Me.FechaDesde = dtFechaDesde
            Me.FechaHasta = dtFechaHasta
        End Sub

    End Class

    <Serializable()> _
   Public Class StGeneracionObrasCampañaResult
        Public NObra(-1) As String
        Public Errores(-1) As String
    End Class

    <Task()> _
    Public Shared Function GeneracionObrasCampaña(ByVal data As StGeneracionObrasCampaña, ByVal services As ServiceProvider) As StGeneracionObrasCampañaResult
        Dim processResult As New StGeneracionObrasCampañaResult()
        If data Is Nothing Then
            Return processResult
        End If

        Dim bdgFinca As New BdgFinca()
        Dim bdgObra As New ObraCabecera()
        For Each dtrOrigen As DataRow In data.Origenes.Rows
            Try
                'Copiar obra
                Dim infoCopia As New dataCopiaObra(data.IDObraModelo)
                infoCopia.ConfiguracionCopia = New dataConfigCopiaObra()
                infoCopia.ConfiguracionCopia.CopiarTrabajos = True
                infoCopia.ConfiguracionCopia.CopiarSubTrabajos = False

                Dim result As ResultadoCopiaObra = ProcessServer.ExecuteTask(Of dataCopiaObra, ResultadoCopiaObra)(AddressOf ObraCabecera.CopiarObra, infoCopia, services)

                'Actualizar los datos de la cabecera
                Dim dttObra As DataTable = bdgObra.SelOnPrimaryKey(result.IDObra)
                '- NObra = "CFinca" + "Año" de la Fecha Hasta.
                dttObra.Rows(0)("NObra") = String.Format("{0}{1}", dtrOrigen("CFinca"), data.FechaHasta.Year)
                result.NObra = dttObra.Rows(0)("NObra")

                '- DescObra = "DescFinca" + "Año" de la Fecha Hasta.
                dttObra.Rows(0)("DescObra") = String.Format("{0}{1}", dtrOrigen("DescFinca"), data.FechaHasta.Year)
                '- IDObraPadre = tbBdgFinca.IDObra
                dttObra.Rows(0)("IDObraPadre") = dtrOrigen("IDObra")

                '- fechas
                dttObra.Rows(0)("FechaInicio") = data.FechaDesde
                dttObra.Rows(0)("FechaFin") = data.FechaHasta

                bdgObra.Update(dttObra)

                '- Finca
                Dim dttFinca As DataTable = bdgFinca.SelOnPrimaryKey(dtrOrigen("IDFinca"))
                If (Not dttFinca Is Nothing AndAlso dttFinca.Rows.Count > 0) Then
                    dttFinca.Rows(0)("IDObraCampaña") = result.IDObra
                    bdgFinca.Update(dttFinca)
                End If

                ReDim Preserve processResult.NObra(UBound(processResult.NObra) + 1)
                processResult.NObra(UBound(processResult.NObra)) = result.NObra
            Catch ex As Exception
                ReDim Preserve processResult.Errores(UBound(processResult.Errores) + 1)
                processResult.Errores(UBound(processResult.Errores)) = ex.Message
            End Try
        Next

        Return processResult

    End Function

#End Region

End Class