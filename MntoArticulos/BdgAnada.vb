Imports System.Collections.Generic

Public Class BdgAnada

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbBdgMaestroAnada"

#End Region

#Region "Eventos Entidad"
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub


#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAnada")) = 0 Then ApplicationService.GenerateError("La Añada es un dato obligatorio.")
    End Sub

#Region "Codificación/Explosión/Alta Artículos"

    <Serializable()> _
    Public Class DataCopiaAnada
        Public AnadaDestino As String
        Public IDArticuloOrigen As String
        Public IDArticuloNew As String
        Public DescArticuloNew As String
        Public Opciones As New Dictionary(Of String, Boolean)

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticuloOrigen As String, ByVal IDArticuloNew As String, ByVal DescArticuloNew As String, ByVal AnadaDestino As String, ByVal Opciones As Dictionary(Of String, Boolean))
            Me.IDArticuloOrigen = IDArticuloOrigen
            Me.IDArticuloNew = IDArticuloNew
            Me.DescArticuloNew = DescArticuloNew
            Me.AnadaDestino = AnadaDestino
            Me.Opciones = Opciones
        End Sub
    End Class

    Private Shared MIDComponente As String = String.Empty
    Private Shared Function LoadElement(ByVal Element As CreateElement) As Boolean
        Return (MIDComponente = Element.NElement)
    End Function

    <Task()> Public Shared Function CopiarAnada(ByVal data As List(Of DataCopiaAnada), ByVal services As ServiceProvider) As LogProcess
        If Not data Is Nothing AndAlso data.Count > 0 Then
            Dim StLog As New LogProcess
            Dim IntPosErrorLog As Integer = 0
            Dim IntPosCreateLog As Integer = 0
            For Each StData As DataCopiaAnada In data
                Dim BlnFind As Boolean = False
                MIDComponente = StData.IDArticuloNew
                If Not StLog.CreatedElements Is Nothing AndAlso StLog.CreatedElements.Length > 0 Then
                    Dim FindElem As CreateElement = Array.Find(StLog.CreatedElements, AddressOf LoadElement)
                    If Not FindElem Is Nothing Then
                        BlnFind = True
                    End If
                End If
                If Not BlnFind Then
                    Dim DtCheckArt As DataTable = New Articulo().SelOnPrimaryKey(StData.IDArticuloNew)
                    If DtCheckArt Is Nothing OrElse DtCheckArt.Rows.Count = 0 Then
                        Dim StDataCopia As New Articulo.DatosArtCopia
                        StDataCopia.IDArticulo = StData.IDArticuloOrigen
                        StDataCopia.IDArticuloNew = StData.IDArticuloNew
                        ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.GeneraNuevoArticulo, StDataCopia, services)
                        StDataCopia.dtArticuloNew.Rows(0)("IDAnada") = StData.AnadaDestino
                        StDataCopia.dtArticuloNew.Rows(0)("DescArticulo") = StData.DescArticuloNew

                        ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarArticuloCosteEstandar, StDataCopia, services)
                        If StData.Opciones("Estructuras") Then
                            ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarArticuloEstructura, StDataCopia, services)
                            ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarEstructura, StDataCopia, services)

                            If Not StDataCopia.dtEstructuraNew Is Nothing AndAlso StDataCopia.dtEstructuraNew.Rows.Count > 0 Then
                                For Each DrEst As DataRow In StDataCopia.dtEstructuraNew.Select
                                    Dim LstFind As List(Of DataCopiaAnada) = (From DataFind As DataCopiaAnada In data _
                                                                              Select DataFind _
                                                                              Where DataFind.IDArticuloOrigen = DrEst("IDComponente")).ToList
                                    If LstFind.Count > 0 Then
                                        DrEst("IDComponente") = LstFind(0).IDArticuloNew
                                    End If
                                Next
                            End If
                        End If
                        Try
                            If StData.Opciones("Rutas") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarArticuloRuta, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarRuta, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarRutaParametro, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarRutaUtillaje, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarRutaOficio, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarRutaAlternativa, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarRutaAMFE, StDataCopia, services)
                            End If
                            If StData.Opciones("CaracteristicasMaq") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarCaracteristicasMaq, StDataCopia, services)
                            End If
                            If StData.Opciones("Caracteristicas") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarCaracteristicas, StDataCopia, services)
                            End If
                            If StData.Opciones("Especificaciones") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarEspecificaciones, StDataCopia, services)
                            End If
                            If StData.Opciones("Promociones") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarPromociones, StDataCopia, services)
                            End If
                            If StData.Opciones("Idiomas") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarIdiomas, StDataCopia, services)
                            End If
                            If StData.Opciones("Documentos") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarDocumentos, StDataCopia, services)
                            End If
                            If StData.Opciones("Analitica") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarAnalitica, StDataCopia, services)
                            End If
                            If StData.Opciones("CostesVarios") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarCostesVarios, StDataCopia, services)
                            End If
                            If StData.Opciones("Proveedores") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarProveedores, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarProveedoresLineas, StDataCopia, services)
                            End If
                            If StData.Opciones("Tarifas") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarTarifasArt, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarTarifasArtLineas, StDataCopia, services)
                            End If
                            If StData.Opciones("Clientes") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarClientes, StDataCopia, services)
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarClientesLineas, StDataCopia, services)
                            End If
                            If StData.Opciones("Unidades") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarUnidades, StDataCopia, services)
                            End If
                            If StData.Opciones("Almacenes") Then
                                ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.CopiarAlmacenes, StDataCopia, services)
                            End If
                            ProcessServer.ExecuteTask(Of Articulo.DatosArtCopia)(AddressOf Articulo.GuardarCopiaArticulo, StDataCopia, services)

                            ReDim Preserve StLog.CreatedElements(IntPosCreateLog)
                            StLog.CreatedElements(IntPosCreateLog) = New CreateElement
                            StLog.CreatedElements(IntPosCreateLog).NElement = StData.IDArticuloNew
                            IntPosCreateLog += 1
                        Catch ex As Exception
                            ReDim Preserve StLog.Errors(IntPosErrorLog)
                            StLog.Errors(IntPosErrorLog) = New ClassErrors
                            StLog.Errors(IntPosErrorLog).Elements = StData.IDArticuloOrigen
                            StLog.Errors(IntPosErrorLog).MessageError = "Ha ocurrido un error en la generación del Artículo: " & StData.IDArticuloNew & " | " & ex.Message
                            IntPosErrorLog += 1
                        End Try
                    End If
                End If
            Next
            Return StLog
        End If
    End Function

    <Serializable()> _
    Public Class DataGetExpArt
        Public IDArticulo As String = String.Empty
        Public IDAnada As String = String.Empty
        Public IDTipo As String = String.Empty
        Public IDFamilia As String = String.Empty
        Public IDSubFamilia As String = String.Empty
        Public IDAnadaDestino As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal IDAnada As String, ByVal IDTipo As String, ByVal IDFamilia As String, ByVal IDSubFamilia As String, ByVal IDAnadaDestino As String)
            Me.IDArticulo = IDArticulo
            Me.IDAnada = IDAnada
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDSubFamilia = IDSubFamilia
            Me.IDAnadaDestino = IDAnadaDestino
        End Sub
    End Class

    <Task()> Public Shared Function GetExplosionArticulos(ByVal data As DataGetExpArt, ByVal services As ServiceProvider) As DataTable
        Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
        SqlCmd.CommandText = "sp_EstructuraExplosionCodificada"
        SqlCmd.CommandType = CommandType.StoredProcedure

        Dim SqlParamIDArticulo As Common.DbParameter = SqlCmd.CreateParameter
        SqlParamIDArticulo.Direction = ParameterDirection.Input
        SqlParamIDArticulo.ParameterName = "@FilIDArticulo"
        SqlParamIDArticulo.Value = data.IDArticulo
        SqlCmd.Parameters.Add(SqlParamIDArticulo)

        Dim SqlParamIDAnada As Common.DbParameter = SqlCmd.CreateParameter
        SqlParamIDAnada.Direction = ParameterDirection.Input
        SqlParamIDAnada.ParameterName = "@FilIDAnada"
        SqlParamIDAnada.Value = data.IDAnada
        SqlCmd.Parameters.Add(SqlParamIDAnada)

        Dim SqlParamIDTipo As Common.DbParameter = SqlCmd.CreateParameter
        SqlParamIDTipo.Direction = ParameterDirection.Input
        SqlParamIDTipo.ParameterName = "@FilIDTipo"
        SqlParamIDTipo.Value = data.IDTipo
        SqlCmd.Parameters.Add(SqlParamIDTipo)

        Dim SqlParamIDFamilia As Common.DbParameter = SqlCmd.CreateParameter
        SqlParamIDFamilia.Direction = ParameterDirection.Input
        SqlParamIDFamilia.ParameterName = "@FilIDFamilia"
        SqlParamIDFamilia.Value = data.IDFamilia
        SqlCmd.Parameters.Add(SqlParamIDFamilia)

        Dim SqlParamIDSubFamilia As Common.DbParameter = SqlCmd.CreateParameter
        SqlParamIDSubFamilia.Direction = ParameterDirection.Input
        SqlParamIDSubFamilia.ParameterName = "@FilIDSubFamilia"
        SqlParamIDSubFamilia.Value = data.IDSubFamilia
        SqlCmd.Parameters.Add(SqlParamIDSubFamilia)

        Dim DtGrid As DataTable = AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteReader)
        Dim StCodif As New DataCheckCodif(DtGrid, data.IDAnadaDestino)
        DtGrid = ProcessServer.ExecuteTask(Of DataCheckCodif, DataTable)(AddressOf CheckCodifArticulos, StCodif, services)

        Return DtGrid
    End Function

    <Serializable()> _
    Public Class DataCheckCodif
        Public DtData As DataTable
        Public IDAnadaDestino As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal DtData As DataTable, ByVal IDAnadaDestino As String)
            Me.DtData = DtData
            Me.IDAnadaDestino = IDAnadaDestino
        End Sub
    End Class

    <Task()> Public Shared Function CheckCodifArticulos(ByVal data As DataCheckCodif, ByVal services As ServiceProvider) As DataTable
        If Not data Is Nothing AndAlso data.DtData.Rows.Count > 0 Then
            'Hay que cachear los detalles de los códigos
            Dim ClsCodifDetalle As New CodificacionDetalle
            Dim DtDetalle As DataTable = ClsCodifDetalle.AddNew()
            Dim IntCodigo As Integer
            For Each DrCodigo As DataRow In data.DtData.Select("IDCodigo IS NOT NULL", "IDCodigo")
                If DrCodigo("IDCodigo") <> IntCodigo Then
                    Dim DtCodifDetalle As DataTable = ClsCodifDetalle.Filter(New FilterItem("IDCodigo", FilterOperator.Equal, DrCodigo("IDCodigo")))
                    If Not DtCodifDetalle Is Nothing AndAlso DtCodifDetalle.Rows.Count > 0 Then
                        For Each DrDetalle As DataRow In DtCodifDetalle.Select()
                            DtDetalle.Rows.Add(DrDetalle.ItemArray)
                        Next
                    End If
                    IntCodigo = DrCodigo("IDCodigo")
                End If
            Next
            DtDetalle.AcceptChanges()

            Dim StrIDComponente As String = String.Empty
            Dim StrNuevoCodigo As String = String.Empty
            Dim StrNuevoDescCodigo As String = String.Empty
            If Not data.DtData.Columns.Contains("Error") Then data.DtData.Columns.Add("Error", GetType(String))
            For Each DrComp As DataRow In data.DtData.Select("IDCodigo IS NOT NULL", "IDComponente")
                If DrComp("IDComponente") <> StrIDComponente Then
                    Dim StData As New Articulo.DataCodifArt
                    StData.IDCodigo = DrComp("IDCodigo")
                    StData.DtDetalle = ClsCodifDetalle.AddNew
                    For Each DrDetalle As DataRow In DtDetalle.Select("IDCodigo = " & DrComp("IDCodigo"), "IDCodigo")
                        StData.DtDetalle.Rows.Add(DrDetalle.ItemArray)
                    Next
                    StData.IDTipo = DrComp("IDTipo")
                    StData.IDFamilia = DrComp("IDFamilia")
                    If Length(DrComp("IDSubFamilia")) > 0 Then StData.IDSubFamilia = DrComp("IDSubFamilia")
                    StData.IDAnada = data.IDAnadaDestino
                    StData.DtArt = New Articulo().SelOnPrimaryKey(DrComp("IDComponente"))
                    Dim StReturn As Articulo.DataCodifReturn = ProcessServer.ExecuteTask(Of Articulo.DataCodifArt, Articulo.DataCodifReturn)(AddressOf Articulo.CodificarArticulo, StData, services)
                    StrIDComponente = DrComp("IDComponente")
                    StrNuevoCodigo = StReturn.IDArticulo
                    StrNuevoDescCodigo = StReturn.DescArticulo
                End If
                DrComp("NuevoCodigo") = StrNuevoCodigo
                DrComp("NuevoDescCodigo") = StrNuevoDescCodigo
                Dim DtArt As DataTable = New Articulo().Filter(New FilterItem("IDArticulo", FilterOperator.Equal, DrComp("NuevoCodigo")))
                If DtArt Is Nothing OrElse DtArt.Rows.Count = 0 Then
                    DrComp("Existe") = False
                Else : DrComp("Existe") = True
                End If
                If Length(DrComp("NuevoCodigo")) > 0 AndAlso Not CStr(DrComp("NuevoCodigo")).Contains("?") Then
                    Dim DrFind() As DataRow = data.DtData.Select("IDArticulo = '" & DrComp("IDArticulo") & "' AND NuevoCodigo = '" & DrComp("NuevoCodigo") & "'")
                    If DrFind.Length <= 1 Then
                        DrComp("Correcto") = True
                    Else
                        DrComp("Correcto") = False
                        DrComp("Error") = "El nuevo código de artículo: " & DrComp("NuevoCodigo") & " está repetido en el Artículo Padre: " & DrComp("IDArticulo") & "."
                        For Each DrNFind As DataRow In DrFind
                            DrNFind("Correcto") = False
                            DrNFind("Error") = "El nuevo código de artículo: " & DrComp("NuevoCodigo") & " está repetido en el Artículo Padre: " & DrComp("IDArticulo") & "."
                        Next
                    End If
                Else
                    DrComp("Correcto") = False
                    DrComp("Error") = "La codificación no está correcta para proponer el nuevo código de artículo."
                End If
            Next
        End If
        Return data.DtData
    End Function

#End Region

#End Region

End Class