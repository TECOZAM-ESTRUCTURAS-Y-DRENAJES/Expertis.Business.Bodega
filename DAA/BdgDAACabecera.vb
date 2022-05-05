Public Class BdgDAACabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbDAACabecera"

#End Region

#Region "Enumerados / Estructuras Públicas"

    Public Enum enumTipoEnvioDAA
        EMCSIntracomunitario = 1
        EMCSInterno = 2
    End Enum

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DataCont As New Contador.DatosDefaultCounterValue(data, "BdgDAACabecera", "NDaa")
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, DataCont, services)
        data("TestEMCS") = False
    End Sub

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("CodigoTipoDestino", AddressOf CambioCodigoTipoDestino)
        Obrl.Add("TipoDestinoInterno", AddressOf CambioCodigoTipoDestino)
        Obrl.Add("IDTipoDocumento", AddressOf CambioIDTipoDocumento)
        Obrl.Add("CAEExpedidor", AddressOf CambioCAEExpedidor)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambioCAEExpedidor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        'Volcamos la Descripción del CAE
        data.Current("Expedidor") = System.DBNull.Value
        If Length(data.Current("CAEExpedidor")) > 0 Then
            Dim f As New Filter
            f.Add("TipoCodigo", enumTipoCodigoEMCS.CodigosCAEExpedidor)
            f.Add("IDCodigo", data.Current("CAEExpedidor"))
            Dim dt As DataTable = New EMCSCodigos().Filter(f)
            If dt.Rows.Count > 0 Then
                data.Current("Expedidor") = dt.Rows(0)("DescCodigo") & String.Empty
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCodigoTipoDestino(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim StDirec As New StDireccionesPorCodigoTipoDestino(data.Current("CodigoTipoDestino"), Nz(data.Current("IDDireccionDestino"), Nz(data.Current("IDDireccionDestino"), 0)), Nz(data.Current("IDDireccionEntrega"), _
                                                             Nz(data.Current("IDDireccionDestino"), 0)), data.Current("TipoDestinoInterno") & String.Empty, data.Current("TipoEnvioDAA"))
        Dim Dir As DireccionesDAA = ProcessServer.ExecuteTask(Of StDireccionesPorCodigoTipoDestino, DireccionesDAA)(AddressOf DireccionesPorCodigoTipoDestino, StDirec, services)

        'Destinatario (Consignee)
        data.Current("CAEDestinatario") = Dir.CAEDestinatario
        data.Current("Destinatario") = Dir.Destinatario
        data.Current("DireccionDestinatario") = Dir.DireccionDestinatario
        data.Current("CodPostalDestinatario") = Dir.CodPostalDestinatario
        data.Current("PoblacionDestinatario") = Dir.PoblacionDestinatario
        data.Current("IDPaisDestinatario") = Dir.IDPaisDestinatario
        data.Current("ISOPaisDestinatario") = Dir.ISOPaisDestinatario
        data.Current("NIFDestinatario") = Dir.NIFDestinatario
        data.Current("CodigoAduanaExportacion") = Dir.CodigoAduanaExportacionDestinatario

        'Lugar de Entrega (Delivery Place)
        data.Current("IDCAELugarEntrega") = Dir.IDCAELugarEntrega
        data.Current("RazonSocialLugarEntrega") = Dir.RazonSocialLugarEntrega
        data.Current("DireccionLugarEntrega") = Dir.DireccionLugarEntrega
        data.Current("CodPostalLugarEntrega") = Dir.CodPostalLugarEntrega
        data.Current("PoblacionLugarEntrega") = Dir.PoblacionLugarEntrega
        data.Current("IDPaisLugarEntrega") = Dir.IDPaisLugarEntrega
        data.Current("ISOPaisLugarEntrega") = Dir.ISOPaisLugarEntrega
    End Sub

    <Task()> Public Shared Sub CambioIDTipoDocumento(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim Tipo_Documento_Albaran As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf GetTipoDocumentoAlbaran, Nothing, services)
        Dim Tipo_Documento_Factura As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf GetTipoDocumentoFactura, Nothing, services)

        Dim IDTipoDocumentoAnt As String = data.Current("IDTipoDocumento") & String.Empty
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDTipoDocumento")) > 0 Then
            Select Case data.Current("IDTipoDocumento")
                Case Tipo_Documento_Albaran
                    If IDTipoDocumentoAnt = Tipo_Documento_Factura Then
                        '//Si veníamos de una Factura, buscamos cual era su Albarán
                        'NumeroDocumento
                        'FechaDocumento

                        If Length(data.Current("NumeroDocumento")) > 0 Then
                            Dim NFactura As String = data.Current("NumeroDocumento")
                            Dim IDBaseDatos As Guid
                            data.Current("NumeroDocumento") = System.DBNull.Value
                            data.Current("FechaDocumento") = System.DBNull.Value

                            If Not IsDBNull(data.Current("IDDAA")) AndAlso Not CType(data.Current("IDDAA"), Guid).Equals(Guid.Empty) Then
                                Dim fDaa As New Filter
                                fDaa.Add(New GuidFilterItem("IDDAA", data.Current("IDDAA")))
                                fDaa.Add(New IsNullFilterItem("NAlbaran", False))
                                Dim dtOrigenes As DataTable = New BdgDAAOrigenes().Filter(fDaa)
                                If dtOrigenes.Rows.Count > 0 Then
                                    Dim IDBaseDatosCurrent As Guid = AdminData.GetConnectionInfo.IDDataBase

                                    For Each drOrigen As DataRow In dtOrigenes.Rows
                                        IDBaseDatos = drOrigen("IDBaseDatos")
                                        Dim IDAlbaran As Integer = drOrigen("IDAlbaran")

                                        Try
                                            If IDBaseDatosCurrent <> IDBaseDatos Then
                                                AdminData.CommitTx(True)
                                                AdminData.SetCurrentConnection(IDBaseDatos)
                                                AdminData.BeginTx()
                                            End If

                                            Dim f As New Filter
                                            f.Add(New NumberFilterItem("IDAlbaran", IDAlbaran))
                                            f.Add(New StringFilterItem("NFactura", NFactura))
                                            Dim dtAlbFra As DataTable = New BE.DataEngine().Filter("vNegAlbaranFacturaVenta", f)
                                            If dtAlbFra.Rows.Count > 0 Then
                                                data.Current("NumeroDocumento") = dtAlbFra.Rows(0)("NAlbaran")
                                                data.Current("FechaDocumento") = dtAlbFra.Rows(0)("FechaAlbaran")
                                            End If
                                        Catch ex As Exception

                                        Finally
                                            If IDBaseDatosCurrent <> IDBaseDatos Then
                                                AdminData.CommitTx(True)
                                                AdminData.SetCurrentConnection(IDBaseDatosCurrent)
                                                AdminData.BeginTx()
                                            End If
                                        End Try
                                    Next
                                End If
                            End If
                        End If
                    End If
                Case Tipo_Documento_Factura
                    If IDTipoDocumentoAnt = Tipo_Documento_Albaran Then
                        '//Si veníamos de un Albarán, buscamos cual era su Factura
                        'NumeroDocumento
                        'FechaDocumento
                        If Length(data.Current("NumeroDocumento")) > 0 Then
                            Dim NAlbaran As String = data.Current("NumeroDocumento")
                            Dim IDBaseDatos As Guid
                            Dim IDAlbaran As Integer
                            data.Current("NumeroDocumento") = System.DBNull.Value
                            data.Current("FechaDocumento") = System.DBNull.Value

                            If Not IsDBNull(data.Current("IDDAA")) AndAlso Not CType(data.Current("IDDAA"), Guid).Equals(Guid.Empty) Then
                                Dim fDaa As New Filter
                                fDaa.Add(New GuidFilterItem("IDDAA", data.Current("IDDAA")))
                                fDaa.Add(New StringFilterItem("NAlbaran", NAlbaran))
                                Dim dtOrigenes As DataTable = New BdgDAAOrigenes().Filter(fDaa)
                                If dtOrigenes.Rows.Count > 0 Then
                                    IDBaseDatos = dtOrigenes.Rows(0)("IDBaseDatos")
                                    IDAlbaran = dtOrigenes.Rows(0)("IDAlbaran")
                                End If

                                If Not IsDBNull(IDBaseDatos) AndAlso Not CType(IDBaseDatos, Guid).Equals(Guid.Empty) Then
                                    Dim IDBaseDatosCurrent As Guid = AdminData.GetConnectionInfo.IDDataBase
                                    Try
                                        If IDBaseDatosCurrent <> IDBaseDatos Then
                                            AdminData.CommitTx(True)
                                            AdminData.SetCurrentConnection(IDBaseDatos)
                                            AdminData.BeginTx()
                                        End If

                                        Dim f As New Filter
                                        f.Add(New NumberFilterItem("IDAlbaran", IDAlbaran))
                                        Dim dtAlbFra As DataTable = New BE.DataEngine().Filter("vNegAlbaranFacturaVenta", f)
                                        If dtAlbFra.Rows.Count > 0 Then
                                            data.Current("NumeroDocumento") = dtAlbFra.Rows(0)("NFactura")
                                            data.Current("FechaDocumento") = dtAlbFra.Rows(0)("FechaFactura")
                                        End If
                                    Catch ex As Exception

                                    Finally
                                        If IDBaseDatosCurrent <> IDBaseDatos Then
                                            AdminData.CommitTx(True)
                                            AdminData.SetCurrentConnection(IDBaseDatosCurrent)
                                            AdminData.BeginTx()
                                        End If
                                    End Try
                                End If
                            End If
                        End If
                    End If
            End Select
        End If
    End Sub

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarReference)
        deleteProcess.AddTask(Of DataRow)(AddressOf GuardarOrigenes)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf BorrarDAAsAlbaranesPedidos)
    End Sub

    <Task()> Public Shared Sub ValidarReference(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("AadReferenceCode")) > 0 AndAlso Strings.InStr(UCase(data("AadReferenceCode")), "TEST") < 0 Then
            ApplicationService.GenerateError("No se puede borrar el DAA. Se ha enviado vía web a AEAT.")
        End If
    End Sub

    <Task()> Public Shared Sub GuardarOrigenes(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim fDaa As New Filter
        fDaa.Add(New GuidFilterItem("IDDAA", data("IDDAA")))
        Dim dtDaaOrigenes As DataTable = New BdgDAAOrigenes().Filter(fDaa)

        Dim lstBBDD As OrigenesDAA = services.GetService(Of OrigenesDAA)()
        lstBBDD.BBDD = (From c In dtDaaOrigenes Select c("IDBaseDatos") Distinct).ToList
    End Sub

    <Task()> Public Shared Sub BorrarDAAsAlbaranesPedidos(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim fDaa As New Filter
        fDaa.Add(New GuidFilterItem("IDDAA", data("IDDAA")))

        Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase
        Dim lstBBDD As OrigenesDAA = services.GetService(Of OrigenesDAA)()
        If Not lstBBDD Is Nothing AndAlso Not lstBBDD.BBDD Is Nothing AndAlso lstBBDD.BBDD.Count > 0 Then
            For Each IDDAABaseDatos As Guid In lstBBDD.BBDD
                Try
                    If currentBBDD <> IDDAABaseDatos Then
                        AdminData.SetCurrentConnection(IDDAABaseDatos)
                        AdminData.CommitTx(True)
                    End If

                    Dim dtAVC As DataTable = New AlbaranVentaCabecera().Filter(fDaa)
                    For Each rwAVC As DataRow In dtAVC.Rows
                        rwAVC("IDDAA") = DBNull.Value
                        rwAVC("NDAA") = DBNull.Value
                        rwAVC("IDDAABaseDatos") = DBNull.Value
                        rwAVC("AadReferenceCode") = DBNull.Value
                    Next
                    BusinessHelper.UpdateTable(dtAVC)


                    Dim dtPVC As DataTable = New PedidoVentaCabecera().Filter(fDaa)
                    For Each rwPVC As DataRow In dtPVC.Rows
                        rwPVC("IDDAA") = DBNull.Value
                        rwPVC("NDAA") = DBNull.Value
                        rwPVC("IDDAABaseDatos") = DBNull.Value
                        rwPVC("AadReferenceCode") = DBNull.Value
                    Next
                    BusinessHelper.UpdateTable(dtPVC)

                Catch ex As Exception
                    If currentBBDD <> IDDAABaseDatos Then
                        AdminData.RollBackTx()
                        AdminData.SetCurrentConnection(currentBBDD)
                        AdminData.CommitTx(True)
                    End If

                Finally
                    If currentBBDD <> IDDAABaseDatos Then
                        AdminData.SetCurrentConnection(currentBBDD)
                        AdminData.CommitTx(True)
                    End If

                End Try
            Next
        End If

    End Sub

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClavePrimaria)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarNContador)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarDAAOrigenes)
    End Sub

    <Task()> Public Shared Sub AsignarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDDAA")) = 0 Then data("IDDAA") = Guid.NewGuid
        End If
    End Sub

    <Task()> Public Shared Sub AsignarNContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDContador")) > 0 AndAlso data.IsNull("NDaa") Then
                data("NDaa") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, data("IDContador"), services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarDAAOrigenes(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified AndAlso Not Nz(data("TestEMCS"), False) AndAlso data("AadReferenceCode") & String.Empty <> data("AadReferenceCode", DataRowVersion.Original) & String.Empty Then
            Dim dtOrigenes As DataTable

            Dim DAAOrig As New BdgDAAOrigenes
            Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase

            Dim fDaa As New Filter
            fDaa.Add(New GuidFilterItem("IDDAA", data("IDDAA")))
            Dim dtOrigen As DataTable = DAAOrig.Filter(fDaa)
            If dtOrigen.Rows.Count > 0 Then
                For Each drOrigen As DataRow In dtOrigen.Select(Nothing, "IDBaseDatos")
                    Try
                        If drOrigen("IDBaseDatos") <> currentBBDD Then
                            AdminData.SetCurrentConnection(drOrigen("IDBaseDatos"))
                            AdminData.CommitTx(True)
                        End If

                        If Nz(drOrigen("IDPedido"), 0) <> 0 Then
                            Dim f As New Filter
                            f.Add(New GuidFilterItem("IDDAA", drOrigen("IDDAA")))
                            f.Add(New NumberFilterItem("IDPedido", drOrigen("IDPedido")))
                            Dim dttPedido As DataTable = New PedidoVentaCabecera().Filter(f)
                            If Not dttPedido Is Nothing AndAlso dttPedido.Rows.Count = 1 Then
                                dttPedido.Rows(0)("AadReferenceCode") = data("AadReferenceCode") & String.Empty
                                BusinessHelper.UpdateTable(dttPedido)
                            End If
                        End If

                        If Nz(drOrigen("IDAlbaran"), 0) <> 0 Then
                            Dim f As New Filter
                            f.Add(New GuidFilterItem("IDDAA", drOrigen("IDDAA")))
                            f.Add(New NumberFilterItem("IDAlbaran", drOrigen("IDAlbaran")))
                            Dim dttAlbaran As DataTable = New AlbaranVentaCabecera().Filter(f)
                            If Not dttAlbaran Is Nothing AndAlso dttAlbaran.Rows.Count = 1 Then
                                dttAlbaran.Rows(0)("AadReferenceCode") = data("AadReferenceCode") & String.Empty
                                BusinessHelper.UpdateTable(dttAlbaran)
                            End If
                        End If

                    Catch ex As Exception
                        If drOrigen("IDBaseDatos") <> currentBBDD Then
                            AdminData.RollBackTx()
                            AdminData.SetCurrentConnection(currentBBDD)
                            AdminData.CommitTx(True)
                        End If

                    Finally
                        If drOrigen("IDBaseDatos") <> currentBBDD Then
                            AdminData.SetCurrentConnection(currentBBDD)
                            AdminData.CommitTx(True)
                        End If
                    End Try
                Next
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function TipoDAAPorPais(ByVal IDDireccion As Long, ByVal services As ServiceProvider) As TipoDAA
        Dim StReturn As New TipoDAA
        StReturn = TipoDAA.Nacional

        Dim oCliDi As Negocio.ClienteDireccion = New Negocio.ClienteDireccion
        Dim rwDirEnvioAV As DataRow = oCliDi.GetItemRow(IDDireccion)
        If Length(rwDirEnvioAV("IDPais")) > 0 Then
            Dim dtPais As DataTable = New Pais().SelOnPrimaryKey(rwDirEnvioAV("IDPais"))
            If dtPais.Rows.Count > 0 Then
                If CBool(dtPais.Rows(0)("CEE")) And CBool(dtPais.Rows(0)("Extranjero")) Then
                    StReturn = TipoDAA.CEE
                ElseIf CBool(dtPais.Rows(0)("Extranjero")) Then
                    StReturn = TipoDAA.Exportacion
                Else
                    StReturn = TipoDAA.Nacional
                End If
            End If
        Else
            ApplicationService.GenerateError("No se ha especificado un País en la Dirección de Envío del Albarán.")
        End If

        Return StReturn
    End Function

    <Serializable()> _
    Public Class StDireccionesPorCodigoTipoDestino
        Public CodigoTipoDestino As String
        Public IDDestino As Integer
        Public IDEntrega As Integer
        Public TipoDestinoInterno As String
        Public TipoEnvioDAA As enumTipoEnvioDAA

        Public Sub New()
        End Sub

        Public Sub New(ByVal CodigoTipoDestino As String, ByVal IDDestino As Integer, ByVal IDEntrega As Integer, ByVal TipoDestinoInterno As String, ByVal TipoEnvioDaa As enumTipoEnvioDAA)
            Me.CodigoTipoDestino = CodigoTipoDestino
            Me.IDDestino = IDDestino
            Me.IDEntrega = IDEntrega
            Me.TipoDestinoInterno = TipoDestinoInterno
            Me.TipoEnvioDAA = TipoEnvioDaa
        End Sub
    End Class

    <Task()> Public Shared Function DireccionesPorCodigoTipoDestino(ByVal data As StDireccionesPorCodigoTipoDestino, ByVal services As ServiceProvider) As DireccionesDAA

        Dim Dir As New DireccionesDAA
        Dim clsPais As New Pais

        Dim rwDirDestino As DataRow = New Negocio.ClienteDireccion().GetItemRow(data.IDDestino)
        Dim stISOPaisDestino As String = String.Empty
        Dim dtPaisDestino As DataTable = clsPais.SelOnPrimaryKey(rwDirDestino("IDPais") & String.Empty)
        If dtPaisDestino.Rows.Count = 1 Then
            stISOPaisDestino = dtPaisDestino.Rows(0)("CodigoISO") & String.Empty
        End If

        If data.TipoEnvioDAA = enumTipoEnvioDAA.EMCSInterno Then

            'Destinatario (Consignee)
            If data.TipoDestinoInterno = "04" Then
                Dir.Destinatario = String.Empty
                Dir.DireccionDestinatario = String.Empty
                Dir.CodPostalDestinatario = String.Empty
                Dir.PoblacionDestinatario = String.Empty
                Dir.IDPaisDestinatario = String.Empty
                Dir.ISOPaisDestinatario = String.Empty
                Dir.NIFDestinatario = String.Empty

                Dir.AutoridadAgroalimentariaDestinatario = String.Empty
                Dir.CodigoAduanaExportacionDestinatario = String.Empty
            Else
                Dir.Destinatario = rwDirDestino("RazonSocial") & String.Empty
                Dir.DireccionDestinatario = rwDirDestino("Direccion") & String.Empty
                Dir.CodPostalDestinatario = rwDirDestino("CodPostal") & String.Empty
                Dir.PoblacionDestinatario = rwDirDestino("Poblacion") & String.Empty
                Dir.IDPaisDestinatario = rwDirDestino("IDPais") & String.Empty
                Dir.ISOPaisDestinatario = stISOPaisDestino
                Dir.NIFDestinatario = rwDirDestino("CifCliente") & String.Empty

                Dir.AutoridadAgroalimentariaDestinatario = rwDirDestino("EMCSAutoridadAgroalimentaria") & String.Empty
                Dir.CodigoAduanaExportacionDestinatario = rwDirDestino("EMCSAduanaExportacion") & String.Empty
            End If
            If data.TipoDestinoInterno = "01" Then
                Dir.CAEDestinatario = rwDirDestino("IDCAE") & String.Empty
            Else
                Dir.CAEDestinatario = String.Empty
            End If

            'Lugar de Entrega (Delivery Place)
            Dir.IDCAELugarEntrega = String.Empty                            'No se cumplimenta
            Dir.RazonSocialLugarEntrega = String.Empty                      'No se cumplimenta
            Dir.DireccionLugarEntrega = String.Empty                        'No se cumplimenta
            Dir.CodPostalLugarEntrega = String.Empty                        'No se cumplimenta
            Dir.PoblacionLugarEntrega = String.Empty                        'No se cumplimenta
            Dir.IDPaisLugarEntrega = String.Empty                           'No se cumplimenta
            Dir.ISOPaisLugarEntrega = String.Empty                           'No se cumplimenta
        Else
            Dim rwDirEntrega As DataRow = New Negocio.ClienteDireccion().GetItemRow(data.IDEntrega)
            Dim stISOPaisEntrega As String = String.Empty
            Dim dtPaisEntrega As DataTable = clsPais.SelOnPrimaryKey(rwDirEntrega("IDPais") & String.Empty)
            If dtPaisEntrega.Rows.Count = 1 Then
                stISOPaisEntrega = dtPaisEntrega.Rows(0)("CodigoISO") & String.Empty
            End If

            Select Case data.CodigoTipoDestino
                Case BdgCodigoTipoDestino.DepositoFiscal, BdgCodigoTipoDestino.DestinatarioRegistrado, BdgCodigoTipoDestino.DestinatarioRegistradoOcasional
                    'Destinatario (Consignee)
                    Dir.CAEDestinatario = rwDirDestino("IDCAE") & String.Empty
                    Dir.Destinatario = rwDirDestino("RazonSocial") & String.Empty
                    Dir.DireccionDestinatario = rwDirDestino("Direccion") & String.Empty
                    Dir.CodPostalDestinatario = rwDirDestino("CodPostal") & String.Empty
                    Dir.PoblacionDestinatario = rwDirDestino("Poblacion") & String.Empty
                    Dir.IDPaisDestinatario = rwDirDestino("IDPais") & String.Empty
                    Dir.ISOPaisDestinatario = stISOPaisDestino

                    Dir.AutoridadAgroalimentariaDestinatario = rwDirDestino("EMCSAutoridadAgroalimentaria") & String.Empty
                    Dir.CodigoAduanaExportacionDestinatario = String.Empty

                    'Lugar de Entrega (Delivery Place)
                    'Para DestinatarioRegistrado y DestinatarioRegistradoOcasional son opcionales
                    Dir.IDCAELugarEntrega = rwDirEntrega("IDCAE") & String.Empty
                    Dir.RazonSocialLugarEntrega = rwDirEntrega("RazonSocial") & String.Empty
                    Dir.DireccionLugarEntrega = rwDirEntrega("Direccion") & String.Empty
                    Dir.CodPostalLugarEntrega = rwDirEntrega("CodPostal") & String.Empty
                    Dir.PoblacionLugarEntrega = rwDirEntrega("Poblacion") & String.Empty
                    Dir.IDPaisLugarEntrega = rwDirEntrega("IDPais") & String.Empty
                    Dir.ISOPaisLugarEntrega = stISOPaisEntrega

                    Dir.AutoridadAgroalimentariaLugarEntrega = rwDirEntrega("EMCSAutoridadAgroalimentaria") & String.Empty
                    Dir.CodigoAduanaExportacionLugarEntrega = String.Empty

                Case BdgCodigoTipoDestino.EntregaDirectaAutorizada
                    'Destinatario (Consignee)
                    Dir.CAEDestinatario = rwDirDestino("IDCAE") & String.Empty
                    Dir.Destinatario = rwDirDestino("RazonSocial") & String.Empty
                    Dir.DireccionDestinatario = rwDirDestino("Direccion") & String.Empty
                    Dir.CodPostalDestinatario = rwDirDestino("CodPostal") & String.Empty
                    Dir.PoblacionDestinatario = rwDirDestino("Poblacion") & String.Empty
                    Dir.IDPaisDestinatario = rwDirDestino("IDPais") & String.Empty
                    Dir.ISOPaisDestinatario = stISOPaisDestino

                    Dir.AutoridadAgroalimentariaDestinatario = rwDirDestino("EMCSAutoridadAgroalimentaria") & String.Empty
                    Dir.CodigoAduanaExportacionDestinatario = String.Empty

                    'Lugar de Entrega (Delivery Place)
                    Dir.IDCAELugarEntrega = String.Empty                                          'No se cumplimenta"
                    Dir.RazonSocialLugarEntrega = rwDirEntrega("RazonSocial") & String.Empty       'Opcional
                    Dir.DireccionLugarEntrega = rwDirEntrega("Direccion") & String.Empty
                    Dir.CodPostalLugarEntrega = rwDirEntrega("CodPostal") & String.Empty
                    Dir.PoblacionLugarEntrega = rwDirEntrega("Poblacion") & String.Empty
                    Dir.IDPaisLugarEntrega = rwDirEntrega("IDPais") & String.Empty
                    Dir.ISOPaisLugarEntrega = stISOPaisEntrega

                    Dir.AutoridadAgroalimentariaLugarEntrega = rwDirEntrega("EMCSAutoridadAgroalimentaria") & String.Empty
                    Dir.CodigoAduanaExportacionLugarEntrega = String.Empty

                Case BdgCodigoTipoDestino.DestinatarioExento
                    'Destinatario (Consignee)
                    Dir.CAEDestinatario = String.Empty                                            'No se cumplimenta"
                    Dir.Destinatario = rwDirDestino("RazonSocial") & String.Empty
                    Dir.DireccionDestinatario = rwDirDestino("Direccion") & String.Empty
                    Dir.CodPostalDestinatario = rwDirDestino("CodPostal") & String.Empty
                    Dir.PoblacionDestinatario = rwDirDestino("Poblacion") & String.Empty
                    Dir.IDPaisDestinatario = rwDirDestino("IDPais") & String.Empty
                    Dir.ISOPaisDestinatario = stISOPaisDestino

                    Dir.AutoridadAgroalimentariaDestinatario = rwDirDestino("EMCSAutoridadAgroalimentaria") & String.Empty
                    Dir.CodigoAduanaExportacionDestinatario = String.Empty

                    'Lugar de Entrega (Delivery Place)
                    Dir.IDCAELugarEntrega = rwDirEntrega("IDCAE") & String.Empty                  'Opcional
                    Dir.RazonSocialLugarEntrega = rwDirEntrega("RazonSocial") & String.Empty      'Opcional
                    Dir.DireccionLugarEntrega = rwDirEntrega("Direccion") & String.Empty          'Opcional
                    Dir.CodPostalLugarEntrega = rwDirEntrega("CodPostal") & String.Empty          'Opcional
                    Dir.PoblacionLugarEntrega = rwDirEntrega("Poblacion") & String.Empty          'Opcional
                    Dir.IDPaisLugarEntrega = rwDirEntrega("IDPais") & String.Empty                'Opcional
                    Dir.ISOPaisLugarEntrega = stISOPaisEntrega                                    'Opcional

                    Dir.AutoridadAgroalimentariaLugarEntrega = rwDirEntrega("EMCSAutoridadAgroalimentaria") & String.Empty
                    Dir.CodigoAduanaExportacionLugarEntrega = String.Empty

                Case BdgCodigoTipoDestino.Exportacion
                    'Destinatario (Consignee)
                    Dir.CAEDestinatario = String.Empty              'Opcional. Si se cumplimenta llevará el NIF del Agente de Aduana.
                    Dir.Destinatario = rwDirDestino("RazonSocial") & String.Empty
                    Dir.DireccionDestinatario = rwDirDestino("Direccion") & String.Empty
                    Dir.CodPostalDestinatario = rwDirDestino("CodPostal") & String.Empty
                    Dir.PoblacionDestinatario = rwDirDestino("Poblacion") & String.Empty
                    Dir.IDPaisDestinatario = rwDirDestino("IDPais") & String.Empty
                    Dir.ISOPaisDestinatario = stISOPaisDestino

                    Dir.AutoridadAgroalimentariaDestinatario = rwDirDestino("EMCSAutoridadAgroalimentaria") & String.Empty
                    Dir.CodigoAduanaExportacionDestinatario = rwDirDestino("EMCSAduanaExportacion") & String.Empty

                    'Lugar de Entrega (Delivery Place)
                    Dir.IDCAELugarEntrega = String.Empty                            'No se cumplimenta
                    Dir.RazonSocialLugarEntrega = String.Empty                      'No se cumplimenta
                    Dir.DireccionLugarEntrega = String.Empty                        'No se cumplimenta
                    Dir.CodPostalLugarEntrega = String.Empty                        'No se cumplimenta
                    Dir.PoblacionLugarEntrega = String.Empty                        'No se cumplimenta
                    Dir.IDPaisLugarEntrega = String.Empty                           'No se cumplimenta
                    Dir.ISOPaisLugarEntrega = String.Empty                          'No se cumplimenta

                    Dir.AutoridadAgroalimentariaLugarEntrega = String.Empty
                    Dir.CodigoAduanaExportacionLugarEntrega = String.Empty

                Case BdgCodigoTipoDestino.DestinoDesconocido
                    'Destinatario (Consignee)
                    Dir.CAEDestinatario = String.Empty                             'No se cumplimenta
                    Dir.Destinatario = String.Empty                                'No se cumplimenta
                    Dir.DireccionDestinatario = String.Empty                       'No se cumplimenta
                    Dir.CodPostalDestinatario = String.Empty                       'No se cumplimenta
                    Dir.PoblacionDestinatario = String.Empty                       'No se cumplimenta
                    Dir.IDPaisDestinatario = String.Empty                          'No se cumplimenta
                    Dir.ISOPaisDestinatario = String.Empty                         'No se cumplimenta

                    Dir.AutoridadAgroalimentariaDestinatario = String.Empty
                    Dir.CodigoAduanaExportacionDestinatario = String.Empty

                    'Lugar de Entrega (Delivery Place)
                    Dir.IDCAELugarEntrega = String.Empty                           'No se cumplimenta
                    Dir.RazonSocialLugarEntrega = String.Empty                     'No se cumplimenta
                    Dir.DireccionLugarEntrega = String.Empty                       'No se cumplimenta
                    Dir.CodPostalLugarEntrega = String.Empty                       'No se cumplimenta
                    Dir.PoblacionLugarEntrega = String.Empty                       'No se cumplimenta
                    Dir.IDPaisLugarEntrega = String.Empty                          'No se cumplimenta
                    Dir.ISOPaisLugarEntrega = String.Empty                         'No se cumplimenta

                    Dir.AutoridadAgroalimentariaLugarEntrega = String.Empty
                    Dir.CodigoAduanaExportacionLugarEntrega = String.Empty

            End Select
        End If
        Return Dir
    End Function


#End Region

#Region "DAA"

#Region "Creación DAAS"

    Public Const CN_VistaDAAAlbaran As String = "NegBdgDAAAlbaranLineas"
    Public Const CN_VistaDAAPedido As String = "NegBdgDAAPedidoLineas"

    Public Enum enumOrigenDAA
        Manual
        Albaran
        Pedido
    End Enum


    <Task()> Public Shared Function CrearDAAAlbaran(ByVal datas As stDatosDAADesdeAlbaran, ByVal services As ServiceProvider) As stCrearDAAInfoResult
        Dim result As stCrearDAAInfoResult
        If Not datas.IDDAA.Equals(Guid.Empty) Then
            result = New stCrearDAAInfoResult
            result.IDDAA = datas.IDDAA
            Dim dtDAA As DataTable = New BdgDAACabecera().SelOnPrimaryKey(datas.IDDAA)
            If dtDAA.Rows.Count > 0 Then
                result.NDAA = dtDAA.Rows(0)("NDAA") & String.Empty
            End If
        Else
            Dim dataDAA As stCrearDAAInfo = ProcessServer.ExecuteTask(Of stDatosDAADesdeAlbaran, stCrearDAAInfo)(AddressOf PrepararDatosDAADesdeAlbaran, datas, services)
            result = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfoResult)(AddressOf CrearDAA, dataDAA, services)
        End If

        If Not result Is Nothing AndAlso Length(result.NDAA) > 0 Then
            Dim bsnAlb As New AlbaranVentaCabecera

            Dim blnError As Boolean
            Dim blnExisteEnOrigen As Boolean
            Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase
            For Each dtrRegistroEmpresa As DataRow In datas.RegistrosEmpresas.Listado.Rows
                Dim NAlbaran As String = String.Empty
                Try
                    blnError = False
                    blnExisteEnOrigen = False
                    If currentBBDD <> dtrRegistroEmpresa("IDBaseDatos") Then
                        AdminData.SetCurrentConnection(dtrRegistroEmpresa("IDBaseDatos"))
                        AdminData.CommitTx(True)
                    End If
                    
                    Dim dttAlbaran As DataTable = bsnAlb.SelOnPrimaryKey(dtrRegistroEmpresa("IDRegistro"))
                    If Not dttAlbaran Is Nothing AndAlso dttAlbaran.Rows.Count = 1 Then
                        dttAlbaran.Rows(0)("IDDAA") = result.IDDAA
                        dttAlbaran.Rows(0)("NDAA") = result.NDAA
                        dttAlbaran.Rows(0)("IDDAABaseDatos") = currentBBDD
                        NAlbaran = dttAlbaran.Rows(0)("NAlbaran")
                        bsnAlb.Update(dttAlbaran)
                        blnExisteEnOrigen = True
                    End If
                Catch ex As Exception
                    If currentBBDD <> dtrRegistroEmpresa("IDBaseDatos") Then
                        AdminData.RollBackTx()
                        AdminData.SetCurrentConnection(currentBBDD)
                        AdminData.CommitTx(True)
                    End If
                    blnError = True
                Finally
                    If currentBBDD <> dtrRegistroEmpresa("IDBaseDatos") Then
                        AdminData.SetCurrentConnection(currentBBDD)
                        AdminData.CommitTx(True)
                    End If

                    If Not blnError AndAlso blnExisteEnOrigen Then
                        Dim BdgOrigen As New BdgDAAOrigenes
                        Dim dtBdgOrigen As DataTable = BdgOrigen.AddNew
                        Dim drNewOrigen As DataRow = dtBdgOrigen.NewRow
                        drNewOrigen("IDDAAOrigen") = Guid.NewGuid
                        drNewOrigen("IDDAA") = result.IDDAA
                        drNewOrigen("IDAlbaran") = dtrRegistroEmpresa("IDRegistro")
                        drNewOrigen("NAlbaran") = NAlbaran
                        drNewOrigen("IDBaseDatos") = dtrRegistroEmpresa("IDBaseDatos")
                        dtBdgOrigen.Rows.Add(drNewOrigen)
                        BdgOrigen.Update(dtBdgOrigen)
                    End If
                End Try
            Next
        End If
        Return result
    End Function

    <Task()> Public Shared Function CrearDAAPedido(ByVal data As stDatosDAADesdePedido, ByVal services As ServiceProvider) As stCrearDAAInfoResult
        Dim dataResult As stCrearDAAInfo = ProcessServer.ExecuteTask(Of stDatosDAADesdePedido, stCrearDAAInfo)(AddressOf PrepararDatosDAADesdePedido, data, services)
        Dim result As stCrearDAAInfoResult = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfoResult)(AddressOf CrearDAA, dataResult, services)
        'TODO => ACTUALIZAR INFO EN TODOS LOS REGISTROS
        If Not result Is Nothing AndAlso Length(result.NDAA) > 0 Then
            Dim bsnPedido As New PedidoVentaCabecera
            Dim blnError As Boolean
            Dim blnExisteEnOrigen As Boolean
            Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase
            For Each dtrRegistroEmpresa As DataRow In data.RegistrosEmpresas.Listado.Rows
                Dim NPedido As String = String.Empty

                Try
                    blnError = False
                    blnExisteEnOrigen = False

                    If currentBBDD <> dtrRegistroEmpresa("IDBaseDatos") Then
                        AdminData.SetCurrentConnection(dtrRegistroEmpresa("IDBaseDatos"))
                        AdminData.CommitTx(True)
                    End If

                    Dim dttPedido As DataTable = bsnPedido.SelOnPrimaryKey(dtrRegistroEmpresa("IDRegistro"))
                    If Not dttPedido Is Nothing AndAlso dttPedido.Rows.Count = 1 Then
                        dttPedido.Rows(0)("IDDAA") = result.IDDAA
                        dttPedido.Rows(0)("NDAA") = result.NDAA
                        dttPedido.Rows(0)("IDDAABaseDatos") = currentBBDD
                        bsnPedido.Update(dttPedido)
                        NPedido = dttPedido.Rows(0)("NPedido")
                        blnExisteEnOrigen = True
                    End If

                Catch ex As Exception
                    If currentBBDD <> dtrRegistroEmpresa("IDBaseDatos") Then
                        AdminData.RollBackTx()
                        AdminData.SetCurrentConnection(currentBBDD)
                        AdminData.CommitTx(True)
                    End If

                    blnError = True

                Finally
                    If currentBBDD <> dtrRegistroEmpresa("IDBaseDatos") Then
                        AdminData.SetCurrentConnection(currentBBDD)
                        AdminData.CommitTx(True)
                    End If

                    If Not blnError AndAlso blnExisteEnOrigen Then
                        Dim BdgOrigen As New BdgDAAOrigenes
                        Dim dtBdgOrigen As DataTable = BdgOrigen.AddNew
                        Dim drNewOrigen As DataRow = dtBdgOrigen.NewRow
                        drNewOrigen("IDDAAOrigen") = Guid.NewGuid
                        drNewOrigen("IDDAA") = result.IDDAA
                        drNewOrigen("IDPedido") = dtrRegistroEmpresa("IDRegistro")
                        drNewOrigen("NPedido") = NPedido
                        drNewOrigen("IDBaseDatos") = dtrRegistroEmpresa("IDBaseDatos")
                        dtBdgOrigen.Rows.Add(drNewOrigen)
                        BdgOrigen.Update(dtBdgOrigen)
                    End If
                End Try
            Next
        End If
        Return result
    End Function

    <Task()> Public Shared Function CrearDAA(ByVal dataDAA As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfoResult
        AdminData.BeginTx()

        Dim result As New stCrearDAAInfoResult
        result.Source = dataDAA.NumeroDocumento

        'Try
        '*********************************************************************************************************************************'
        '**** CABECERA COMUN INTRACOMUNITARIO****'
        '********************************************************************************************************************************'

        'BASE
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraBase, dataDAA, services)

        '5 - DESTINATARIO
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraDestinatario, dataDAA, services)

        '2 - EXPEDIDOR
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraExpedidor, dataDAA, services)

        '3 - LUGARDESPACHO
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraLugarExpedicion, dataDAA, services)

        '4 - OFICINA IMPORTACIÓN
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraOficinaImportacion, dataDAA, services)

        '6 - MIEMBRO - COMPLEMENT CONSIGNEE
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraConsignatario, dataDAA, services)

        '7 - LUGAR DE ENTREGA
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraLugarEntrega, dataDAA, services)

        '8 - D PLACE CUSTOMS OFFICE 
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraOficinaExportacion, dataDAA, services)

        '10 - DISPATCH OFFICE
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraOficinaExpedicion, dataDAA, services)

        '14 - RESPONSABLE TRANSPORTE
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraResponsableTransporte, dataDAA, services)

        '15 - TRANSPORTISTA
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraTransportista, dataDAA, services)

        '16.- Transport Details  OK
        ' tbDaaDetalleTransporte
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraDetalleTransporte, dataDAA, services)

        '18.- Document Certificate OK
        ' tbDaaDetalleDocumento
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraDetalleDocumento, dataDAA, services)

        '13 - TRANSPORT MODE y 1.-Header E-AAD (JUNTO CON 13 EN MO) OK
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraExtraTransporte, dataDAA, services)

        '11 - TIPO GARANTE y 12 - GARANTE
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraGarantia, dataDAA, services)

        '9 E-AAD DRAFT
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraExtra, dataDAA, services)

        '*********************************************************************************************************************************'
        '**** CABECERA INTERNO ****'
        '*********************************************************************************************************************************'

        '6 INTERNO - ORGANISMO EXENTO
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraOrganismoExentoInterno, dataDAA, services)

        '11 INTERNO - AUTORIDAD AGROALIMENTARIA
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAACabeceraAutoridadAgroAlimentariaInterno, dataDAA, services)

        Dim bdgDAA As New BdgDAACabecera
        bdgDAA.Update(dataDAA.Cabecera)

        '*********************************************************************************************************************************'
        '**** LÍNEAS ****'
        '*********************************************************************************************************************************'

        '17 - LÍNEAS
        dataDAA = ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf CrearDAALineas, dataDAA, services)

        'TABLAS AUXILIARES: LINEAS, DOCUMENTO, ETC
        Dim bdgDAAL As New BdgDAALinea
        bdgDAAL.Update(dataDAA.Lineas)

        If Not dataDAA.DetalleTransporte Is Nothing AndAlso dataDAA.DetalleTransporte.Rows.Count > 0 Then
            Dim bdgDAAT As New BdgDAADetalleTransporte
            bdgDAAT.Update(dataDAA.DetalleTransporte)
        End If


        If Not dataDAA.LineasPaquete Is Nothing AndAlso dataDAA.LineasPaquete.Rows.Count > 0 Then
            Dim bdgDAALP As New BdgDAALineaPaquete
            bdgDAALP.Update(dataDAA.LineasPaquete)
        End If


        If Not dataDAA.LineasOperacion Is Nothing AndAlso dataDAA.LineasOperacion.Rows.Count > 0 Then
            Dim bdgDAALT As New BdgDAALineaOperacion
            bdgDAALT.Update(dataDAA.LineasOperacion)
        End If

        result.NDAA = dataDAA.Cabecera.Rows(0)("NDAA")
        result.IDDAA = dataDAA.Cabecera.Rows(0)("IDDAA")

        'Catch ex As Exception
        '    result.ErrorMessage = ex.Message
        'End Try
        Return result
    End Function

    <Serializable()> _
    Public Class stCrearDAAInfo

#Region "           Orígenes de datos"

        Public Contador As String

        Public Cliente As String
        Public Direccion As String
        Public DireccionEntrega As String
        Public DireccionDestino As String
        Public FormaEnvio As String
        Public Matricula As String
        Public Remolque As String
        Public Precinto As String
        Public Contenedor As String

        Public EstadoMiembro As String
        Public NumeroCertificado As String
        Public Aduana As String

        Public NIFResponsableTransporte As String
        Public ResponsableTransporte As String

        Public NIFTransportista As String
        Public Transportista As String

        Public IDModoTransporte As String
        Public IDDocumento As String
        Public IDTipoDocumento As String
        Public NumeroDocumento As String
        Public FechaDocumento As Date

        Public CodigoTipoDestino As String
        Public CodigoTipoDestinoInterno As String
        Public TipoDAA As TipoDAA
        Public TipoEnvioDAA As enumTipoEnvioDAA

        'Public Filtro As Filter 'origen
        Public RegistrosEmpresas As New RegistroEmpresaInfo
        Public CampoAgrupacionOrigenLineasPorDefecto As String = "IDLineaAlbaran"
        Public CampoID As String = "IDAlbaran"
        Public VistaOrigenLineas As String = CN_VistaDAAAlbaran

        'Public OrigenExterno As DataTable  '//Para accesos desde otra BBDD

#End Region

#Region "           Ayudas"

        Public Origen As enumOrigenDAA
        Public DefaultDAA As DataTable
        Public DefaultDAALinea As DataTable
        Public DefaultDAAOperacion As DataTable

#End Region

#Region "           Resultados"

        Public Cabecera As DataTable
        Public Lineas As DataTable
        Public LineasOperacion As DataTable
        Public LineasPaquete As DataTable
        Public DetalleTransporte As DataTable
        Public DetalleDocumento As DataTable

        Public Direcciones As DireccionesDAA

#End Region

#Region "           Constantes"

        Public Const CN_DefaultIntracomunitario As String = "00000000000000000000000000000001"
        Public Const CN_DefaultInterno As String = "00000000000000000000000000000000"


#End Region

#Region "           Métodos"

        'REVISAR
        Public Sub New()
            InitializeData(enumTipoEnvioDAA.EMCSInterno)
        End Sub

        Public Sub New(ByVal TipoEnvioDAA As enumTipoEnvioDAA)
            InitializeData(TipoEnvioDAA)
        End Sub

        Protected Sub InitializeData(ByVal TipoEnvioDAA As enumTipoEnvioDAA)
            Me.TipoEnvioDAA = TipoEnvioDAA
            EstablecerDAADefault()
            Cabecera = New BdgDAACabecera().AddNew
            Lineas = New BdgDAALinea().AddNew
            LineasOperacion = New BdgDAALineaOperacion().AddNew
            LineasPaquete = New BdgDAALineaPaquete().AddNew
            DetalleDocumento = New BdgDAADetalleDocumento().AddNew
            DetalleTransporte = New BdgDAADetalleTransporte().AddNew
            Origen = enumOrigenDAA.Manual
        End Sub

        Protected Sub EstablecerDAADefault()
            If (Nz(Me.TipoEnvioDAA, enumTipoEnvioDAA.EMCSIntracomunitario) = enumTipoEnvioDAA.EMCSIntracomunitario) Then
                Me.DefaultDAA = New DataEngine().Filter("tbDAACabecera", New GuidFilterItem("IDDaa", New Guid(CN_DefaultIntracomunitario)))
            Else
                Me.DefaultDAA = New DataEngine().Filter("tbDAACabecera", New GuidFilterItem("IDDaa", New Guid(CN_DefaultInterno)))
            End If

            If Not Me.DefaultDAA Is Nothing AndAlso Me.DefaultDAA.Rows.Count > 0 Then
                DefaultDAALinea = New DataEngine().Filter("tbDAALinea", New GuidFilterItem("IDDaa", DefaultDAA.Rows(0)("IDDaa")))
                If Not Me.DefaultDAALinea Is Nothing AndAlso Me.DefaultDAALinea.Rows.Count > 0 Then
                    Me.DefaultDAAOperacion = New DataEngine().Filter("tbDAALineaOperacion", New GuidFilterItem("IDDaaLinea", DefaultDAALinea.Rows(0)("IDDaaLinea")))
                End If
            End If
        End Sub

#End Region

    End Class

    <Serializable()> _
    Public Class RegistroEmpresaInfo

        Public Listado As DataTable

        Public Sub New()
            Listado = New DataTable()
            Listado.Columns.Add("IDRegistro", GetType(Integer))
            Listado.Columns.Add("IDBaseDatos", GetType(Guid))
        End Sub

        Public Sub Add(ByVal intIDRegistro As Integer, ByVal idbbdd As Guid)
            Dim dtr As DataRow = Listado.NewRow
            dtr("IDRegistro") = intIDRegistro
            dtr("IDBaseDatos") = idbbdd
            Listado.Rows.Add(dtr)
        End Sub

    End Class


    <Serializable()> _
    Public Class stCrearDAAInfoResult
        Public NDAA As String
        Public IDDAA As Guid
        Public Source As String
        Public SourceEntity As String
        Public ErrorMessage As String
    End Class

    <Serializable()> _
Public Class stCrearDAAInfoCollection
        Public DAAInfoCollection() As stCrearDAAInfo
    End Class
    <Serializable()> _
    Public Class stCrearDAAInfoCollectionResult
        Public InfoResult() As stCrearDAAInfoResult
    End Class

    <Serializable()> Public Class stDatosDAADesdeAlbaran
        Public IDContador As String
        Public RegistrosEmpresas As New RegistroEmpresaInfo
        Public DireccionEntrega As Integer
        Public DireccionDestino As Integer
        Public TipoEnvioDAA As Integer
        Public Aduana As String
        Public IDDAA As Guid
    End Class

    <Task()> Public Shared Function GetTipoDocumentoAlbaran(ByVal data As Object, ByVal services As ServiceProvider) As String
        Return "ALB"
    End Function

    <Task()> Public Shared Function GetTipoDocumentoFactura(ByVal data As Object, ByVal services As ServiceProvider) As String
        Return "FAC"
    End Function

    <Task()> Public Shared Function GetTipoDocumentoOtros(ByVal data As Object, ByVal services As ServiceProvider) As String
        Return "OTR"
    End Function

    <Task()> Public Shared Function PrepararDatosDAADesdeAlbaran(ByVal data As stDatosDAADesdeAlbaran, ByVal services As ServiceProvider) As stCrearDAAInfo
        Dim Tipo_Documento_Albaran As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf GetTipoDocumentoAlbaran, Nothing, services)
        Dim Tipo_Documento_Factura As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf GetTipoDocumentoFactura, Nothing, services)
        Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase

        Dim strAlbaran As String = String.Empty
        Dim AlbaranBBDD As Guid = Guid.Empty

        'Tomamos los datos de cabecera del primer documento de la empresa actual y sino del primer documento del resto
        For Each dr As DataRow In data.RegistrosEmpresas.Listado.Select("IDBaseDatos = '" & currentBBDD.ToString & "'")
            strAlbaran = dr("IDRegistro")
            AlbaranBBDD = dr("IDBaseDatos")
            If Length(strAlbaran) > 0 Then Exit For
        Next

        If Length(strAlbaran) = 0 Then
            strAlbaran = data.RegistrosEmpresas.Listado(0)("IDRegistro")
            AlbaranBBDD = data.RegistrosEmpresas.Listado(0)("IDBaseDatos")
        End If


        '//Hay que crear los datos por defecto del DAA desde la BBDD en la que nos encontramos
        Dim dataResult As New stCrearDAAInfo(data.TipoEnvioDAA)

        If currentBBDD <> AlbaranBBDD Then
            AdminData.SetCurrentConnection(AlbaranBBDD)
            AdminData.CommitTx(True)
        End If

        Dim dttDatosAlbaran As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(strAlbaran)
        If (dttDatosAlbaran Is Nothing OrElse dttDatosAlbaran.Rows.Count = 0) Then
            ApplicationService.GenerateError("El albarán indicado no existe")
        End If

        dataResult.Contador = data.IDContador
        dataResult.IDDocumento = dttDatosAlbaran.Rows(0)("IDAlbaran")
        dataResult.FechaDocumento = dttDatosAlbaran.Rows(0)("FechaAlbaran")
        'dataResult.NumeroCertificado = dttDatosAlbaran.Rows(0)("NAlbaran")
        dataResult.IDModoTransporte = dttDatosAlbaran.Rows(0)("IDModoTransporte") & String.Empty
        dataResult.Transportista = dttDatosAlbaran.Rows(0)("EmpresaTransp") & String.Empty
        dataResult.NIFResponsableTransporte = dttDatosAlbaran.Rows(0)("DNIConductor") & String.Empty
        dataResult.ResponsableTransporte = dttDatosAlbaran.Rows(0)("Conductor") & String.Empty
        dataResult.FormaEnvio = dttDatosAlbaran.Rows(0)("IDFormaEnvio") & String.Empty
        dataResult.IDTipoDocumento = Tipo_Documento_Albaran
        dataResult.NumeroDocumento = dttDatosAlbaran.Rows(0)("NAlbaran") & String.Empty
        dataResult.FechaDocumento = Nz(dttDatosAlbaran.Rows(0)("FechaAlbaran"), cnMinDate)
        Dim dtFra As DataTable = New BE.DataEngine().Filter("vNegAlbaranFacturaVenta", New NumberFilterItem("IDAlbaran", strAlbaran))
        If dtFra.Rows.Count > 0 Then
            dataResult.IDTipoDocumento = Tipo_Documento_Factura
            dataResult.NumeroDocumento = dtFra.Rows(0)("NFactura") & String.Empty
            dataResult.FechaDocumento = Nz(dtFra.Rows(0)("FechaFactura"), cnMinDate)
        End If
        dataResult.Cliente = dttDatosAlbaran.Rows(0)("IDCliente") & String.Empty
        dataResult.Matricula = dttDatosAlbaran.Rows(0)("Matricula") & String.Empty
        dataResult.Remolque = dttDatosAlbaran.Rows(0)("Remolque") & String.Empty
        dataResult.Precinto = dttDatosAlbaran.Rows(0)("Precinto") & String.Empty
        dataResult.Contenedor = dttDatosAlbaran.Rows(0)("NContenedor") & String.Empty

        dataResult.RegistrosEmpresas = data.RegistrosEmpresas
        dataResult.VistaOrigenLineas = CN_VistaDAAAlbaran
        dataResult.CampoAgrupacionOrigenLineasPorDefecto = "IDLineaAlbaran"
        dataResult.CampoID = "IDAlbaran"
        dataResult.Origen = enumOrigenDAA.Albaran

        dataResult.Direccion = dttDatosAlbaran.Rows(0)("IDDireccion")
        dataResult.DireccionDestino = data.DireccionDestino
        dataResult.DireccionEntrega = data.DireccionEntrega
        dataResult.Aduana = data.Aduana

        dataResult.TipoDAA = ProcessServer.ExecuteTask(Of Long, TipoDAA)(AddressOf TipoDAAPorPais, dttDatosAlbaran.Rows(0)("IDDireccion"), services)

        If currentBBDD <> AlbaranBBDD Then
            AdminData.SetCurrentConnection(currentBBDD)
            AdminData.CommitTx(True)
        End If

        Return dataResult
    End Function

    <Serializable()> Public Class stDatosDAADesdePedido
        Public IDContador As String
        Public RegistrosEmpresas As New RegistroEmpresaInfo
        Public DireccionEntrega As Integer
        Public DireccionDestino As Integer
        Public TipoEnvioDAA As Integer
        Public Aduana As String
    End Class

    <Task()> Public Shared Function PrepararDatosDAADesdePedido(ByVal data As stDatosDAADesdePedido, ByVal services As ServiceProvider) As stCrearDAAInfo

        Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase

        Dim strPedido As String = String.Empty
        Dim PedidoBBDD As Guid = Guid.Empty

        'Tomamos los datos de cabecera del primer documento de la empresa actual y sino del primer documento del resto
        For Each dr As DataRow In data.RegistrosEmpresas.Listado.Select("IDBaseDatos = '" & currentBBDD.ToString & "'")
            strPedido = dr("IDRegistro")
            PedidoBBDD = dr("IDBaseDatos")
            If Length(strPedido) > 0 Then Exit For
        Next

        If Length(strPedido) = 0 Then
            strPedido = data.RegistrosEmpresas.Listado(0)("IDRegistro")
            PedidoBBDD = data.RegistrosEmpresas.Listado(0)("IDBaseDatos")
        End If

        If currentBBDD <> PedidoBBDD Then
            AdminData.SetCurrentConnection(PedidoBBDD)
            AdminData.CommitTx(True)
        End If

        Dim dttDatosPedido As DataTable = New PedidoVentaCabecera().SelOnPrimaryKey(strPedido)
        If (dttDatosPedido Is Nothing OrElse dttDatosPedido.Rows.Count = 0) Then
            ApplicationService.GenerateError("El Pedido indicado no existe")
        End If

        Dim dataReturn As New stCrearDAAInfo(data.TipoEnvioDAA)
        dataReturn.Contador = data.IDContador
        dataReturn.FechaDocumento = dttDatosPedido.Rows(0)("FechaPedido")
        'dataReturn.NumeroCertificado = dttDatosPedido.Rows(0)("NPedido")
        dataReturn.IDModoTransporte = dttDatosPedido.Rows(0)("IDModoTransporte") & String.Empty
        dataReturn.FormaEnvio = dttDatosPedido.Rows(0)("IDFormaEnvio") & String.Empty
        dataReturn.NumeroDocumento = dttDatosPedido.Rows(0)("NPedido") & String.Empty
        dataReturn.Cliente = dttDatosPedido.Rows(0)("IDCliente") & String.Empty
        dataReturn.RegistrosEmpresas = data.RegistrosEmpresas
        dataReturn.VistaOrigenLineas = CN_VistaDAAPedido
        dataReturn.CampoAgrupacionOrigenLineasPorDefecto = "IDLineaPedido"
        dataReturn.CampoID = "IDPedido"
        dataReturn.Origen = enumOrigenDAA.Pedido

        dataReturn.Direccion = dttDatosPedido.Rows(0)("IDDireccionEnvio")
        dataReturn.DireccionDestino = data.DireccionDestino
        dataReturn.DireccionEntrega = data.DireccionEntrega
        dataReturn.Aduana = data.Aduana

        dataReturn.TipoDAA = ProcessServer.ExecuteTask(Of Long, TipoDAA)(AddressOf TipoDAAPorPais, dttDatosPedido.Rows(0)("IDDireccionEnvio"), services)

        If currentBBDD <> PedidoBBDD Then
            AdminData.SetCurrentConnection(currentBBDD)
            AdminData.CommitTx(True)
        End If

        Return dataReturn
    End Function

#End Region

#Region "Métodos Individuales Secciones DAA"

    <Task()> Public Shared Function CrearDAACabeceraBase(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        'El motivo de los varios 'case' es mantener el orden de asignación según aparecen en el documento

        '01. Parte común
        data.Cabecera = New BdgDAACabecera().AddNewForm
        data.Cabecera.Rows(0)("IDDaa") = Guid.NewGuid
        data.Cabecera.Rows(0)("TestEMCS") = data.DefaultDAA.Rows(0)("TestEMCS")
        If Length(data.Contador) > 0 Then
            Dim StDatos As New Contador.DatosCounterValue(data.Contador, New BdgDAACabecera, "NDAA", "FechaDocumento", data.FechaDocumento)
            data.Cabecera.Rows(0)("NDAA") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
            data.Cabecera.Rows(0)("IDContador") = data.Contador
        End If

        If (data.Cabecera.Rows(0).IsNull("NDAA")) Then
            Dim StCont As New Contador.DatosDefaultCounterValue(data.Cabecera.Rows(0), "BdgDAACabecera", "NDAA")
            ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StCont, services)

            If (data.Cabecera.Rows(0).IsNull("NDAA")) Then
                ApplicationService.GenerateError("No se ha definido un contador predeterminado para los DAAs")
            End If
        End If

        Dim oDE As DatosEmpresa = New DatosEmpresa
        Dim dtDE As DataTable = oDE.Filter(New Filter())

        If dtDE.Rows.Count = 0 Then
            ApplicationService.GenerateError("No se puede generar un DAA sin Datos de Empresa.")
        ElseIf dtDE.Rows(0).IsNull("IdPais") Then
            ApplicationService.GenerateError("No se ha especificado un Pais en Datos de Empresa.")
        End If

        data.Cabecera.Rows(0)("EmisorMensaje") = dtDE.Rows(0)("CIF")                'SENDER
        data.Cabecera.Rows(0)("NombreEmisorMensaje") = dtDE.Rows(0)("DescEmpresa")  'RAZON SOCIAL
        data.Cabecera.Rows(0)("DestinatarioMensaje") = "NDEA.ES" '("CIF")           'RECIPIENT
        data.Cabecera.Rows(0)("FechaPreparacion") = Date.Today                      'DATEPREPARATION
        data.Cabecera.Rows(0)("HoraPreparacion") = Date.Now                         'TIMEPREPARATION
        'data.Cabecera.Rows(0)("IDMensaje") = Date.Now                               'MIDENTIFIER
        data.Cabecera.Rows(0)("CMensaje") = ""                                      'CIDENTIFIER

        data.Cabecera.Rows(0)("TipoInformacion") = data.DefaultDAA.Rows(0)("TipoInformacion")                          'MTYPE

        data.Cabecera.Rows(0)("IndicadorPresentacionDiferida") = CInt(Nz(data.DefaultDAA.Rows(0)("IndicadorPresentacionDiferida"), 0))      'DEFERRED SUBMISSION

        data.Cabecera.Rows(0)("LenguaGruposDatos") = data.DefaultDAA.Rows(0)("LenguaGruposDatos")

        If Length(data.Direccion) = 0 Then data.Direccion = data.DireccionDestino
        data.Cabecera(0)("TipoDAA") = data.TipoDAA 'ProcessServer.ExecuteTask(Of stCrearDAAInfo, String)(AddressOf ObtenerTipoDAA, data, services)
        data.Cabecera(0)("TipoEnvioDAA") = data.TipoEnvioDAA

        If data.Direcciones Is Nothing Then
            If data.DireccionDestino <> 0 Then
                Dim oCliDi As New ClienteDireccion
                Dim rowDir As DataRow = oCliDi.GetItemRow(data.DireccionDestino)
                'Tipo Destino Intracomunitario
                data.CodigoTipoDestino = rowDir("CodigoTipoDestino") & String.Empty
                If Length(data.CodigoTipoDestino) = 0 Then data.CodigoTipoDestino = data.DefaultDAA.Rows(0)("CodigoTipoDestino") & String.Empty

                'Tipo Destino Interno
                data.CodigoTipoDestinoInterno = rowDir("TipoDestinoInterno") & String.Empty
                If Length(data.CodigoTipoDestinoInterno) = 0 Then data.CodigoTipoDestinoInterno = data.DefaultDAA.Rows(0)("TipoDestinoInterno") & String.Empty

                If Length(data.CodigoTipoDestino) > 0 Then data.Cabecera(0)("CodigoTipoDestino") = data.CodigoTipoDestino
                data.Cabecera(0)("TipoDestinoInterno") = data.CodigoTipoDestinoInterno
            End If

            Dim StData As New StDireccionesPorCodigoTipoDestino(data.CodigoTipoDestino, data.DireccionDestino, data.DireccionEntrega, data.CodigoTipoDestinoInterno, data.TipoEnvioDAA)
            data.Direcciones = ProcessServer.ExecuteTask(Of StDireccionesPorCodigoTipoDestino, DireccionesDAA)(AddressOf DireccionesPorCodigoTipoDestino, StData, services)
        End If

        data.Cabecera(0)("IDDireccionDestino") = data.DireccionDestino
        data.Cabecera(0)("IDDireccionEntrega") = data.DireccionEntrega

        Return data
    End Function

    '5 - DESTINATARIO CEE
    '4 - DESTINATARIO INTERNO
    <Task()> Public Shared Function CrearDAACabeceraDestinatario(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        If (data.Direcciones Is Nothing) Then
            ApplicationService.GenerateError("No existen direcciones relacionadas con el destinatario.")
        End If
        data.Cabecera.Rows(0)("CAEDestinatario") = data.Direcciones.CAEDestinatario             '5a TRADERID
        data.Cabecera.Rows(0)("Destinatario") = data.Direcciones.Destinatario                   '5b TRADERNAME
        data.Cabecera.Rows(0)("DireccionDestinatario") = data.Direcciones.DireccionDestinatario '5c STREET-NAME
        data.Cabecera.Rows(0)("CodPostalDestinatario") = data.Direcciones.CodPostalDestinatario '5e POSTCODE
        data.Cabecera.Rows(0)("PoblacionDestinatario") = data.Direcciones.PoblacionDestinatario '5f CITY
        data.Cabecera.Rows(0)("IDPaisDestinatario") = data.Direcciones.IDPaisDestinatario       '
        data.Cabecera.Rows(0)("ISOPaisDestinatario") = data.Direcciones.ISOPaisDestinatario     '

        If data.TipoEnvioDAA = enumTipoEnvioDAA.EMCSInterno Then
            data.Cabecera.Rows(0)("NIFDestinatario") = data.Direcciones.NIFDestinatario       '4b
        End If

        Return data
    End Function

    '2 - EXPEDIDOR CEE
    '1 - EXPEDIDOR INTERNO
    <Task()> Public Shared Function CrearDAACabeceraExpedidor(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        Dim dttDatosEmpresa As DataTable = New DatosEmpresa().Filter()
        If (dttDatosEmpresa Is Nothing OrElse dttDatosEmpresa.Rows.Count = 0) Then
            ApplicationService.GenerateError("No existen los datos de la empresa Expedidora. Compruebe el mantenimiento de Datos Empresa.")
        End If
        
        '2a TRADEREXCISEN
        Dim datCAE As New DataGetCAEExpedidor(data.DefaultDAA, dttDatosEmpresa)
        Dim CAEExpedidor As String = ProcessServer.ExecuteTask(Of DataGetCAEExpedidor, String)(AddressOf GetCAEExpedidor, datCAE, services)
        If Length(CAEExpedidor) > 0 Then
            data.Cabecera.Rows(0)("CAEExpedidor") = CAEExpedidor
        End If

        If (Not data.DefaultDAA Is Nothing AndAlso data.DefaultDAA.Rows.Count > 0 AndAlso Length(data.DefaultDAA.Rows(0)("Expedidor")) > 0) Then
            data.Cabecera.Rows(0)("Expedidor") = data.DefaultDAA.Rows(0)("Expedidor")             '2b TRADERNAME
            data.Cabecera.Rows(0)("DireccionExpedidor") = data.DefaultDAA.Rows(0)("DireccionExpedidor")      '2c STREET-NAME
            data.Cabecera.Rows(0)("CodPostalExpedidor") = data.DefaultDAA.Rows(0)("CodPostalExpedidor")      '2e POSTCODE
            data.Cabecera.Rows(0)("PoblacionExpedidor") = data.DefaultDAA.Rows(0)("PoblacionExpedidor")      '2f CITY
        Else
            data.Cabecera.Rows(0)("Expedidor") = dttDatosEmpresa.Rows(0)("DescEmpresa")             '2b TRADERNAME
            data.Cabecera.Rows(0)("DireccionExpedidor") = dttDatosEmpresa.Rows(0)("Direccion")      '2c STREET-NAME
            data.Cabecera.Rows(0)("CodPostalExpedidor") = dttDatosEmpresa.Rows(0)("CodPostal")      '2e POSTCODE
            data.Cabecera.Rows(0)("PoblacionExpedidor") = dttDatosEmpresa.Rows(0)("Poblacion")      '2f CITY
        End If

        'EMCS INTERNO
        If (Not data.DefaultDAA Is Nothing AndAlso data.DefaultDAA.Rows.Count > 0 AndAlso Length(data.DefaultDAA.Rows(0)("CIFExpedidor")) > 0) Then
            data.Cabecera.Rows(0)("CIFExpedidor") = data.DefaultDAA.Rows(0)("CIFExpedidor") '2a TRADEREXCISEN
        Else
            data.Cabecera.Rows(0)("CIFExpedidor") = dttDatosEmpresa.Rows(0)("Cif")  '2a TRADEREXCISEN
        End If

        Return data
    End Function

    <Serializable()> _
    Public Class DataGetCAEExpedidor
        Public DefaultDAA As DataTable
        Public dttDatosEmpresa As DataTable

        Public Sub New(Optional ByVal DefaultDAA As DataTable = Nothing, Optional ByVal dttDatosEmpresa As DataTable = Nothing)
            Me.DefaultDAA = DefaultDAA
            Me.dttDatosEmpresa = dttDatosEmpresa
        End Sub
    End Class
    <Task()> Public Shared Function GetCAEExpedidor(ByVal data As DataGetCAEExpedidor, ByVal services As ServiceProvider) As String
        Dim CAEExpedidor As String

        Dim IDCentroGestion As String
        Dim cgu As UsuarioCentroGestion.UsuarioCentroGestionInfo = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
        If Not cgu Is Nothing AndAlso Length(cgu.IDCentroGestion) > 0 Then
            IDCentroGestion = cgu.IDCentroGestion
        End If

        If Length(IDCentroGestion) > 0 Then
            Dim CentrosGestion As EntityInfoCache(Of CentroGestionInfo) = services.GetService(Of EntityInfoCache(Of CentroGestionInfo))()
            Dim CentroG As CentroGestionInfo = CentrosGestion.GetEntity(IDCentroGestion)
            CAEExpedidor = CentroG.IDCAE
        End If
        If Length(CAEExpedidor) = 0 Then
            If (Not data.DefaultDAA Is Nothing AndAlso data.DefaultDAA.Rows.Count > 0 AndAlso Length(data.DefaultDAA.Rows(0)("CAEExpedidor")) > 0) Then
                CAEExpedidor = data.DefaultDAA.Rows(0)("CAEExpedidor") '2a TRADEREXCISEN
            Else
                If data.dttDatosEmpresa Is Nothing Then data.dttDatosEmpresa = New DatosEmpresa().Filter()
                CAEExpedidor = data.dttDatosEmpresa.Rows(0)("IDCAE")  '2a TRADEREXCISEN
            End If
        End If
        Return CAEExpedidor
    End Function

    '3 - LUGARDESPACHO 'DUDA esto no sería el lugar de datos empresa????
    <Task()> Public Shared Function CrearDAACabeceraLugarExpedicion(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo

        Dim oDE As DatosEmpresa = New DatosEmpresa
        Dim dttDatosEmpresa As DataTable = oDE.Filter(New Filter())


        If (Not data.DefaultDAA Is Nothing AndAlso data.DefaultDAA.Rows.Count > 0 AndAlso Length(data.DefaultDAA.Rows(0)("CAELugarDespacho")) > 0) Then
            data.Cabecera.Rows(0)("CAELugarDespacho") = data.DefaultDAA.Rows(0)("CAELugarDespacho") '3a NTAXWAREHOUSE
        Else
            data.Cabecera.Rows(0)("CAELugarDespacho") = dttDatosEmpresa.Rows(0)("IDCAE") '3a NTAXWAREHOUSE
        End If
        If (Not data.DefaultDAA Is Nothing AndAlso data.DefaultDAA.Rows.Count > 0 AndAlso Length(data.DefaultDAA.Rows(0)("LugarDespacho")) > 0) Then
            data.Cabecera.Rows(0)("LugarDespacho") = data.DefaultDAA.Rows(0)("LugarDespacho")             '3b TRADERNAME
            data.Cabecera.Rows(0)("DireccionLugarDespacho") = data.DefaultDAA.Rows(0)("DireccionLugarDespacho")      '3c STREET-NAME
            data.Cabecera.Rows(0)("CodPostalLugarDespacho") = data.DefaultDAA.Rows(0)("CodPostalLugarDespacho")      '3e POSTCODE
            data.Cabecera.Rows(0)("PoblacionLugarDespacho") = data.DefaultDAA.Rows(0)("PoblacionLugarDespacho")      '3f CITY
        Else
            data.Cabecera.Rows(0)("LugarDespacho") = dttDatosEmpresa.Rows(0)("DescEmpresa")             '3b TRADERNAME
            data.Cabecera.Rows(0)("DireccionLugarDespacho") = dttDatosEmpresa.Rows(0)("Direccion")      '3c STREET-NAME
            data.Cabecera.Rows(0)("CodPostalLugarDespacho") = dttDatosEmpresa.Rows(0)("CodPostal")      '3e POSTCODE
            data.Cabecera.Rows(0)("PoblacionLugarDespacho") = dttDatosEmpresa.Rows(0)("Poblacion")      '3f CITY
        End If

        Return data
    End Function

    '4.- Import Office OK
    <Task()> Public Shared Function CrearDAACabeceraOficinaImportacion(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        'Código 	Oficina 	    Código 	Oficina 	Código 	Oficina 	Código 	Oficina 
        'D01200 	Álava 	        D15200 	A Coruña 	D29200 	Málaga 	    D43200 	Tarragona 
        'D02200 	Albacete 	    D16200 	Cuenca 	    D30200 	Murcia 	    D44200 	Teruel 
        'D03200 	Alicante 	    D17200 	Girona 	    D31200 	Navarra 	D45200 	Toledo 
        'D04200 	Almería 	    D18200 	Granada 	D32200 	Ourense 	D46200 	Valencia 
        'D05200 	Ávila 	        D19200 	Guadalajara D33200 	Oviedo 	    D47200 	Valladolid 
        'D06200 	Badajoz 	    D20200 	Guipúzcoa 	D34200 	Palencia 	D48200 	Vizcaya 
        'D07200 	Illes Baleares 	D21200 	Huelva 	    D49200 	Zamora 	    D08200 	Barcelona 
        'D22200 	Huesca 	        D36200 	Pontevedra 	D50200 	Zaragoza 	D09200 	Burgos 
        'D23200 	Jaén 	        D37200 	Salamanca 	D51200 	Cartagena 	D10200 	Cáceres 
        'D24200 	León 	        D52200 	Gijón 	    D11200 	Cádiz 	    D25200 	Lleida 
        'D39200 	Santander 	    D53200 	Jerez       D12200 	Castellón 	D26200 	La Rioja 
        'D40200 	Segovia 	    D54200 	Vigo 	    D13200 	Ciudad Real	D27200 	Lugo 
        'D41200 	Sevilla 	    D14200 	Córdoba 	D28200 	Madrid 	    D42200 	Soria 
        data.Cabecera.Rows(0)("CodigoAduanaImportacion") = data.DefaultDAA.Rows(0)("CodigoAduanaImportacion")   '4a REFERENCENº
        Return data
    End Function

    '6.- Complement Consignee OK
    <Task()> Public Shared Function CrearDAACabeceraConsignatario(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        data.Cabecera.Rows(0)("EstadoMiembro") = data.EstadoMiembro 'Lista 11 -     '6a MEMBERSTATE
        data.Cabecera.Rows(0)("NumeroCertificado") = data.NumeroCertificado         '6b NCERTIFICATE
        Return data
    End Function

    '7.- DeliveryPlace OK
    <Task()> Public Shared Function CrearDAACabeceraLugarEntrega(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        If (data.Direcciones Is Nothing) Then
            ApplicationService.GenerateError("No existen direcciones relacionadas con el destinatario.")
        End If

        data.Cabecera.Rows(0)("IDCAELugarEntrega") = data.Direcciones.IDCAELugarEntrega 'C074 / R045 -      '7a TRADERID
        data.Cabecera.Rows(0)("RazonSocialLugarEntrega") = data.Direcciones.RazonSocialLugarEntrega 'C079 - '7b TRADERNAME
        data.Cabecera.Rows(0)("DireccionLugarEntrega") = data.Direcciones.DireccionLugarEntrega 'C078 -     '7c STREETNAME
        data.Cabecera.Rows(0)("CodPostalLugarEntrega") = data.Direcciones.CodPostalLugarEntrega 'C078 -     '7e POSTCODE
        data.Cabecera.Rows(0)("PoblacionLugarEntrega") = data.Direcciones.PoblacionLugarEntrega 'C078 -     '7f CITY
        Return data

    End Function

    '8.- D.Place-Customs Office C013    CEE
    '5.- Aduana Exportación             INTERNO
    <Task()> Public Shared Function CrearDAACabeceraOficinaExportacion(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo

        Select Case data.TipoEnvioDAA
            Case enumTipoEnvioDAA.EMCSIntracomunitario
                'data.Cabecera.Rows(0)("CodigoAduanaExportacion") = data.CodigoAduanaExportacion '8a

                '8a CodigoAduanaExportacion
                data.Cabecera.Rows(0)("CodigoAduanaExportacion") = data.Direcciones.CodigoAduanaExportacionLugarEntrega
                If Length(data.Cabecera.Rows(0)("CodigoAduanaExportacion")) = 0 AndAlso Nz(data.Cabecera.Rows(0)("CodigoTipoDestino"), -1) = BdgCodigoTipoDestino.Exportacion Then
                    data.Cabecera.Rows(0)("CodigoAduanaExportacion") = data.DefaultDAA.Rows(0)("CodigoAduanaExportacion")
                End If
            Case enumTipoEnvioDAA.EMCSInterno
                '5a CodigoAduanaExportacion
                data.Cabecera.Rows(0)("CodigoAduanaExportacion") = data.Direcciones.CodigoAduanaExportacionDestinatario
                If Length(data.Cabecera.Rows(0)("CodigoAduanaExportacion")) = 0 Then
                    data.Cabecera.Rows(0)("CodigoAduanaExportacion") = data.DefaultDAA.Rows(0)("CodigoAduanaExportacion")
                End If
                '5b ISOPaisDestinatario
                data.Cabecera.Rows(0)("IDPaisDestinatario") = data.Direcciones.IDPaisDestinatario
                data.Cabecera.Rows(0)("ISOPaisDestinatario") = data.Direcciones.ISOPaisDestinatario
        End Select

        Return data
    End Function

    '10.- Dispatch Office OK
    <Task()> Public Shared Function CrearDAACabeceraOficinaExpedicion(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        '10A CEE
        '2i  INTERNO
        data.Cabecera.Rows(0)("CodigoAduanaDespacho") = data.DefaultDAA.Rows(0)("CodigoAduanaDespacho")

        '2j  INTERNO
        data.Cabecera.Rows(0)("Garantia") = data.DefaultDAA.Rows(0)("Garantia")

        Return data
    End Function

    '14.- T.Arranger Trader OK
    <Task()> Public Shared Function CrearDAACabeceraResponsableTransporte(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        data.Cabecera.Rows(0)("NIFResponsableTransporte") = data.NIFResponsableTransporte '14a
        data.Cabecera.Rows(0)("ResponsableTransporte") = data.ResponsableTransporte '14b
        data.Cabecera.Rows(0)("DireccionResponableTransporte") = "" '14c
        data.Cabecera.Rows(0)("CodigoPostalResponableTransporte") = "" '14e
        data.Cabecera.Rows(0)("PoblacionResponableTransporte") = "" '14f
        Return data
    End Function

    '15.- First Transporter Trader
    <Task()> Public Shared Function CrearDAACabeceraTransportista(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        If Length(data.FormaEnvio) = 0 Then
            data.Cabecera.Rows(0)("NIFTransportista") = data.NIFTransportista '15a
            data.Cabecera.Rows(0)("Transportista") = data.Transportista '15b
        Else
            Dim oFormaEnvio As New FormaEnvio
            Dim dttFormaEnvio As DataTable

            dttFormaEnvio = oFormaEnvio.SelOnPrimaryKey(data.FormaEnvio)
            If Not dttFormaEnvio Is Nothing And dttFormaEnvio.Rows.Count > 0 And Not IsDBNull(dttFormaEnvio.Rows(0)("IDProveedor")) Then
                Dim oProveedor As New Proveedor
                Dim dttProveedor As DataTable = oProveedor.SelOnPrimaryKey(dttFormaEnvio.Rows(0)("IDProveedor"))
                If Not dttProveedor Is Nothing And dttProveedor.Rows.Count > 0 Then
                    data.Cabecera.Rows(0)("Transportista") = dttProveedor.Rows(0)("DescProveedor") & " - " & dttProveedor.Rows(0)("CifProveedor") & String.Empty
                    data.Cabecera.Rows(0)("NifTransportista") = dttProveedor.Rows(0)("CifProveedor")
                    data.Cabecera.Rows(0)("IDPaisTransportista") = dttProveedor.Rows(0)("IDPais")
                    Dim clsPais As New Pais
                    Dim dtPais As DataTable = clsPais.SelOnPrimaryKey(dttProveedor.Rows(0)("IDPais") & String.Empty)
                    If dtPais.Rows.Count = 1 Then
                        data.Cabecera.Rows(0)("ISOPaisTransportista") = dtPais.Rows(0)("CodigoISO") & String.Empty
                    End If
                    data.Cabecera.Rows(0)("DireccionTransportista") = dttProveedor.Rows(0)("Direccion")
                    data.Cabecera.Rows(0)("PoblacionTransportista") = dttProveedor.Rows(0)("Poblacion")
                    data.Cabecera.Rows(0)("CodigoPostalTransportista") = dttProveedor.Rows(0)("CodPostal")
                End If
            End If

        End If
        Return data
    End Function

    '16.- Transport Details : tbDaaDetalleTransporte
    <Task()> Public Shared Function CrearDAACabeceraDetalleTransporte(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        Return ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf BdgDAADetalleTransporte.CrearDAADetalleTransporte, data, services)

    End Function

    '18.- Document Certificate : tbDaaDetalleDocumento
    <Task()> Public Shared Function CrearDAACabeceraDetalleDocumento(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        Return ProcessServer.ExecuteTask(Of stCrearDAAInfo, stCrearDAAInfo)(AddressOf BdgDAADetalleDocumento.CrearDAADetalleDocumento, data, services)
    End Function

    '13.- Transport Mode 
    '1.-Header E-AAD (JUNTO CON 13 EN MO) OK
    <Task()> Public Shared Function CrearDAACabeceraExtraTransporte(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        Dim f As New Filter
        f.Add("IDDireccion", data.DireccionDestino)
        Dim dttClienteDireccion As DataTable = New ClienteDireccion().Filter(f)
        If (dttClienteDireccion Is Nothing OrElse dttClienteDireccion.Rows.Count = 0) Then
            ApplicationService.GenerateError("El registro no tiene una dirección de entrega válida.")
        End If

        data.Cabecera.Rows(0)("DuracionTransporte") = dttClienteDireccion.Rows(0)("DuracionTransporte") 'R054    1b
        data.Cabecera.Rows(0)("OrganizacionTransporte") = data.DefaultDAA.Rows(0)("OrganizacionTransporte") 'LISTA 70    1c
        data.Cabecera.Rows(0)("IDModoTransporte") = data.IDModoTransporte 'LISTA 67    13a
        data.Cabecera.Rows(0)("CodigoMedioTransporte") = data.DefaultDAA.Rows(0)("CodigoMedioTransporte") 'data.IDModoTransporte 'LISTA 67    13a

        Dim dttMedioTransp As DataTable = New ModoTrasporte().SelOnPrimaryKey(data.IDModoTransporte)
        If Not dttMedioTransp Is Nothing AndAlso dttMedioTransp.Rows.Count > 0 Then
            data.Cabecera.Rows(0)("InfoExtraMedioTransporte") = dttMedioTransp.Rows(0)("DescModoTransporte")    '13b
        End If
        Return data
    End Function

    '11.- Movement Guarantee 
    '12.- Guaranter Trader
    <Task()> Public Shared Function CrearDAACabeceraGarantia(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        data.Cabecera.Rows(0)("TipoGarante") = data.DefaultDAA.Rows(0)("TipoGarante") 'LISTA 29 11a
        data.Cabecera.Rows(0)("CAEGarante") = data.DefaultDAA.Rows(0)("CAEGarante") 'LISTA 29 11a
        data.Cabecera.Rows(0)("NumeroGarante") = data.DefaultDAA.Rows(0)("NumeroGarante") 'LISTA 29 11a
        data.Cabecera.Rows(0)("Garante") = data.DefaultDAA.Rows(0)("Garante") 'LISTA 29 11a
        data.Cabecera.Rows(0)("DireccionGarante") = data.DefaultDAA.Rows(0)("DireccionGarante") 'LISTA 29 11a
        data.Cabecera.Rows(0)("CodigoPostalDireccionGarante") = data.DefaultDAA.Rows(0)("CodigoPostalDireccionGarante") 'LISTA 29 11a
        data.Cabecera.Rows(0)("PoblacionGarante") = data.DefaultDAA.Rows(0)("PoblacionGarante") 'LISTA 29 11a

        Return data
    End Function

    '--17.- Body E-ADD LÍNEAS
    '-- 17.1 - Package
    '-- 17.2 - Wine Producto
    '-- 17.2.1 Tipo de operacion 
    '-- 17.3. Biocarburantes
    <Task()> Public Shared Function CrearDAALineas(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        ProcessServer.ExecuteTask(Of stCrearDAAInfo)(AddressOf BdgDAALinea.CrearDAALineas, data, services)
        Return data
    End Function

    '-- 9. E-AAD DRAFT
    '-- 2. DATOS EXPEDICION INTERNO
    <Task()> Public Shared Function CrearDAACabeceraExtra(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        If Length(data.IDTipoDocumento) > 0 Then data.Cabecera.Rows(0)("IDTipoDocumento") = data.IDTipoDocumento '9b
        data.Cabecera.Rows(0)("NumeroDocumento") = data.NumeroDocumento '9b
        data.Cabecera.Rows(0)("TipoExpedidorDocumento") = data.DefaultDAA.Rows(0)("TipoExpedidorDocumento") '9d
        data.Cabecera.Rows(0)("FechaDocumento") = data.FechaDocumento
        data.Cabecera.Rows(0)("FechaDespacho") = data.FechaDocumento '9e
        data.Cabecera.Rows(0)("HoraDespacho") = Now.AddHours(1) '9f
        data.Cabecera.Rows(0)("AutCesionInfAgri") = data.DefaultDAA.Rows(0)("AutCesionInfAgri") '9g       

        If data.TipoEnvioDAA = enumTipoEnvioDAA.EMCSInterno Then
            data.Cabecera.Rows(0)("RegimenFiscal") = data.DefaultDAA.Rows(0)("RegimenFiscal") '2c       
        End If

        Return data
    End Function

    '--6 INTERNO - ORGANISMO EXENTO
    <Task()> Public Shared Function CrearDAACabeceraOrganismoExentoInterno(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo
        data.Cabecera.Rows(0)("TipoOrganismoExento") = data.DefaultDAA.Rows(0)("TipoOrganismoExento") '6a
        data.Cabecera.Rows(0)("IDPaisOrganismoExento") = data.DefaultDAA.Rows(0)("IDPaisOrganismoExento")
        data.Cabecera.Rows(0)("ISOPaisOrganismoExento") = data.DefaultDAA.Rows(0)("ISOPaisOrganismoExento") '6b
        Return data
    End Function

    '--11 INTERNO - AUTORIDAD AGROALIMENTARIA
    <Task()> Public Shared Function CrearDAACabeceraAutoridadAgroAlimentariaInterno(ByVal data As stCrearDAAInfo, ByVal services As ServiceProvider) As stCrearDAAInfo

        If data.TipoEnvioDAA = enumTipoEnvioDAA.EMCSInterno Then
            'CodigoAutoridadExpedicion '11a
            data.Cabecera.Rows(0)("CodigoAutoridadExpedicion") = data.DefaultDAA.Rows(0)("CodigoAutoridadExpedicion") '11a
            'CodigoAutoridadDestino '11b
            If Length(data.Direcciones.AutoridadAgroalimentariaDestinatario) > 0 Then
                data.Cabecera.Rows(0)("CodigoAutoridadDestino") = data.Direcciones.AutoridadAgroalimentariaDestinatario '11b
            Else
                data.Cabecera.Rows(0)("CodigoAutoridadDestino") = data.DefaultDAA.Rows(0)("CodigoAutoridadDestino") '11b
            End If
            'CesionInformacion '11c
            data.Cabecera.Rows(0)("CesionInformacion") = data.DefaultDAA.Rows(0)("CesionInformacion") '11c
        End If

        Return data
    End Function
#End Region





    <Serializable()> _
    Public Class DataDesvincularDAAOrigenes
        Public Pedidos As New Dictionary(Of String, List(Of String))
        Public Albaranes As New Dictionary(Of String, List(Of String))
    End Class
    <Task()> Public Shared Function DesvincularDAAOrigenes(ByVal IDDaa As Guid, ByVal services As ServiceProvider) As DataDesvincularDAAOrigenes
        Dim dat As New DataDesvincularDAAOrigenes

        Dim currentBBDD As Guid = AdminData.GetConnectionInfo.IDDataBase
        Dim AVC As New AlbaranVentaCabecera
        Dim PVC As New PedidoVentaCabecera
        Dim dtOrigenes As DataTable = New BdgDAAOrigenes().Filter(New GuidFilterItem("IDDaa", IDDaa))
        If dtOrigenes.Rows.Count > 0 Then
            Dim f As New Filter
            f.Add(New GuidFilterItem("IDDAA", IDDaa))
            f.Add(New GuidFilterItem("IDDAABaseDatos", currentBBDD))

            Dim OrigenesOrdenados As List(Of DataRow) = (From c In dtOrigenes Order By c("IDBaseDatos")).ToList()
            For Each dr As DataRow In OrigenesOrdenados
                If currentBBDD <> dr("IDBaseDatos") Then
                    AdminData.SetCurrentConnection(dr("IDBaseDatos"))
                End If

                Dim DataBaseName As String = AdminData.GetSessionInfo.DataBase.DataBaseDescription

                Dim dtPedidos As DataTable = PVC.Filter(f)
                If dtPedidos.Rows.Count > 0 Then
                    For Each drPedido As DataRow In dtPedidos.Rows
                        drPedido("IDDAA") = System.DBNull.Value
                        drPedido("NDAA") = System.DBNull.Value
                        drPedido("IDDAABaseDatos") = System.DBNull.Value
                        drPedido("AadReferenceCode") = System.DBNull.Value

                        If dat.Pedidos.ContainsKey(DataBaseName) Then
                            Dim lstPedidos As List(Of String) = dat.Pedidos(DataBaseName)
                            lstPedidos.Add(drPedido("NPedido"))
                        Else
                            Dim lstPedidos As New List(Of String)
                            lstPedidos.Add(drPedido("NPedido"))
                            dat.Pedidos.Add(DataBaseName, lstPedidos)
                        End If
                    Next
                    BusinessHelper.UpdateTable(dtPedidos)
                End If

                Dim dtAlbaranes As DataTable = AVC.Filter(f)
                If dtAlbaranes.Rows.Count > 0 Then
                    For Each drAlbaran As DataRow In dtAlbaranes.Rows
                        drAlbaran("IDDAA") = System.DBNull.Value
                        drAlbaran("NDAA") = System.DBNull.Value
                        drAlbaran("IDDAABaseDatos") = System.DBNull.Value
                        drAlbaran("AadReferenceCode") = System.DBNull.Value

                        If dat.Albaranes.ContainsKey(DataBaseName) Then
                            Dim lstAlbaranes As List(Of String) = dat.Albaranes(DataBaseName)
                            lstAlbaranes.Add(drAlbaran("NAlbaran"))
                        Else
                            Dim lstAlbaranes As New List(Of String)
                            lstAlbaranes.Add(drAlbaran("NAlbaran"))
                            dat.Albaranes.Add(DataBaseName, lstAlbaranes)
                        End If
                    Next
                    BusinessHelper.UpdateTable(dtAlbaranes)
                End If
                If currentBBDD <> dr("IDBaseDatos") Then
                    AdminData.SetCurrentConnection(currentBBDD)
                End If
               
            Next
        End If

        Return dat
    End Function

#End Region

End Class

Public Enum TipoDAA
    Nacional
    CEE
    Exportacion
End Enum

Public Enum BdgCodigoTipoDestino
    DepositoFiscal = 1
    DestinatarioRegistrado = 2
    DestinatarioRegistradoOcasional = 3
    EntregaDirectaAutorizada = 4
    DestinatarioExento = 5
    Exportacion = 6
    DestinoDesconocido = 8
End Enum

<Serializable()> _
Public Class DireccionesDAA

    'Destinatario (Consignee)
    Public CAEDestinatario As String = String.Empty
    Public Destinatario As String = String.Empty
    Public DireccionDestinatario As String = String.Empty
    Public CodPostalDestinatario As String = String.Empty
    Public PoblacionDestinatario As String = String.Empty
    Public IDPaisDestinatario As String = String.Empty
    Public ISOPaisDestinatario As String = String.Empty
    Public AutoridadAgroalimentariaDestinatario As String = String.Empty
    Public CodigoAduanaExportacionDestinatario As String = String.Empty
    Public NIFDestinatario As String = String.Empty

    'Lugar de Entrega (Delivery Place)
    Public IDCAELugarEntrega As String = String.Empty
    Public RazonSocialLugarEntrega As String = String.Empty
    Public DireccionLugarEntrega As String = String.Empty
    Public CodPostalLugarEntrega As String = String.Empty
    Public PoblacionLugarEntrega As String = String.Empty
    Public IDPaisLugarEntrega As String = String.Empty
    Public ISOPaisLugarEntrega As String = String.Empty
    Public AutoridadAgroalimentariaLugarEntrega As String = String.Empty
    Public CodigoAduanaExportacionLugarEntrega As String = String.Empty

End Class


<Serializable()> _
Public Class OrigenesDAA
    Public BBDD As New List(Of Object)
End Class