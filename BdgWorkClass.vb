Public Class BdgWorkClass

    <Serializable()> _
    Public Class VinoData
        Public IDVino As Guid
        Public IDDeposito As String
        Public TipoDeposito As TipoDeposito
        Public IDArticulo As String
        Public Lote As String
        Public IDAlmacen As String
        Public Fecha As Date
        Public IDEstadoVino As String
        Public Origen As BdgOrigenVino
        Public NOperacion As String
        Public IDUdMedida As String
        'Public CosteTotal As String
        'Public CosteFiscal As String
        'Public CosteVariable As String
        Public IDBarrica As String
        Public Estructura As VinoComponente()

        Public Sub New(ByVal Deposito As String, _
                        ByVal TipDep As TipoDeposito, _
                        ByVal Articulo As String, _
                        ByVal Fecha As Date, _
                        ByVal Origen As BdgOrigenVino, _
                        ByVal UDMedida As String, _
                        ByVal Lote As String, _
                        ByVal Almacen As String)


            Me.New(Guid.NewGuid, Deposito, TipDep, Articulo, Fecha, Origen, UDMedida, Lote, Almacen)
        End Sub

        Private Sub New(ByVal NewGuid As Guid, _
                        ByVal Deposito As String, _
                        ByVal TipDep As TipoDeposito, _
                        ByVal Articulo As String, _
                        ByVal Fecha As Date, _
                        ByVal Origen As BdgOrigenVino, _
                        ByVal UDMedida As String, _
                        ByVal Lote As String, _
                        ByVal Almacen As String)

            Me.IDVino = NewGuid

            Me.IDDeposito = Deposito
            Me.TipoDeposito = TipDep
            Me.IDArticulo = Articulo
            Me.Fecha = Fecha
            Me.Origen = Origen
            Me.IDUdMedida = UDMedida
            Me.Lote = Lote
            Me.IDAlmacen = Almacen

        End Sub

        Public Function Clone() As VinoData
            Return New VinoData(IDDeposito, TipoDeposito, IDArticulo, Fecha, BdgOrigenVino.Interno, IDUdMedida, Lote, IDAlmacen)
        End Function

        'Public Function Clone(ByVal IDVino As Guid) As VinoData
        '    Return New VinoData(IDVino, IDDeposito, TipoDeposito, IDArticulo, Fecha, BdgOrigenVino.Interno, IDUdMedida, Lote, IDAlmacen)
        'End Function

    End Class

    <Serializable()> _
    Public Class StCrearVino
        Public IDDeposito As String
        Public IDArticulo As String
        Public Lote As String
        Public Fecha As Date
        Public Origen As BdgOrigenVino
        Public UDMedida As String
        Public Cantidad As Double
        Public IDEstadoVino As String
        Public NOperacion As String
        Public Estrct As VinoComponente()
        Public SubDeps As VinoSubDep()
        Public IDBarrica As String
        Public IDAlmacen As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal IDArticulo As String, ByVal Lote As String, ByVal Fecha As Date, ByVal Origen As BdgOrigenVino, _
                       ByVal UDMedida As String, ByVal Cantidad As Double, Optional ByVal IDEstadoVino As String = Nothing, Optional ByVal NOperacion As String = Nothing, _
                       Optional ByVal Estrct As VinoComponente() = Nothing, Optional ByVal SubDeps() As VinoSubDep = Nothing, Optional ByVal IDBarrica As String = Nothing)
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.Fecha = Fecha
            Me.Origen = Origen
            Me.UDMedida = UDMedida
            Me.Cantidad = Cantidad
            Me.IDEstadoVino = IDEstadoVino
            Me.NOperacion = NOperacion
            Me.Estrct = Estrct
            Me.SubDeps = SubDeps
            Me.IDBarrica = IDBarrica
        End Sub
    End Class

    <Task()> Public Shared Function GetAlmacenDeposito(ByVal IDDeposito As String, ByVal services As ServiceProvider) As String
        Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
        Dim DptoInfo As BdgDepositoInfo = Depositos.GetEntity(IDDeposito)

        Dim IDAlmacen As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf BdgNave.Almacen, DptoInfo.IDNave & String.Empty, services)
        If Length(IDAlmacen) = 0 Then
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            IDAlmacen = AppParams.Almacen
        End If
        If Length(IDAlmacen) = 0 Then Throw New ApplicationException("No hay un almacén predeterminado para el depósito")

        Return IDAlmacen
    End Function

    <Task()> Public Shared Function CrearVino(ByVal data As StCrearVino, ByVal services As ServiceProvider) As Guid
        Dim rwDep As DataRow = New BdgDeposito().GetItemRow(data.IDDeposito)
        'TODO buscar parametro

        Dim IDAlmacenDeposito As String = ProcessServer.ExecuteTask(Of String, String)(AddressOf GetAlmacenDeposito, data.IDDeposito & String.Empty, services)
        If Length(data.IDAlmacen) > 0 AndAlso data.IDAlmacen <> IDAlmacenDeposito Then

            ApplicationService.GenerateError("El Depósito {0} de Bodega está asociado al Almacén {1}, no se puede realizar el Movimiento de la Ubicación {2} con el Almacén {3}.{4}Artículo: {5}, Lote: {6}, Cantidad: {7}.", _
                                             Quoted(data.IDDeposito), Quoted(IDAlmacenDeposito), Quoted(data.IDDeposito), Quoted(data.IDAlmacen), vbNewLine, Quoted(data.IDArticulo), Quoted(data.Lote), Quoted(data.Cantidad))
        End If
        Dim StCheck As New StCheckCapVinoArticulo(rwDep.Table, data.Cantidad, data.Origen, data.IDArticulo, data.UDMedida)
        ProcessServer.ExecuteTask(Of StCheckCapVinoArticulo)(AddressOf CheckCapacityVinoArticulo, StCheck, services)

        Dim oVinoDat As New VinoData(data.IDDeposito, rwDep(_D.TipoDeposito), data.IDArticulo, data.Fecha, data.Origen, data.UDMedida, data.Lote, IDAlmacenDeposito)
        oVinoDat.IDEstadoVino = data.IDEstadoVino
        oVinoDat.IDBarrica = data.IDBarrica
        oVinoDat.NOperacion = data.NOperacion
        oVinoDat.Estructura = data.Estrct

        Dim StCrear As New StCrearVinoMultiples(oVinoDat, data.Cantidad, rwDep(_D.MultiplesVinos))
        Dim IDVinoRslt As Guid = ProcessServer.ExecuteTask(Of StCrearVinoMultiples, Guid)(AddressOf CrearVinoMultiples, StCrear, services)

        Dim StGetFactor As New StGetFactorConversion(data.IDArticulo, data.UDMedida, rwDep(_D.IDUDMedida) & String.Empty)

        rwDep(_D.Ocupacion) += data.Cantidad * ProcessServer.ExecuteTask(Of StGetFactorConversion, Double)(AddressOf GetFactorConversion, StGetFactor, services)
        BusinessHelper.UpdateTable(rwDep.Table)

        Dim Depositos As EntityInfoCache(Of BdgDepositoInfo) = services.GetService(Of EntityInfoCache(Of BdgDepositoInfo))()
        Dim DptoInfo As BdgDepositoInfo = Depositos.GetEntity(data.IDDeposito)
        DptoInfo.Ocupacion = Nz(rwDep(_D.Ocupacion), 0)

        Return IDVinoRslt
    End Function

    <Serializable()> _
    Public Class StCrearVinoOrigenCantidad
        Public IDDeposito As String
        Public IDArticulo As String
        Public Lote As String
        Public Fecha As Date
        Public Origen As BdgOrigenVino
        Public Cantidad As Double
        Public IDAlmacen As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal IDArticulo As String, ByVal Lote As String, ByVal Fecha As Date, _
                       ByVal Origen As BdgOrigenVino, ByVal Cantidad As Double)
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.Fecha = Fecha
            Me.Origen = Origen
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoOrigenCantidad(ByVal data As StCrearVinoOrigenCantidad, ByVal services As ServiceProvider) As Guid
        Dim rwArt As DataRow = New Articulo().GetItemRow(data.IDArticulo)
        Dim strIDUdInt As String = String.Empty
        If Not rwArt.IsNull("IDUDInterna") Then strIDUdInt = rwArt("IDUDInterna")
        Dim StCrear As New StCrearVino(data.IDDeposito, data.IDArticulo, data.Lote, data.Fecha, data.Origen, strIDUdInt, data.Cantidad)
        If Length(data.IDAlmacen) > 0 Then StCrear.IDAlmacen = data.IDAlmacen
        Return ProcessServer.ExecuteTask(Of StCrearVino, Guid)(AddressOf CrearVino, StCrear, services)
    End Function

    <Serializable()> _
    Public Class StCrearVinoMultiples
        Public oVinoDat As VinoData
        Public Cantidad As Double
        Public MultiplesVinos As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal oVinoDat As VinoData, ByVal Cantidad As Double, ByVal MultiplesVinos As Boolean)
            Me.oVinoDat = oVinoDat
            Me.Cantidad = Cantidad
            Me.MultiplesVinos = MultiplesVinos
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoMultiples(ByVal data As StCrearVinoMultiples, ByVal services As ServiceProvider) As Guid
        Dim blVinoExistente As Boolean
        Dim blVinoDistintoEnDeposito As Boolean

        Dim rwDepV As DataRow
        Dim dtDepV As DataTable

        'Comprobar si el vino está vivo
        Dim StExiste As New StExisteMismoVinoVivo(data.oVinoDat.IDDeposito, data.oVinoDat.IDArticulo, data.oVinoDat.Lote, data.oVinoDat.IDAlmacen)
        rwDepV = ProcessServer.ExecuteTask(Of StExisteMismoVinoVivo, DataRow)(AddressOf ExisteMismoVinoVivo, StExiste, services)
        If rwDepV Is Nothing Then
            blVinoExistente = False
        Else
            blVinoExistente = True
        End If

        If Not data.MultiplesVinos And Not blVinoExistente Then
            'En los depósitos que sólo permiten un vino comprobamos si tienen un vino diferente al que queremos crear.
            dtDepV = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgDepositoVino.SelOnIDDeposito, data.oVinoDat.IDDeposito, services)
            If dtDepV.Rows.Count = 0 Then
                '//el deposito está vacio
                blVinoDistintoEnDeposito = False
            Else
                blVinoDistintoEnDeposito = True
            End If
        End If

        If Not blVinoExistente And Not blVinoDistintoEnDeposito Then
            Dim StCrear As New StCrearVinoYMeterEnDeposito(data.oVinoDat, data.Cantidad)
            Return ProcessServer.ExecuteTask(Of StCrearVinoYMeterEnDeposito, Guid)(AddressOf CrearVinoYMeterEnDeposito, StCrear, services)
        Else
            'Preparar datos del vino existente
            Dim IDVinoActual As Guid
            Dim CantidadActual As Double
            Dim rwVinoActual As DataRow
            If blVinoDistintoEnDeposito Then
                If Not dtDepV Is Nothing AndAlso (dtDepV.Rows.Count > 0) Then
                    IDVinoActual = dtDepV.Rows(0)(_DV.IDVino)
                    CantidadActual = dtDepV.Rows(0)(_DV.Cantidad)
                    rwVinoActual = New BdgVino().GetItemRow(IDVinoActual)
                End If
            Else
                IDVinoActual = rwDepV(_DV.IDVino)
                CantidadActual = rwDepV(_DV.Cantidad)
                rwVinoActual = New BdgVino().GetItemRow(IDVinoActual)
            End If

            'VinoContenidoTieneSalidas

            Dim StSal As New StVinoContenidoTienesalidas(IDVinoActual)
            Dim StEsMismo As New StEsMismoOrigenArtLote(data.oVinoDat.Origen, data.oVinoDat.IDArticulo, data.oVinoDat.Lote, rwVinoActual.Table, data.oVinoDat.IDAlmacen, data.oVinoDat.IDDeposito)
            If ProcessServer.ExecuteTask(Of StVinoContenidoTienesalidas, Boolean)(AddressOf VinoContenidoTieneSalidas, StSal, services) Then
                If ProcessServer.ExecuteTask(Of BdgOrigenVino, Boolean)(AddressOf OrigenEsUvaOCompra, data.oVinoDat.Origen, services) Then
                    '//Crear dos vinos. 
                    '//Uno identifica lo que entra y otro la mezcla de lo que entra con el contenido
                    Dim StCrear As New StCrearVinoOrigenYVinoMezcla(IDVinoActual, CantidadActual, data.oVinoDat, data.Cantidad)
                    Return ProcessServer.ExecuteTask(Of StCrearVinoOrigenYVinoMezcla, Guid)(AddressOf CrearVinoOrigenYVinoMezcla, StCrear, services)
                Else
                    Dim StCrear As New StCrearVinoMezcla(IDVinoActual, CantidadActual, data.oVinoDat, data.Cantidad)
                    Return ProcessServer.ExecuteTask(Of StCrearVinoMezcla, Guid)(AddressOf CrearVinoMezcla, StCrear, services)
                End If
            ElseIf ProcessServer.ExecuteTask(Of StEsMismoOrigenArtLote, Boolean)(AddressOf EsMismoOrigenArticuloLote, StEsMismo, services) Then
                If ProcessServer.ExecuteTask(Of BdgOrigenVino, Boolean)(AddressOf OrigenEsUvaOCompra, data.oVinoDat.Origen, services) Then
                    '//El vino a devolver es el contenido
                    Dim StSumar As New StSumarCantidadAContenido(IDVinoActual, data.Cantidad)
                    ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)

                    Return IDVinoActual
                Else
                    Dim StAñadir As New StAñadirOrigenComoComponente(IDVinoActual, data.oVinoDat.Origen, data.oVinoDat.NOperacion, data.oVinoDat.Estructura, data.Cantidad)
                    Return ProcessServer.ExecuteTask(Of StAñadirOrigenComoComponente, Guid)(AddressOf AñadirOrigenComoComponente, StAñadir, services)
                End If
            ElseIf data.oVinoDat.Origen = BdgOrigenVino.Compra Then
                'Si el vino que entra es de Compra de Vino, el vino resultante debe tener el Artículo y el Lote del Vino que entra.
                '//Crear dos vinos. 
                '//Uno identifica lo que entra y otro la mezcla de lo que entra con el contenido
                Dim StCrear As New StCrearVinoOrigenYVinoMezcla(IDVinoActual, CantidadActual, data.oVinoDat, data.Cantidad)
                Return ProcessServer.ExecuteTask(Of StCrearVinoOrigenYVinoMezcla, Guid)(AddressOf CrearVinoOrigenYVinoMezcla, StCrear, services)
            ElseIf ProcessServer.ExecuteTask(Of BdgOrigenVino, Boolean)(AddressOf OrigenEsUvaOCompra, data.oVinoDat.Origen, services) Then
                If ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf VinoContenidoEsInterno, rwVinoActual, services) Then
                    Dim StMismo As New StComponenteConMismoOrigenArtLoteSinSal(IDVinoActual, data.oVinoDat.Origen, data.oVinoDat.IDArticulo, data.oVinoDat.Lote, data.oVinoDat.IDAlmacen, data.oVinoDat.IDDeposito)
                    Dim IDComponente As Guid = ProcessServer.ExecuteTask(Of StComponenteConMismoOrigenArtLoteSinSal, Guid)(AddressOf ComponenteConMismoOrigenArticuloLoteSinSalidas, StMismo, services)
                    If IDComponente.Equals(Guid.Empty) Then
                        '//Crear un nuevo vino componente
                        Dim StCrear As New StCrearVinoOrigenComoComp(IDVinoActual, CantidadActual, data.oVinoDat, data.Cantidad)
                        Return ProcessServer.ExecuteTask(Of StCrearVinoOrigenComoComp, Guid)(AddressOf CrearVinoOrigenComoComponente, StCrear, services)
                    Else
                        '//El vino a devolver es el componente
                        Dim StSumar As New StSumarCantidadAComponente(IDVinoActual, IDComponente, data.Cantidad)
                        Return ProcessServer.ExecuteTask(Of StSumarCantidadAComponente, Guid)(AddressOf SumarCantidadAComponente, StSumar, services)
                    End If
                Else
                    '//Crear dos vinos. 
                    '//Uno identifica lo que entra y otro la mezcla de lo que entra con el contenido
                    Dim StCrear As New StCrearVinoOrigenYVinoMezcla(IDVinoActual, CantidadActual, data.oVinoDat, data.Cantidad)
                    Return ProcessServer.ExecuteTask(Of StCrearVinoOrigenYVinoMezcla, Guid)(AddressOf CrearVinoOrigenYVinoMezcla, StCrear, services)
                End If
            Else
                Dim StMezcla As New StCrearVinoMezcla(IDVinoActual, CantidadActual, data.oVinoDat, data.Cantidad)
                Return ProcessServer.ExecuteTask(Of StCrearVinoMezcla, Guid)(AddressOf CrearVinoMezcla, StMezcla, services)
            End If
        End If
    End Function

    <Serializable()> _
    Public Class StIncrementarIDVino
        Public IDVino As Guid
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Cantidad As Double)
            Me.IDVino = IDVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub IncrementarCantidadIDVino(ByVal data As StIncrementarIDVino, ByVal services As ServiceProvider)
        '//Cantidad, es el incremento de cantidad respecto a la que ya había. Puede ser negativo
        Dim rwVino As DataRow = New BdgVino().GetItemRow(data.IDVino)
        Dim Origen As BdgOrigenVino = rwVino(_V.Origen)
        Dim rwDep As DataRow = New BdgDeposito().GetItemRow(rwVino(_V.IDDeposito))
        Dim StGetFactor As New StGetFactorConversion(rwVino(_V.IDArticulo), rwVino(_V.IDUdMedida) & String.Empty, rwDep(_D.IDUDMedida) & String.Empty)
        Dim DblCantidad As Double = data.Cantidad
        DblCantidad = data.Cantidad * ProcessServer.ExecuteTask(Of StGetFactorConversion, Double)(AddressOf GetFactorConversion, StGetFactor, services)

        Dim StCheck As New StCheckCapOrigen(rwDep.Table, DblCantidad, Origen)
        ProcessServer.ExecuteTask(Of StCheckCapOrigen)(AddressOf CheckCapacityOrigen, StCheck, services)

        Dim StInc As New StIncrementarCantidad(rwVino.Table, data.Cantidad)
        ProcessServer.ExecuteTask(Of StIncrementarCantidad)(AddressOf IncrementarCantidad, StInc, services)
        rwDep(_D.Ocupacion) += DblCantidad
        BusinessHelper.UpdateTable(rwDep.Table)
    End Sub

    <Serializable()> _
    Public Class StIncrementarCantidad
        Public DtVino As DataTable
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtVino As DataTable, ByVal Cantidad As Double)
            Me.DtVino = DtVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub IncrementarCantidad(ByVal data As StIncrementarCantidad, ByVal services As ServiceProvider)
        Dim IDVino As Guid = data.DtVino.Rows(0)(_V.IDVino)

        Dim dtDepV As DataTable = New BdgDepositoVino().SelOnPrimaryKey(IDVino)
        If dtDepV.Rows.Count = 0 Then
            '//El vino no esta en el deposito
            If ProcessServer.ExecuteTask(Of BdgOrigenVino, Boolean)(AddressOf OrigenEsUvaOCompra, data.DtVino.Rows(0)(_V.Origen), services) Then
                Dim StModificar As New StModificarCantidadCompEnNivel(IDVino, data.Cantidad, data.DtVino.Rows(0)(_V.IDDeposito))
                Dim IDVinoEnDep As Guid = ProcessServer.ExecuteTask(Of StModificarCantidadCompEnNivel, Guid)(AddressOf ModificarCantidadComponenteEnCualquierNivel, StModificar, services)
                If CType(ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgVinoEstructura.SelOnComponente, IDVino, services), DataTable).Rows.Count = 0 Then
                    ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgVino.DeleteVino, IDVino, services)
                    Dim rwVinoEnDep As DataRow = New BdgVino().GetItemRow(IDVinoEnDep)
                    Dim rwVinoEnDep2 As DataRow = ProcessServer.ExecuteTask(Of DataRow, DataRow)(AddressOf ReducirEstructura, rwVinoEnDep, services)
                    Dim StReemp As New StReemplazar(rwVinoEnDep.Table, rwVinoEnDep2.Table)
                    ProcessServer.ExecuteTask(Of StReemplazar)(AddressOf ReemplazarVinoEnDeposito, StReemp, services)
                    IDVinoEnDep = rwVinoEnDep2(_V.IDVino)
                End If
                Dim StSumar As New StSumarCantidadAContenido(IDVinoEnDep, data.Cantidad)
                ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)
            Else
                Dim stError As String = "El vino ya no se encuentra en el depósito."
                Dim StMensaje As New StMensajeError(IDVino, stError)
                stError = String.Format(ProcessServer.ExecuteTask(Of StMensajeError, String)(AddressOf MensajeErrorVino, StMensaje, services))
                ApplicationService.GenerateError(stError)
            End If
        Else
            '//El vino está en el depósito
            Dim rwDepV As DataRow = dtDepV.Rows(0)
            Dim dtDepVino As DataTable = dtDepV.Clone
            dtDepVino.ImportRow(rwDepV)
            Dim DtVin As DataTable = data.DtVino.Clone
            DtVin.ImportRow(data.DtVino.Rows(0))
            Dim StCambiar As New StCambiarOcupYStock(dtDepVino, data.Cantidad, DtVin)
            ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambiar, services)

            If dtDepVino.Rows.Count = 0 Then  '//Al cambiar la ocupación, se ha quedado vacío el depósito, entonces eliminamos el vino (si no tiene salidas)
                Dim StVino As New StVinoContenidoTienesalidas(IDVino)
                If Not ProcessServer.ExecuteTask(Of StVinoContenidoTienesalidas, Boolean)(AddressOf VinoContenidoTieneSalidas, StVino, services) Then
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf BdgVino.DeleteRw, data.DtVino.Rows(0), services)
                End If
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class StReemplazar
        Public DtVinoEnDep As DataTable
        Public DtNuevoVinoEnDep As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtVinoEnDep As DataTable, ByVal DtNuevoVinoEnDep As DataTable)
            Me.DtVinoEnDep = DtVinoEnDep
            Me.DtNuevoVinoEnDep = DtNuevoVinoEnDep
        End Sub
    End Class

    <Task()> Public Shared Sub ReemplazarVinoEnDeposito(ByVal data As StReemplazar, ByVal services As ServiceProvider)
        If DirectCast(data.DtVinoEnDep.Rows(0)(_V.IDVino), Guid).Equals(data.DtNuevoVinoEnDep.Rows(0)(_V.IDVino)) Then
        Else
            Dim rwDepV As DataRow = New BdgDepositoVino().GetItemRow(data.DtVinoEnDep.Rows(0)(_V.IDVino))
            If data.DtVinoEnDep.Rows(0)(_V.IDArticulo) = data.DtNuevoVinoEnDep.Rows(0)(_V.IDArticulo) AndAlso data.DtVinoEnDep.Rows(0)(_V.Lote) = data.DtNuevoVinoEnDep.Rows(0)(_V.Lote) Then
                rwDepV(_DV.IDVino) = data.DtNuevoVinoEnDep.Rows(0)(_V.IDVino)
                BusinessHelper.UpdateTable(rwDepV.Table)
            Else
                Dim Cantidad As Double = rwDepV(_DV.Cantidad)
                Dim DtVin As DataTable = data.DtVinoEnDep.Clone
                DtVin.ImportRow(data.DtVinoEnDep.Rows(0))
                Dim StCambio1 As New StCambiarOcupYStock(rwDepV.Table, -Cantidad, DtVin)
                ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambio1, services)
                Dim StGet As New StGetNew(Nothing, data.DtNuevoVinoEnDep.Rows(0)(_V.IDVino), data.DtNuevoVinoEnDep.Rows(0)(_V.IDDeposito))
                Dim DtVin2 As DataTable = data.DtNuevoVinoEnDep.Clone
                DtVin2.ImportRow(data.DtNuevoVinoEnDep.Rows(0))
                Dim StCambio2 As New StCambiarOcupYStock(ProcessServer.ExecuteTask(Of StGetNew, DataRow)(AddressOf GetNewDepVRow, StGet, services).Table, Cantidad, DtVin2)
                ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambio2, services)
            End If
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf BdgVino.DeleteRw, data.DtVinoEnDep.Rows(0), services)
        End If
    End Sub

    <Serializable()> _
    Public Class StCrearVinoMezcla
        Public IDVinoContenido As Guid
        Public CantidadActual As Double
        Public oVino As VinoData
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVinoContenido As Guid, ByVal CantidadActual As Double, ByVal oVino As VinoData, ByVal Cantidad As Double)
            Me.IDVinoContenido = IDVinoContenido
            Me.CantidadActual = CantidadActual
            Me.oVino = oVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoMezcla(ByVal data As StCrearVinoMezcla, ByVal services As ServiceProvider) As Guid
        If data.oVino.Estructura Is Nothing Then
            ReDim data.oVino.Estructura(0)
        Else
            ReDim Preserve data.oVino.Estructura(data.oVino.Estructura.Length)
        End If
        data.oVino.Estructura(data.oVino.Estructura.Length - 1) = New VinoComponente(data.IDVinoContenido, data.CantidadActual, 1)

        ProcessServer.ExecuteTask(Of VinoData)(AddressOf CrearVinoYEstructura, data.oVino, services)

        '//salida del vino contenido
        Dim rwDepV As DataRow = New BdgDepositoVino().GetItemRow(data.IDVinoContenido)
        Dim StCambio1 As New StCambiarOcupYStock(rwDepV.Table, -data.CantidadActual, New BdgVino().GetItemRow(data.IDVinoContenido).Table)
        ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambio1, services)

        '//entrada del vino mezcla
        Dim StGet As New StGetNew(rwDepV.Table, data.oVino.IDVino, data.oVino.IDDeposito)
        rwDepV = ProcessServer.ExecuteTask(Of StGetNew, DataRow)(AddressOf GetNewDepVRow, StGet, services)

        Dim StCambio2 As New StCambiarOcupYStock(rwDepV.Table, data.CantidadActual + data.Cantidad, New BdgVino().GetItemRow(data.oVino.IDVino).Table)
        ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambio2, services)

        Return data.oVino.IDVino
    End Function

    <Serializable()> _
    Public Class StSumarCantidadAComponente
        Public IDVino As Guid
        Public IDVinoComponente As Guid
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal IDVinoComponente As Guid, ByVal Cantidad As Double)
            Me.IDVino = IDVino
            Me.IDVinoComponente = IDVinoComponente
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function SumarCantidadAComponente(ByVal data As StSumarCantidadAComponente, ByVal services As ServiceProvider) As Guid
        Dim rwVE As DataRow = New BdgVinoEstructura().GetItemRow(data.IDVino, data.IDVinoComponente, "Entrada")
        rwVE(_VE.Cantidad) += data.Cantidad
        BusinessHelper.UpdateTable(rwVE.Table)

        Dim StSumar As New StSumarCantidadAContenido(data.IDVino, data.Cantidad)
        ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)

        Return data.IDVinoComponente
    End Function

    <Serializable()> _
    Public Class StSumarCantidadAContenido
        Public IDVinoContenido As Guid
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVinoContenido As Guid, ByVal Cantidad As Double)
            Me.IDVinoContenido = IDVinoContenido
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub SumarCantidadAContenido(ByVal data As StSumarCantidadAContenido, ByVal services As ServiceProvider)
        Dim StCambio As New StCambiarOcupYStock(New BdgDepositoVino().GetItemRow(data.IDVinoContenido).Table, data.Cantidad, New BdgVino().GetItemRow(data.IDVinoContenido).Table)
        ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambio, services)
    End Sub

    <Serializable()> _
    Public Class StCrearVinoOrigenComoComp
        Public IDVinoContenido As Guid
        Public CantidadActual As Double
        Public oVino As VinoData
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVinoContenido As Guid, ByVal CantidadActual As Double, ByVal oVino As VinoData, ByVal Cantidad As Double)
            Me.IDVinoContenido = IDVinoContenido
            Me.CantidadActual = CantidadActual
            Me.oVino = oVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoOrigenComoComponente(ByVal data As StCrearVinoOrigenComoComp, ByVal services As ServiceProvider) As Guid
        '//vino origen (uva o compra)
        ProcessServer.ExecuteTask(Of VinoData)(AddressOf CrearVinoYEstructura, data.oVino, services)
        Dim StEst As New StCrearEstructuraOperacion(data.IDVinoContenido, data.oVino.Origen, data.oVino.NOperacion, New VinoComponente() {New VinoComponente(data.oVino.IDVino, data.Cantidad, 1)})
        ProcessServer.ExecuteTask(Of StCrearEstructuraOperacion)(AddressOf CrearEstructuraOperacion, StEst, services)
        Dim StSumar As New StSumarCantidadAContenido(data.IDVinoContenido, data.Cantidad)
        ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)

        Return data.oVino.IDVino

    End Function

    <Serializable()> _
    Public Class StCrearVinoOrigenYVinoMezcla
        Public IDVinoContenido As Guid
        Public CantidadActual As Double
        Public oVino As VinoData
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVinoContenido As Guid, ByVal CantidadActual As Double, ByVal oVino As VinoData, ByVal Cantidad As Double)
            Me.IDVinoContenido = IDVinoContenido
            Me.CantidadActual = CantidadActual
            Me.oVino = oVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoOrigenYVinoMezcla(ByVal data As StCrearVinoOrigenYVinoMezcla, ByVal services As ServiceProvider) As Guid
        '//vino origen (uva o compra)
        ProcessServer.ExecuteTask(Of VinoData)(AddressOf CrearVinoYEstructura, data.oVino, services)

        Dim oMezcla As VinoData = data.oVino.Clone()
        Dim Estructura() As VinoComponente = {New VinoComponente(data.oVino.IDVino, data.Cantidad, 1)}
        oMezcla.Estructura = Estructura

        Dim StMezcla As New StCrearVinoMezcla(data.IDVinoContenido, data.CantidadActual, oMezcla, data.Cantidad)
        ProcessServer.ExecuteTask(Of StCrearVinoMezcla)(AddressOf CrearVinoMezcla, StMezcla, services)

        Return data.oVino.IDVino
    End Function

    <Serializable()> _
    Public Class StCrearVinoYMeterEnDeposito
        Public oVino As VinoData
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal oVino As VinoData, ByVal Cantidad As Double)
            Me.oVino = oVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function CrearVinoYMeterEnDeposito(ByVal data As StCrearVinoYMeterEnDeposito, ByVal services As ServiceProvider) As Guid
        ProcessServer.ExecuteTask(Of VinoData)(AddressOf CrearVinoYEstructura, data.oVino, services)
        Dim StMeter As New StMeterVinoEnDeposito(data.oVino.IDVino, data.oVino.IDDeposito, data.Cantidad)
        ProcessServer.ExecuteTask(Of StMeterVinoEnDeposito)(AddressOf MeterVinoEnDeposito, StMeter, services)
        Return data.oVino.IDVino
    End Function

    <Serializable()> _
    Public Class StMeterVinoEnDeposito
        Public IDVino As Guid
        Public IDDeposito As String
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal IDDeposito As String, ByVal Cantidad As Double)
            Me.IDVino = IDVino
            Me.IDDeposito = IDDeposito
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub MeterVinoEnDeposito(ByVal data As StMeterVinoEnDeposito, ByVal services As ServiceProvider)
        Dim StGet As New StGetNew(Nothing, data.IDVino, data.IDDeposito)
        Dim f As New Filter
        f.Add(New GuidFilterItem("IDVino", data.IDVino))
        Dim dtDptoVino As DataTable = New BdgDepositoVino().Filter(f)
        Dim rwDepV As DataRow
        If dtDptoVino.Rows.Count > 0 Then
            rwDepV = dtDptoVino.Rows(0)
        Else
            rwDepV = ProcessServer.ExecuteTask(Of StGetNew, DataRow)(AddressOf GetNewDepVRow, StGet, services)
        End If
        Dim StCambio As New StCambiarOcupYStock(rwDepV.Table, data.Cantidad, New BdgVino().GetItemRow(data.IDVino).Table)
        ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambio, services)
    End Sub

    <Serializable()> _
    Public Class StGetNew
        Public DtDepV As DataTable
        Public IDVino As Guid
        Public IDDeposito As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtDepV As DataTable, ByVal IDVino As Guid, ByVal IDDeposito As String)
            Me.DtDepV = DtDepV
            Me.IDVino = IDVino
            Me.IDDeposito = IDDeposito
        End Sub
    End Class

    <Task()> Public Shared Function GetNewDepVRow(ByVal data As StGetNew, ByVal services As ServiceProvider) As DataRow
        If data.DtDepV Is Nothing Then
            data.DtDepV = New BdgDepositoVino().AddNew
        End If

        Dim rwDepV As DataRow = data.DtDepV.NewRow
        rwDepV(_DV.IDVino) = data.IDVino
        rwDepV(_DV.IDDeposito) = data.IDDeposito
        rwDepV(_DV.Cantidad) = 0
        data.DtDepV.Rows.Add(rwDepV)

        Return rwDepV
    End Function

    <Task()> Public Shared Sub CrearVinoYEstructura(ByVal oVino As VinoData, ByVal services As ServiceProvider)
        Dim dtVino As DataTable = New BdgVino().AddNew
        AddVinoToDT(dtVino, oVino)
        BusinessHelper.UpdateTable(dtVino)
        Dim StEst As New StCrearEstructura(oVino)
        ProcessServer.ExecuteTask(Of StCrearEstructura)(AddressOf CrearEstructura, StEst, services)
    End Sub

    <Serializable()> _
    Public Class StAñadirOrigenComoComponente
        Public IDVino As Guid
        Public Origen As BdgOrigenVino
        Public NOperacion As String
        Public Estructura() As VinoComponente
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Origen As BdgOrigenVino, ByVal NOperacion As String, ByVal Estructura() As VinoComponente, ByVal Cantidad As Double)
            Me.IDVino = IDVino
            Me.Origen = Origen
            Me.NOperacion = NOperacion
            Me.Estructura = Estructura
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function AñadirOrigenComoComponente(ByVal data As StAñadirOrigenComoComponente, ByVal services As ServiceProvider) As Guid
        Dim StCrear As New StCrearEstructuraOperacion(data.IDVino, data.Origen, data.NOperacion, data.Estructura)
        ProcessServer.ExecuteTask(Of StCrearEstructuraOperacion)(AddressOf CrearEstructuraOperacion, StCrear, services)
        Dim StSumar As New StSumarCantidadAContenido(data.IDVino, data.Cantidad)
        ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)
        Return data.IDVino
    End Function

    <Serializable()> _
    Public Class StVinoContenidoTienesalidas
        Public IDVino As Guid
        Public StOperaciones As String

        Public Sub New(ByVal IDVino As Guid, Optional ByVal StOperaciones As String = "")
            Me.IDVino = IDVino
            Me.StOperaciones = StOperaciones
        End Sub
    End Class

    <Task()> Public Shared Function VinoContenidoTieneSalidas(ByVal data As StVinoContenidoTienesalidas, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgVinoEstructura.SelOnComponente, data.IDVino, services)
        Dim Operaciones As List(Of String) = (From c In dt Where Not c.IsNull(_VE.Operacion) Select CStr(c(_VE.Operacion)) Distinct).ToList

        Dim f As New Filter
        f.Add(New GuidFilterItem("IDVino", data.IDVino))
        f.Add(New BooleanFilterItem("Destino", False))
        Dim dtOpVinoEnOrigen As DataTable = New BE.DataEngine().Filter("frmBdgOperacionVinoOrigen", f)
        If dtOpVinoEnOrigen.Rows.Count > 0 Then
            Dim OperacionesVinoEnOrigen As List(Of String) = (From c In dtOpVinoEnOrigen Where Not c.IsNull("NOperacion") Select CStr(c("NOperacion")) Distinct).ToList
            If Not OperacionesVinoEnOrigen Is Nothing AndAlso OperacionesVinoEnOrigen.Count > 0 Then
                For Each Op As String In OperacionesVinoEnOrigen
                    If Operaciones.IndexOf(Op) < 0 Then
                        Operaciones.Add(Op)
                    End If
                Next
            End If
        End If

        If Not Operaciones Is Nothing AndAlso Operaciones.Count > 0 Then
            data.StOperaciones = Strings.Join(Operaciones.ToArray, ",")
        End If

        Return (Not Operaciones Is Nothing AndAlso Operaciones.Count > 0)
    End Function

    <Serializable()> _
    Public Class StVinoComponenteTieneSalidas
        Public IDVino As Guid
        Public IDVinoComponente As Guid

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal IDVinoComponente As Guid)
            Me.IDVino = IDVino
            Me.IDVinoComponente = IDVinoComponente
        End Sub
    End Class

    <Task()> Public Shared Function VinoComponenteTieneSalidas(ByVal data As StVinoComponenteTieneSalidas, ByVal services As ServiceProvider) As Boolean
        Dim oFltr As New Filter
        oFltr.Add(New GuidFilterItem(_VE.IDVino, FilterOperator.NotEqual, data.IDVino))
        oFltr.Add(New GuidFilterItem(_VE.IDVinoComponente, data.IDVinoComponente))
        Dim dt As DataTable = New BdgVinoEstructura().Filter(oFltr)
        Return dt.Rows.Count <> 0
    End Function

    <Task()> Public Shared Function OrigenEsUvaOCompra(ByVal Origen As BdgOrigenVino, ByVal services As ServiceProvider) As Boolean
        Return Origen = BdgOrigenVino.Uva OrElse Origen = BdgOrigenVino.Compra
    End Function

    <Serializable()> _
    Public Class StExisteMismoVinoVivo
        Public IDDeposito As String
        Public IDArticulo As String
        Public Lote As String
        Public IDAlmacen As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal IDArticulo As String, ByVal Lote As String, ByVal IDAlmacen As String)
            Me.IDDeposito = IDDeposito
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Function ExisteMismoVinoVivo(ByVal data As StExisteMismoVinoVivo, ByVal services As ServiceProvider) As DataRow
        Dim oFltr As New Filter
        oFltr.Add(New StringFilterItem(_V.IDDeposito, data.IDDeposito))
        oFltr.Add(New StringFilterItem(_V.IDArticulo, data.IDArticulo))
        oFltr.Add(New StringFilterItem(_V.Lote, data.Lote))
        oFltr.Add(New StringFilterItem(_V.IDAlmacen, data.IDAlmacen))

        Dim dt As DataTable = AdminData.GetData("negBdgExisteMismoVinoVivo", oFltr)
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
    End Function

    <Serializable()> _
    Public Class StExisteMismoOrigenArticuloLote
        Public IDDeposito As String
        Public Origen As BdgOrigenVino
        Public IDArticulo As String
        Public Lote As String
        Public IDAlmacen As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDDeposito As String, ByVal Origen As BdgOrigenVino, ByVal IDArticulo As String, ByVal Lote As String, ByVal IDAlmacen As String)
            Me.IDDeposito = IDDeposito
            Me.Origen = Origen
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Function ExisteMismoOrigenArticuloLote(ByVal data As StExisteMismoOrigenArticuloLote, ByVal services As ServiceProvider) As DataRow
        Dim oFltr As New Filter
        oFltr.Add(New StringFilterItem(_V.IDDeposito, data.IDDeposito))
        oFltr.Add(New NumberFilterItem(_V.Origen, data.Origen))
        oFltr.Add(New StringFilterItem(_V.IDArticulo, data.IDArticulo))
        oFltr.Add(New StringFilterItem(_V.Lote, data.Lote))
        oFltr.Add(New StringFilterItem(_V.IDAlmacen, data.IDAlmacen))

        Dim dt As DataTable = AdminData.GetData("negBdgExisteMismoOrigenArticuloLote", oFltr)
        If dt.Rows.Count > 0 Then Return dt.Rows(0)
    End Function

    <Serializable()> _
    Public Class StComponenteConMismoOrigenArtLoteSinSal
        Public IDVino As Guid
        Public Origen As BdgOrigenVino
        Public IDArticulo As String
        Public Lote As String
        Public IDAlmacen As String
        Public IDDeposito As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Origen As BdgOrigenVino, ByVal IDArticulo As String, ByVal Lote As String, ByVal IDAlmacen As String, ByVal IDDeposito As String)
            Me.IDVino = IDVino
            Me.Origen = Origen
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.IDAlmacen = IDAlmacen
            Me.IDDeposito = IDDeposito
        End Sub
    End Class

    <Task()> Public Shared Function ComponenteConMismoOrigenArticuloLoteSinSalidas(ByVal data As StComponenteConMismoOrigenArtLoteSinSal, ByVal services As ServiceProvider) As Guid
        Dim oFltr As New Filter
        oFltr.Add(New GuidFilterItem(_VE.IDVino, data.IDVino))
        oFltr.Add(New NumberFilterItem(_V.Origen, data.Origen))
        oFltr.Add(New StringFilterItem(_V.IDArticulo, data.IDArticulo))
        oFltr.Add(New StringFilterItem(_V.Lote, data.Lote))
        oFltr.Add(New StringFilterItem(_V.IDAlmacen, data.IDAlmacen))
        oFltr.Add(New StringFilterItem(_V.IDDeposito, data.IDDeposito))

        Dim dt As DataTable = New BE.DataEngine().Filter("negBdgComponenteOrigenArticuloLote", oFltr)

        For Each oRw As DataRow In dt.Rows
            Dim StVinoComp As New StVinoComponenteTieneSalidas(oRw(_VE.IDVino), oRw(_VE.IDVinoComponente))
            If Not ProcessServer.ExecuteTask(Of StVinoComponenteTieneSalidas, Boolean)(AddressOf VinoComponenteTieneSalidas, StVinoComp, services) Then
                Return oRw(_VE.IDVinoComponente)
            End If
        Next
    End Function

    <Serializable()> _
    Public Class StEsMismoOrigenArtLote
        Public Origen As BdgOrigenVino
        Public IDArticulo As String
        Public Lote As String
        Public DtVinoActual As DataTable
        Public IDAlmacen As String
        Public IDDeposito As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal Origen As BdgOrigenVino, ByVal IDArticulo As String, ByVal Lote As String, ByVal DtVinoActual As DataTable, _
                       ByVal IDAlmacen As String, ByVal IDDeposito As String)
            Me.Origen = Origen
            Me.IDArticulo = IDArticulo
            Me.Lote = Lote
            Me.DtVinoActual = DtVinoActual
            Me.IDAlmacen = IDAlmacen
            Me.IDDeposito = IDDeposito
        End Sub
    End Class

    <Task()> Public Shared Function EsMismoOrigenArticuloLote(ByVal data As StEsMismoOrigenArtLote, ByVal services As ServiceProvider) As Boolean
        Return data.Origen = data.DtVinoActual.Rows(0)(_V.Origen) _
                AndAlso data.IDArticulo = data.DtVinoActual.Rows(0)(_V.IDArticulo) _
                AndAlso data.Lote = data.DtVinoActual.Rows(0)(_V.Lote) _
                AndAlso data.IDAlmacen = data.DtVinoActual.Rows(0)(_V.IDAlmacen) _
                AndAlso data.IDDeposito = data.DtVinoActual.Rows(0)(_V.IDDeposito)
    End Function

    <Task()> Public Shared Function VinoContenidoEsInterno(ByVal DrVinoActual As DataRow, ByVal services As ServiceProvider) As Boolean
        Return DrVinoActual(_V.Origen) = BdgOrigenVino.Interno
    End Function

    Public Shared Sub AddVinoToDT(ByVal dtVino As DataTable, _
                                    ByVal oVinoDat As VinoData)

        Dim rwVino As DataRow = dtVino.NewRow
        rwVino = dtVino.NewRow
        rwVino(_V.IDVino) = oVinoDat.IDVino
        rwVino(_V.IDDeposito) = oVinoDat.IDDeposito
        rwVino(_V.TipoDeposito) = oVinoDat.TipoDeposito
        rwVino(_V.IDArticulo) = oVinoDat.IDArticulo
        rwVino(_V.Lote) = oVinoDat.Lote
        rwVino(_V.Fecha) = oVinoDat.Fecha
        If Len(oVinoDat.IDEstadoVino) Then rwVino(_V.IDEstadoVino) = oVinoDat.IDEstadoVino
        rwVino(_V.Origen) = oVinoDat.Origen
        If Len(oVinoDat.NOperacion) <> 0 Then rwVino(_V.NOperacion) = oVinoDat.NOperacion
        rwVino(_V.IDUdMedida) = oVinoDat.IDUdMedida
        If Len(oVinoDat.IDBarrica) > 0 Then rwVino(_V.IDBarrica) = oVinoDat.IDBarrica
        rwVino(_V.IDAlmacen) = oVinoDat.IDAlmacen
        dtVino.Rows.Add(rwVino)

    End Sub

    <Serializable()> _
    Public Class StCrearEstructura
        Public oVinoDat As VinoData
        Public Dels As List(Of Guid)

        Public Sub New()
        End Sub

        Public Sub New(ByVal oVinoDat As VinoData, Optional ByVal Dels As List(Of Guid) = Nothing)
            Me.oVinoDat = oVinoDat
            Me.Dels = Dels
        End Sub
    End Class

    <Task()> Public Shared Sub CrearEstructura(ByVal data As StCrearEstructura, ByVal services As ServiceProvider)
        Dim StEst As New StCrearEstructuraOperacion(data.oVinoDat.IDVino, data.oVinoDat.Origen, data.oVinoDat.NOperacion, data.oVinoDat.Estructura, data.Dels)
        ProcessServer.ExecuteTask(Of StCrearEstructuraOperacion)(AddressOf CrearEstructuraOperacion, StEst, services)
    End Sub

    <Serializable()> _
    Public Class StCrearEstructuraOperacion
        Public IDVino As Guid
        Public NOperacion As String
        Public Estructura() As VinoComponente
        Public Dels As List(Of Guid)
        Public Origen As BdgOrigenVino

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Origen As BdgOrigenVino, ByVal NOperacion As String, ByVal Estructura() As VinoComponente, Optional ByVal Dels As List(Of Guid) = Nothing)
            Me.IDVino = IDVino
            Me.Origen = Origen
            Me.NOperacion = NOperacion
            Me.Estructura = Estructura
            Me.Dels = Dels
        End Sub
    End Class

    <Task()> Public Shared Sub CrearEstructuraOperacion(ByVal data As StCrearEstructuraOperacion, ByVal services As ServiceProvider)
        Dim strNOperacion As String = String.Empty
        Dim strOperacion As String = String.Empty

        If data.Origen = BdgOrigenVino.AlbaranTransferencia Then
            strNOperacion = String.Empty
            If Len(data.NOperacion) > 0 Then
                strOperacion = data.NOperacion
            Else
                strOperacion = "AVTransfer"
            End If
        Else
            strNOperacion = data.NOperacion
            If Len(strNOperacion) > 0 Then
                strOperacion = strNOperacion
            Else
                strOperacion = "Entrada"
            End If
        End If

        Dim StEst As New BdgVinoEstructura.StSelOnVinoOperacion(data.IDVino, strOperacion)
        Dim dtVE As DataTable = ProcessServer.ExecuteTask(Of BdgVinoEstructura.StSelOnVinoOperacion, DataTable)(AddressOf BdgVinoEstructura.SelOnVinoOperacion, StEst, services)
        Dim dwVE As DataView = dtVE.DefaultView
        dwVE.Sort = _VE.IDVinoComponente

        If Not data.Estructura Is Nothing Then
            For Each oVC As VinoComponente In data.Estructura
                Dim idx As Integer = dwVE.Find(oVC.IDVino)
                Dim rwVE As DataRow
                If idx >= 0 Then
                    rwVE = dwVE(idx).Row
                Else
                    rwVE = dtVE.NewRow
                    dtVE.Rows.Add(rwVE)
                    rwVE(_VE.IDVino) = data.IDVino
                    rwVE(_VE.IDVinoComponente) = oVC.IDVino

                    rwVE(_VE.Operacion) = strOperacion
                    rwVE(_VE.NOperacion) = strNOperacion

                    rwVE(_VE.Cantidad) = 0
                    rwVE(_VE.Merma) = 0
                End If
                rwVE(_VE.Factor) = oVC.Factor
                rwVE(_VE.Cantidad) = oVC.Cantidad
                rwVE(_VE.Merma) = oVC.Merma
            Next
        End If

        If Not data.Dels Is Nothing Then
            For Each IDComp As Guid In data.Dels
                Dim idx As Integer = dwVE.Find(IDComp)
                If idx >= 0 Then dwVE.Delete(idx)
            Next
        End If
        BusinessHelper.UpdateTable(dtVE)
    End Sub

    <Task()> Public Shared Function ReducirEstructura(ByVal DrVino As DataRow, ByVal services As ServiceProvider) As DataRow
        Dim rslt As DataRow = DrVino
        Dim dtVinoE As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgVinoEstructura.SelOnVino, DrVino(_V.IDVino), services)
        If dtVinoE.Rows.Count = 1 Then
            Dim drVinoE As DataRow = dtVinoE.Rows(0)
            Dim rwVinoComponente As DataRow = New BdgVino().GetItemRow(drVinoE(_VE.IDVinoComponente))
            If rwVinoComponente(_V.IDDeposito) = DrVino(_V.IDDeposito) Then
                rslt = rwVinoComponente
                drVinoE.Delete()
                BusinessHelper.UpdateTable(drVinoE.Table)
            End If
        End If
        Return rslt
    End Function

    <Task()> Public Shared Function GetLotePredeterminado(ByVal Lote As String, ByVal services As ServiceProvider) As String
        If Length(Lote) = 0 Then
            Dim AppParams As BdgParametrosOperaciones = services.GetService(Of BdgParametrosOperaciones)()
            Lote = AppParams.LotePorDefecto
            If Length(Lote) = 0 Then
                ApplicationService.GenerateError("No se ha especificado un número de lote predeterminado")
            End If
        End If
        Return Lote
    End Function

    <Serializable()> _
    Public Class StCambiarOcupacion
        Public IDVino As Guid
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Cantidad As Double)
            Me.IDVino = IDVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub CambiarOcupacion(ByVal data As StCambiarOcupacion, ByVal services As ServiceProvider)
        Dim rwVino As DataRow = New BdgVino().GetItemRow(data.IDVino)
        Dim IDDeposito As String = rwVino(_V.IDDeposito)
        Dim rwDep As DataRow = New BdgDeposito().GetItemRow(IDDeposito)

        Dim StCheck As New StCheckCapVino(rwDep, data.Cantidad, rwVino(_V.Origen), rwVino.Table)
        ProcessServer.ExecuteTask(Of StCheckCapVino)(AddressOf CheckCapacityVino, StCheck, services)

        Dim rwDepV As DataRow

        Dim dtDepV As DataTable = New BdgDepositoVino().SelOnPrimaryKey(data.IDVino)
        If dtDepV.Rows.Count = 0 Then
            If Not CBool(rwDep(_D.MultiplesVinos)) Then
                If CType(ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf BdgDepositoVino.SelOnIDDeposito, IDDeposito, services), DataTable).Rows.Count <> 0 Then
                    ApplicationService.GenerateError("El vino contenido en el depósito '|' no es el mismo que se intenta mover.", IDDeposito)
                End If
            End If
            Dim StGet As New StGetNew(dtDepV, data.IDVino, IDDeposito)
            rwDepV = ProcessServer.ExecuteTask(Of StGetNew, DataRow)(AddressOf GetNewDepVRow, StGet, services)
        Else
            rwDepV = dtDepV.Rows(0)
        End If

        Dim StCambiar As New StCambiarOcupYStock(rwDepV.Table, data.Cantidad, rwVino.Table)
        ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambiar, services)

        Dim StGetFactor As New StGetFactorConversion(rwVino(_V.IDArticulo), rwVino(_V.IDUdMedida) & String.Empty, rwDep(_D.IDUDMedida) & String.Empty)
        rwDep(_D.Ocupacion) += data.Cantidad * ProcessServer.ExecuteTask(Of StGetFactorConversion, Double)(AddressOf GetFactorConversion, StGetFactor, services)

        BusinessHelper.UpdateTable(rwDep.Table)

    End Sub


    <Serializable()> _
    Public Class StCambiarOcupYStock
        Public DtDepV As DataTable
        Public Cantidad As Double
        Public DtVino As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtDepV As DataTable, ByVal Cantidad As Double, ByVal DtVino As DataTable)
            Me.DtDepV = DtDepV
            Me.Cantidad = Cantidad
            Me.DtVino = DtVino
        End Sub
    End Class

    <Task()> Public Shared Sub CambiarOcupacionYStock(ByVal data As StCambiarOcupYStock, ByVal services As ServiceProvider)
        Dim bNuevoVino As Boolean
        Dim IDVinoTratar As Guid = data.DtDepV.Rows(0)("IDVino")

        data.DtDepV.Rows(0)(_DV.Cantidad) += data.Cantidad
        If data.DtDepV.Rows(0)(_DV.Cantidad) < 0 Then
            Dim stError As String = "No se puede realizar el movimiento porque la cantidad a mover excede la cantidad existente."
            Dim StMensaje As New StMensajeError(IDVinoTratar, stError)
            stError = String.Format(ProcessServer.ExecuteTask(Of StMensajeError, String)(AddressOf MensajeErrorVino, StMensaje, services))
            ApplicationService.GenerateError(stError)
        End If

        If data.DtDepV.Rows(0)(_DV.Cantidad) = 0 Then
            data.DtDepV.Rows(0).Delete()
            bNuevoVino = False
        Else
            bNuevoVino = True
        End If

        AdminData.SetData(data.DtDepV)

        If bNuevoVino Then 'Nuevo Vino
            ProcessServer.ExecuteTask(Of Guid)(AddressOf ActualizarDiasVino, IDVinoTratar, services)
            Dim StActua As New BdgWorkClass.StActuaQTotVino(IDVinoTratar, True)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StActuaQTotVino)(AddressOf BdgWorkClass.ActualizarQTotVino, StActua, services)
        Else 'Salida de Vino
            Dim StActua As New BdgWorkClass.StActuaQTotVino(IDVinoTratar)
            ProcessServer.ExecuteTask(Of BdgWorkClass.StActuaQTotVino)(AddressOf BdgWorkClass.ActualizarQTotVino, StActua, services)
        End If
        Dim StReg As New StRegistrarMovimiento(data.DtVino, data.Cantidad)
        ProcessServer.ExecuteTask(Of StRegistrarMovimiento)(AddressOf RegistrarMovimiento, StReg, services)
    End Sub

    <Task()> Public Shared Function QTotObtener(ByVal IDVino As Guid, ByVal services As ServiceProvider) As Double
        Dim Total As Double = 0
        Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
        SqlCmd.CommandText = "fBdgQTotVino"
        SqlCmd.CommandType = CommandType.StoredProcedure

        Dim SqlParam1 As Common.DbParameter = SqlCmd.CreateParameter
        SqlCmd.Parameters.Add(SqlParam1)
        SqlParam1.ParameterName = "@pIDVino"
        SqlParam1.Value = IDVino

        Dim sqlParam As Common.DbParameter = SqlCmd.CreateParameter
        SqlCmd.Parameters.Add(sqlParam)
        sqlParam.ParameterName = "fBdgQTotVino"
        sqlParam.Direction = ParameterDirection.ReturnValue

        Nz(AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteScalar), 0)

        Total = Nz(sqlParam.Value, 0)
        Return Total
    End Function

    <Serializable()> _
    Public Class StActuaQTotVino
        Public IDVino As Guid
        Public Nuevo As Boolean

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, Optional ByVal Nuevo As Boolean = False)
            Me.IDVino = IDVino
            Me.Nuevo = Nuevo
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarQTotVino(ByVal data As StActuaQTotVino, ByVal services As ServiceProvider)
        'Si el vino es Nuevo se pone en la tabla de vinos (BdgVino) a 0.
        'Al sacar el vino se calcula su QTot 
        Dim clsVino As New BdgVino
        Dim dtVino As DataTable = clsVino.Filter("*", "IDVino='" & data.IDVino.ToString & "'")
        If Not IsNothing(dtVino) AndAlso dtVino.Rows.Count > 0 Then
            If data.Nuevo Then
                dtVino.Rows(0)("QTotal") = 0
            Else
                ProcessServer.ExecuteTask(Of Guid)(AddressOf QTotObtener, data.IDVino, services)
            End If
        End If
        BusinessHelper.UpdateTable(dtVino)
    End Sub

    <Task()> Public Shared Sub ActualizarDiasVino(ByVal IDVino As Guid, ByVal services As ServiceProvider)
        'Se deben guardar en BdgVino los datos acumulados de los Dias (Deposito,Barrica,Botellero) de estancia del Vino.
        Dim clsVino As New BdgVino
        Dim Fecha As DateTime
        Dim dtVino As DataTable = clsVino.Filter("*", "IDVino='" & IDVino.ToString & "'")
        If Not IsNothing(dtVino) AndAlso dtVino.Rows.Count > 0 Then
            Fecha = dtVino.Rows(0)("Fecha")
        Else
            Fecha = Date.Today
        End If
        'CREATE  FUNCTION fBdgDiasDeposito
        '            @pIDVino uniqueidentifier,
        '            @pFechaSalida datetime,
        '            @pTipoDep int
        'RETURNS(Int)

        For TipoDepo As Integer = 0 To 3
            Select Case TipoDepo
                Case 0 'Deposito
                    Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
                    SqlCmd.CommandText = "fBdgDiasDeposito"
                    SqlCmd.CommandType = CommandType.StoredProcedure

                    Dim SqlParam1 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam1)
                    SqlParam1.ParameterName = "@pIDVino"
                    SqlParam1.Value = IDVino

                    Dim SqlParam2 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam2)
                    SqlParam2.ParameterName = "@pFechaSalida"
                    SqlParam2.Value = Fecha

                    Dim SqlParam3 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam3)
                    SqlParam3.ParameterName = "@pTipoDep"
                    SqlParam3.Value = TipoDepo

                    dtVino.Rows(0)("DiasDeposito") = Nz(AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteScalar), 0)
                Case 1 'Barrica
                    Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
                    SqlCmd.CommandText = "fBdgDiasDeposito"
                    SqlCmd.CommandType = CommandType.StoredProcedure

                    Dim SqlParam1 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam1)
                    SqlParam1.ParameterName = "@pIDVino"
                    SqlParam1.Value = IDVino

                    Dim SqlParam2 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam2)
                    SqlParam2.ParameterName = "@pFechaSalida"
                    SqlParam2.Value = Fecha

                    Dim SqlParam3 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam3)
                    SqlParam3.ParameterName = "@pTipoDep"
                    SqlParam3.Value = TipoDepo

                    dtVino.Rows(0)("DiasBarrica") = Nz(AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteScalar), 0)
                Case 2 'Botellero
                    Dim SqlCmd As Common.DbCommand = AdminData.GetCommand
                    SqlCmd.CommandText = "fBdgDiasDeposito"
                    SqlCmd.CommandType = CommandType.StoredProcedure

                    Dim SqlParam1 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam1)
                    SqlParam1.ParameterName = "@pIDVino"
                    SqlParam1.Value = IDVino

                    Dim SqlParam2 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam2)
                    SqlParam2.ParameterName = "@pFechaSalida"
                    SqlParam2.Value = Fecha

                    Dim SqlParam3 As Common.DbParameter = SqlCmd.CreateParameter
                    SqlCmd.Parameters.Add(SqlParam3)
                    SqlParam3.ParameterName = "@pTipoDep"
                    SqlParam3.Value = TipoDepo

                    dtVino.Rows(0)("DiasBotellero") = Nz(AdminData.Execute(SqlCmd, ExecuteCommand.ExecuteScalar), 0)
            End Select
        Next
        BusinessHelper.UpdateTable(dtVino)
    End Sub

    <Serializable()> _
    Public Class StRegistrarMovimiento
        Public DtVino As DataTable
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtVino As DataTable, ByVal Cantidad As Double)
            Me.DtVino = DtVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub RegistrarMovimiento(ByVal data As StRegistrarMovimiento, ByVal services As ServiceProvider)
        Dim Stk As arStockData = services.GetService(Of arStockData)()
        Dim StStockData As New StStockDataFromVino(data.DtVino, data.Cantidad)
        ReDim Preserve Stk.Stocks(Stk.Stocks.Length)
        Stk.Stocks(Stk.Stocks.Length - 1) = ProcessServer.ExecuteTask(Of StStockDataFromVino, StockData)(AddressOf StockDataFromVino, StStockData, services)
    End Sub

    <Serializable()> _
    Public Class StCheckCapOrigen
        Public DtDep As DataTable
        Public Cantidad As Double
        Public Origen As BdgOrigenVino

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtDep As DataTable, ByVal Cantidad As Double, ByVal Origen As BdgOrigenVino)
            Me.DtDep = DtDep
            Me.Cantidad = Cantidad
            Me.Origen = Origen
        End Sub
    End Class

    <Task()> Public Shared Sub CheckCapacityOrigen(ByVal data As StCheckCapOrigen, ByVal services As ServiceProvider)
        If CBool(data.DtDep.Rows(0)(_D.SinLimite)) Then
        Else
            'Comprobar Cantidad <= Capacidad - Ocupacion deposito Destino
            Dim strFieldName As String
            If data.Origen = BdgOrigenVino.Uva Then
                strFieldName = _D.CapacidadKg
            Else
                strFieldName = _D.Capacidad
            End If
            If data.DtDep.Rows(0)(_D.Ocupacion) + data.Cantidad > data.DtDep.Rows(0)(strFieldName) Then
                ApplicationService.GenerateError("No se puede realizar el movimiento porque se superaría la capacidad del depósito '|'.", data.DtDep.Rows(0)(_D.IDDeposito))
            End If
        End If
        If data.Cantidad < 0 Then
            If data.DtDep.Rows(0)(_D.Ocupacion) + data.Cantidad < 0 Then
                ApplicationService.GenerateError("No se puede realizar el movimiento. La cantidad a mover en el Depósito '|', excede la cantidad existente.", data.DtDep.Rows(0)(_D.IDDeposito))
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class StCheckCapVino
        Public DrDep As DataRow
        Public Cantidad As Double
        Public Origen As BdgOrigenVino
        Public DtVino As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal DrDep As DataRow, ByVal Cantidad As Double, ByVal Origen As BdgOrigenVino, ByVal DtVino As DataTable)
            Me.DrDep = DrDep
            Me.Cantidad = Cantidad
            Me.Origen = Origen
            Me.DtVino = DtVino
        End Sub
    End Class

    <Task()> Public Shared Sub CheckCapacityVino(ByVal data As StCheckCapVino, ByVal services As ServiceProvider)
        Dim StCheck As New StCheckCapVinoArticulo(data.DrDep.Table, data.Cantidad, data.Origen, data.DtVino.Rows(0)(_V.IDArticulo), data.DtVino.Rows(0)(_V.IDUdMedida))
        ProcessServer.ExecuteTask(Of StCheckCapVinoArticulo)(AddressOf CheckCapacityVinoArticulo, StCheck, services)
    End Sub

    <Serializable()> _
    Public Class StCheckCapVinoArticulo
        Public DtDep As DataTable
        Public Cantidad As Double
        Public Origen As BdgOrigenVino
        Public IDArticulo As String
        Public IDUDMedidaArticulo As String

        Public Sub New(ByVal DtDep As DataTable, ByVal Cantidad As Double, ByVal Origen As BdgOrigenVino, ByVal IDArticulo As String, ByVal IDUDMedidaArticulo As String)
            Me.DtDep = DtDep
            Me.Cantidad = Cantidad
            Me.Origen = Origen
            Me.IDArticulo = IDArticulo
            Me.IDUDMedidaArticulo = IDUDMedidaArticulo
        End Sub
    End Class

    <Task()> Public Shared Sub CheckCapacityVinoArticulo(ByVal data As StCheckCapVinoArticulo, ByVal services As ServiceProvider)
        'La Ocupación del depósito está en la unidad de medida del depósito.
        'La Cantidad está en la unidad de medida del artículo del vino.

        Dim StGetFactor As New StGetFactorConversion(data.IDArticulo, data.IDUDMedidaArticulo, data.DtDep.Rows(0)(_D.IDUDMedida) & String.Empty)
        Dim dbCantidad As Double = data.Cantidad * ProcessServer.ExecuteTask(Of StGetFactorConversion, Double)(AddressOf GetFactorConversion, StGetFactor, services)

        If Not CBool(data.DtDep.Rows(0)(_D.SinLimite)) Then
            'Comprobar Cantidad <= Capacidad - Ocupacion deposito Destino
            Dim strFieldName As String
            If data.Origen = BdgOrigenVino.Uva Then
                strFieldName = _D.CapacidadKg
            Else
                strFieldName = _D.Capacidad
            End If
            If data.DtDep.Rows(0)(_D.Ocupacion) + dbCantidad > data.DtDep.Rows(0)(strFieldName) Then
                ApplicationService.GenerateError("No se puede realizar el movimiento porque se superaría la capacidad del depósito '|'.", data.DtDep.Rows(0)(_D.IDDeposito))
            End If
        End If

        If dbCantidad < 0 Then
            If data.DtDep.Rows(0)(_D.Ocupacion) + dbCantidad < 0 Then
                ApplicationService.GenerateError("No se puede realizar el movimiento. La cantidad a mover en el Depósito '|' del Artículo '|', excede la cantidad existente.", data.DtDep.Rows(0)(_D.IDDeposito), data.IDArticulo)
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class StDeshacerVinoInt
        Public IDVino As Guid
        Public Origen As BdgOrigenVino
        Public Operacion As String
        Public NOperacion As String
        Public Cantidad As Double
        Public IDDeposito As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Origen As BdgOrigenVino, ByVal Operacion As String, ByVal NOperacion As String, ByVal Cantidad As Double, _
                       ByVal IDDeposito As String)
            Me.IDVino = IDVino
            Me.Origen = Origen
            Me.Operacion = Operacion
            Me.NOperacion = NOperacion
            Me.Cantidad = Cantidad
            Me.IDDeposito = IDDeposito
        End Sub
    End Class

    <Task()> Public Shared Sub DeshacerVinoInt(ByVal data As StDeshacerVinoInt, ByVal services As ServiceProvider)
        If New BdgDepositoVino().SelOnPrimaryKey(data.IDVino).Rows.Count > 0 Then
            '//esta en deposito
            Dim StEliminar As New StEliminarOpYActua(data.IDVino, data.Origen, data.Operacion, data.NOperacion, data.Cantidad, data.IDDeposito)
            ProcessServer.ExecuteTask(Of StEliminarOpYActua)(AddressOf EliminarOperacionDeEstructuraYActualizar, StEliminar, services)
        Else
            '//no esta en deposito
            Dim StEliminar As New StEliminarCompNivel(data.IDVino, data.Origen, data.Operacion, data.NOperacion, data.Cantidad, data.IDDeposito)
            ProcessServer.ExecuteTask(Of StEliminarCompNivel)(AddressOf EliminarComponenteDeEstructuraEnCualquierNivel, StEliminar, services)
        End If
    End Sub

    <Serializable()> _
    Public Class StDeshacerVino
        Public IDVino As Guid
        Public Operacion As String
        Public NOperacion As String
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Operacion As String, ByVal NOperacion As String, ByVal Cantidad As Double)
            Me.IDVino = IDVino
            Me.Operacion = Operacion
            Me.NOperacion = NOperacion
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Sub DeshacerVino(ByVal data As StDeshacerVino, ByVal services As ServiceProvider)
        Dim rwVino As DataRow = New BdgVino().GetItemRow(data.IDVino)

        Dim stError As String = String.Empty
        If rwVino(_V.Origen) = BdgOrigenVino.Uva Then
            stError = "El vino siguiente procede de Entradas de Uva. No se puede eliminar."
        ElseIf rwVino(_V.Origen) = BdgOrigenVino.Compra Then
            stError = "El vino siguiente procede de Compras de Vino. No se puede eliminar."
        End If
        If Length(stError) > 0 Then
            Dim StMsg As New StMensajeError(data.IDVino, stError)
            stError = String.Format(ProcessServer.ExecuteTask(Of StMensajeError, String)(AddressOf MensajeErrorVino, StMsg, services))
            ApplicationService.GenerateError(stError)
        End If

        Dim rwDep As DataRow = New BdgDeposito().GetItemRow(rwVino(_V.IDDeposito))

        Dim StCheck As New StCheckCapVino(rwDep, -data.Cantidad, BdgOrigenVino.Interno, rwVino.Table)
        ProcessServer.ExecuteTask(Of StCheckCapVino)(AddressOf CheckCapacityVino, StCheck, services)

        Dim StDeshacer As New StDeshacerVinoInt(data.IDVino, rwVino(_V.Origen), data.Operacion, data.NOperacion, data.Cantidad, rwVino(_V.IDDeposito))
        ProcessServer.ExecuteTask(Of StDeshacerVinoInt)(AddressOf DeshacerVinoInt, StDeshacer, services)

        Dim StGetFactor As New StGetFactorConversion(rwVino(_V.IDArticulo), rwVino(_V.IDUdMedida) & String.Empty, rwDep(_D.IDUDMedida) & String.Empty)
        rwDep(_D.Ocupacion) += -data.Cantidad * ProcessServer.ExecuteTask(Of StGetFactorConversion, Double)(AddressOf GetFactorConversion, StGetFactor, services)
        BusinessHelper.UpdateTable(rwDep.Table)
    End Sub

    <Serializable()> _
    Public Class StGetFactorConversion
        Public IDArticulo As String
        Public UDOrigen As String
        Public UDDestino As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal UDOrigen As String, ByVal UDDestino As String)
            Me.IDArticulo = IDArticulo
            Me.UDOrigen = UDOrigen
            Me.UDDestino = UDDestino
        End Sub
    End Class

    <Task()> Public Shared Function GetFactorConversion(ByVal data As StGetFactorConversion, ByVal services As ServiceProvider) As Double
        If Length(data.UDOrigen) > 0 And Length(data.UDDestino) > 0 Then
            Dim dataAB As New ArticuloUnidadAB.DatosFactorConversion(data.IDArticulo, data.UDOrigen, data.UDDestino)
            Return ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, dataAB, services)
        Else
            Return 1
        End If
    End Function

    <Serializable()> _
    Public Class StEliminarOp
        Public DtVinoE As DataTable
        Public Origen As BdgOrigenVino
        Public Operacion As String
        Public NOperacion As String
        Public Cantidad As Double
        Public IDVino As Guid

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtVinoE As DataTable, ByVal Origen As BdgOrigenVino, ByVal Operacion As String, ByVal NOperacion As String, _
                       ByVal Cantidad As Double, ByVal IDVino As Guid)
            Me.DtVinoE = DtVinoE
            Me.Origen = Origen
            Me.Operacion = Operacion
            Me.NOperacion = NOperacion
            Me.Cantidad = Cantidad
            Me.IDVino = IDVino
        End Sub
    End Class

    <Task()> Public Shared Function EliminarOperacionDeEstructura(ByVal data As StEliminarOp, ByVal services As ServiceProvider) As VinoQ
        Dim rslt As VinoQ
        Dim blHayMezcla As Boolean

        Dim strFilter As String = String.Empty
        If data.Origen = BdgOrigenVino.AlbaranTransferencia Then
            strFilter = "Operacion = '" & data.Operacion & "'"
        Else
            strFilter = "NOperacion = '" & data.NOperacion & "'"
        End If

        'Si hay mezclas se devuelven los datos del vino mezclado para volver a darlo de alta.
        For Each rwVinoE As DataRow In data.DtVinoE.Select(strFilter)
            'Los vinos componentes de la estructura que no estén el Origen de la Operación son Mezclas
            Dim StSelOn As New BdgOperacionVino.StSelOnNOperacionVinoOrigen(data.NOperacion, rwVinoE(_VE.IDVinoComponente))
            Dim oOV As New BdgOperacionVino
            Dim dtOV As DataTable = ProcessServer.ExecuteTask(Of BdgOperacionVino.StSelOnNOperacionVinoOrigen, DataTable)(AddressOf BdgOperacionVino.SelOnNOperacionVinoOrigen, StSelOn, services)
            If dtOV.Rows.Count = 0 And data.Origen = BdgOrigenVino.Interno Then
                blHayMezcla = True
            Else
                blHayMezcla = False
            End If

            If blHayMezcla Then
                Dim StBuscar As New StBuscarOtraOperacion(data.Operacion, data.DtVinoE)
                Dim StrBuscar As String = ProcessServer.ExecuteTask(Of StBuscarOtraOperacion, String)(AddressOf BuscarOtraOperacion, StBuscar, services)
                If Len(StrBuscar) > 0 Then
                    rwVinoE(_VE.Operacion) = StrBuscar
                    rwVinoE(_VE.NOperacion) = rwVinoE(_VE.Operacion)
                Else
                    rslt = New VinoQ
                    rslt.IDVino = rwVinoE(_VE.IDVinoComponente)
                    rslt.Cantidad = rwVinoE(_VE.Cantidad)

                    rwVinoE.Delete()
                End If
            Else
                rwVinoE.Delete()
            End If
        Next
        AdminData.SetData(data.DtVinoE)
        Return rslt
    End Function

    <Serializable()> _
    Public Class StEliminarOpYActua
        Public IDVino As Guid
        Public Origen As BdgOrigenVino
        Public Operacion As String
        Public NOperacion As String
        Public Cantidad As Double
        Public IDDeposito As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Origen As BdgOrigenVino, ByVal Operacion As String, ByVal NOperacion As String, ByVal Cantidad As Double, _
                       ByVal IDDeposito As String)
            Me.IDVino = IDVino
            Me.Origen = Origen
            Me.Operacion = Operacion
            Me.NOperacion = NOperacion
            Me.Cantidad = Cantidad
            Me.IDDeposito = IDDeposito
        End Sub
    End Class

    'TODO puede sobrar (EliminarComponenteDeEstructuraEnCualquierNivel)
    <Task()> Public Shared Sub EliminarOperacionDeEstructuraYActualizar(ByVal data As StEliminarOpYActua, ByVal services As ServiceProvider)
        Dim dtVinoE As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgVinoEstructura.SelOnVino, data.IDVino, services)

        Dim StBuscar As New StBuscarOtraOperacion(data.Operacion, dtVinoE)
        If Len(ProcessServer.ExecuteTask(Of StBuscarOtraOperacion, String)(AddressOf BuscarOtraOperacion, StBuscar, services)) = 0 Then
            '//si el vino sólo se ha compuesto con una unica operación y tiene salidas, no se puede eliminar
            Dim stOperaciones As String
            Dim StVino As New StVinoContenidoTienesalidas(data.IDVino, stOperaciones)
            If ProcessServer.ExecuteTask(Of StVinoContenidoTienesalidas, Boolean)(AddressOf VinoContenidoTieneSalidas, StVino, services) Then
                Dim stError As String = "El vino siguiente tiene salidas. No se puede eliminar."
                Dim StMensaje As New StMensajeError(data.IDVino, stError)
                stError = String.Format(ProcessServer.ExecuteTask(Of StMensajeError, String)(AddressOf MensajeErrorVino, StMensaje, services) & vbCrLf & "Operaciones: {0}", StVino.StOperaciones)
                ApplicationService.GenerateError(stError)
            End If
        End If

        Dim StEliminar As New StEliminarOp(dtVinoE, data.Origen, data.Operacion, data.NOperacion, data.Cantidad, data.IDVino)
        Dim OtroVino As VinoQ = ProcessServer.ExecuteTask(Of StEliminarOp, VinoQ)(AddressOf EliminarOperacionDeEstructura, StEliminar, services)
        If Not OtroVino Is Nothing Then
            Dim StMeter As New StMeterVinoEnDeposito(OtroVino.IDVino, data.IDDeposito, OtroVino.Cantidad)
            ProcessServer.ExecuteTask(Of StMeterVinoEnDeposito)(AddressOf MeterVinoEnDeposito, StMeter, services)
        End If

        If dtVinoE.Rows.Count = 0 Then
            Dim dttOtrasOper As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgOperacionVino.SelOnVino, data.IDVino, services)
            If (Not dttOtrasOper Is Nothing AndAlso dttOtrasOper.Rows.Count > 0) Then
                Dim StSumar As New StSumarCantidadAContenido(data.IDVino, -data.Cantidad)
                ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)
            Else
                ProcessServer.ExecuteTask(Of Guid)(AddressOf EliminarVinoDeDeposito, data.IDVino, services)
                ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgVino.DeleteVino, data.IDVino, services)
            End If
        Else
            Dim StSumar As New StSumarCantidadAContenido(data.IDVino, -data.Cantidad)
            ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)
        End If
    End Sub

    <Serializable()> _
    Public Class StModificarCantidadCompEnNivel
        Public IDVino As Guid
        Public Cantidad As Double
        Public IDDeposito As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Cantidad As Double, ByVal IDDeposito As String)
            Me.IDVino = IDVino
            Me.Cantidad = Cantidad
            Me.IDDeposito = IDDeposito
        End Sub
    End Class

    <Task()> Public Shared Function ModificarCantidadComponenteEnCualquierNivel(ByVal data As StModificarCantidadCompEnNivel, ByVal services As ServiceProvider) As Guid
        '//solo es modificable un vino a cualquier nivel si no tiene bifurcaciones
        '//y todos sus descendientes han permanecido en el mismo deposito
        '//O sea: es componente de un vino porque se le han ido añadiendo vinos desde otros depositos

        '//implosion del vino
        Dim IDVinoPadre As Guid = data.IDVino
        Do
            Dim dtVinoEImplosion As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgVinoEstructura.SelOnComponente, IDVinoPadre, services)
            If dtVinoEImplosion.Rows.Count = 0 Then Exit Do
            If dtVinoEImplosion.Rows.Count > 1 Then
                Dim stOperaciones As String
                For Each Dr As DataRow In dtVinoEImplosion.Rows
                    If Len(stOperaciones) = 0 Then
                        stOperaciones = Dr(_VE.NOperacion)
                    Else
                        stOperaciones = stOperaciones & ", " & Dr(_VE.NOperacion)
                    End If
                Next

                Dim stError As String = "El vino siguiente tiene salidas. No se puede eliminar."
                Dim StMensaje As New StMensajeError(IDVinoPadre, stError)
                stError = String.Format(ProcessServer.ExecuteTask(Of StMensajeError, String)(AddressOf MensajeErrorVino, StMensaje, services) & vbCrLf & "Operaciones: {0}", stOperaciones)
                ApplicationService.GenerateError(stError)
            End If
            '//dtVinoEImplosion.Rows.Count = 1
            IDVinoPadre = dtVinoEImplosion.Rows(0)(_VE.IDVino)
            If New BdgVino().GetItemRow(IDVinoPadre)(_V.IDDeposito) <> data.IDDeposito Then
                Dim stError As String = "El vino ya no se encuentra en el depósito. No se puede eliminar."
                Dim StMensaje As New StMensajeError(IDVinoPadre, stError)
                stError = String.Format(ProcessServer.ExecuteTask(Of StMensajeError, String)(AddressOf MensajeErrorVino, StMensaje, services))
                ApplicationService.GenerateError(stError)
            End If

            Dim rwVinoE As DataRow = dtVinoEImplosion.Rows(0)

            rwVinoE(_VE.Cantidad) += data.Cantidad
            If rwVinoE(_VE.Cantidad) < 0 Then ApplicationService.GenerateError("No se puede realizar la modificación. La cantidad en el depósito es inferior a la requerida.")
            If rwVinoE(_VE.Cantidad) = 0 Then rwVinoE.Delete()

            BusinessHelper.UpdateTable(dtVinoEImplosion)
        Loop
        Return IDVinoPadre
    End Function

    <Serializable()> _
    Public Class StEliminarCompNivel
        Public IDVino As Guid
        Public Origen As BdgOrigenVino
        Public Operacion As String
        Public NOperacion As String
        Public Cantidad As Double
        Public IDDeposito As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Origen As BdgOrigenVino, ByVal Operacion As String, ByVal NOperacion As String, ByVal Cantidad As Double, _
                       ByVal IDDeposito As String)
            Me.IDVino = IDVino
            Me.Origen = Origen
            Me.Operacion = Operacion
            Me.NOperacion = NOperacion
            Me.Cantidad = Cantidad
            Me.IDDeposito = IDDeposito
        End Sub
    End Class

    <Task()> Public Shared Sub EliminarComponenteDeEstructuraEnCualquierNivel(ByVal data As StEliminarCompNivel, ByVal services As ServiceProvider)
        Dim StModificar As New StModificarCantidadCompEnNivel(data.IDVino, -data.Cantidad, data.IDDeposito)
        Dim IDVinoEnDeposito As Guid = ProcessServer.ExecuteTask(Of StModificarCantidadCompEnNivel, Guid)(AddressOf ModificarCantidadComponenteEnCualquierNivel, StModificar, services)
        Dim dtVinoE As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgVinoEstructura.SelOnVino, data.IDVino, services)
        Dim StEliminar As New StEliminarOp(dtVinoE, data.Origen, data.Operacion, data.NOperacion, data.Cantidad, data.IDVino)
        Dim OtroVino As VinoQ = ProcessServer.ExecuteTask(Of StEliminarOp, VinoQ)(AddressOf EliminarOperacionDeEstructura, StEliminar, services)
        If Not OtroVino Is Nothing Then
            Dim dtVinoEImplosion As DataTable = ProcessServer.ExecuteTask(Of Guid, DataTable)(AddressOf BdgVinoEstructura.SelOnComponente, data.IDVino, services)
            If dtVinoEImplosion.Rows.Count > 0 Then
                dtVinoEImplosion.Rows(0)(_VE.IDVinoComponente) = OtroVino.IDVino
                BusinessHelper.UpdateTable(dtVinoEImplosion)
            Else
                ApplicationService.GenerateError("La operación no se puede borrar.")
            End If
            ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgVino.DeleteVino, data.IDVino, services)
        End If
        If dtVinoE.Rows.Count = 0 Then ProcessServer.ExecuteTask(Of Guid)(AddressOf BdgVino.DeleteVino, data.IDVino, services)
        Dim StSumar As New StSumarCantidadAContenido(IDVinoEnDeposito, -data.Cantidad)
        ProcessServer.ExecuteTask(Of StSumarCantidadAContenido)(AddressOf SumarCantidadAContenido, StSumar, services)
    End Sub

    <Task()> Public Shared Sub EliminarVinoDeDeposito(ByVal IDVino As Guid, ByVal services As ServiceProvider)
        Dim rwDepV As DataRow = New BdgDepositoVino().GetItemRow(IDVino)
        Dim StCambio As New StCambiarOcupYStock(rwDepV.Table, -rwDepV(_DV.Cantidad), New BdgVino().GetItemRow(rwDepV(_DV.IDVino)).Table)
        ProcessServer.ExecuteTask(Of StCambiarOcupYStock)(AddressOf CambiarOcupacionYStock, StCambio, services)
    End Sub

    <Serializable()> _
    Public Class StBuscarOtraOperacion
        Public Operacion As String
        Public DtVinoE As DataTable

        Public Sub New()
        End Sub

        Public Sub New(ByVal Operacion As String, ByVal DtVinoE As DataTable)
            Me.Operacion = Operacion
            Me.DtVinoE = DtVinoE
        End Sub
    End Class

    <Task()> Public Shared Function BuscarOtraOperacion(ByVal data As StBuscarOtraOperacion, ByVal services As ServiceProvider) As String
        Dim rslt As String
        For Each oRw As DataRow In data.DtVinoE.Rows
            If oRw.RowState = DataRowState.Unchanged Then
                If oRw(_VE.Operacion) <> data.Operacion Then
                    If Len(rslt) = 0 OrElse rslt > oRw(_VE.Operacion) Then rslt = oRw(_VE.Operacion)
                End If
            End If
        Next
        Return rslt
    End Function

    <Serializable()> _
    Public Class StEjecutarMovimientos
        Public Documento As String
        Public FechaDocumento As Date
        Public Precio As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal Documento As String, ByVal FechaDocumento As Date, Optional ByVal Precio As Double = 0)
            Me.Documento = Documento
            Me.FechaDocumento = FechaDocumento
            Me.Precio = Precio
        End Sub
    End Class

    <Task()> Public Shared Function EjecutarMovimientos(ByVal data As StEjecutarMovimientos, ByVal services As ServiceProvider) As Integer
        'TODO no esta resuelto el tema de los precios de los movimientos
        Dim Stk As arStockData = services.GetService(Of arStockData)()
        If Stk.Stocks.Length = 0 Then Exit Function
        Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
        Dim StEj As New StEjecutarMovimientosNumero(NumeroMovimiento, data.Documento, data.FechaDocumento, data.Precio)
        ProcessServer.ExecuteTask(Of StEjecutarMovimientosNumero)(AddressOf EjecutarMovimientosNumero, StEj, services)
        Return NumeroMovimiento
    End Function

    <Serializable()> _
    Public Class StEjecutarMovimientosNumero
        Public NumeroMovimiento As Integer
        Public Documento As String
        Public FechaDocumento As Date
        Public Precio As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal NumeroMovimiento As Integer, ByVal Documento As String, ByVal FechaDocumento As Date, Optional ByVal Precio As Double = 0)
            Me.NumeroMovimiento = NumeroMovimiento
            Me.Documento = Documento
            Me.FechaDocumento = FechaDocumento
            Me.Precio = Precio
        End Sub
    End Class

    <Task()> Public Shared Sub EjecutarMovimientosNumero(ByVal data As StEjecutarMovimientosNumero, ByVal services As ServiceProvider)
        Dim Stk As arStockData = services.GetService(Of arStockData)()
        If Stk.Stocks.Length > 0 Then
            Dim strError As String = String.Empty
            Dim oMI As MonedaInfo = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaB, data.FechaDocumento, services)
            For Each oSd As StockData In Stk.Stocks
                oSd.Documento = data.Documento
                oSd.FechaDocumento = data.FechaDocumento.Date
                oSd.PrecioA = data.Precio
                oSd.PrecioB = xRound(data.Precio * oMI.CambioB, oMI.NDecimalesPrecio)
            Next
            Dim dataMovGenericoES As New ProcesoStocks.DataMovimientosGenericosES(data.NumeroMovimiento, Stk.Stocks, False)
            Dim SUD As StockUpdateData() = ProcessServer.ExecuteTask(Of ProcesoStocks.DataMovimientosGenericosES, StockUpdateData())(AddressOf ProcesoStocks.MovimientosGenericosES, dataMovGenericoES, services)
            For Each oSUD As StockUpdateData In SUD
                If oSUD.Estado = EstadoStock.NoActualizado Then
                    strError &= oSUD.Log
                End If
            Next
            ReDim Stk.Stocks(-1)
            If Len(strError) Then Throw New Exception(strError)
        End If
    End Sub

    <Serializable()> _
    Public Class StStockDataFromVino
        Public DtVino As DataTable
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal DtVino As DataTable, ByVal Cantidad As Double)
            Me.DtVino = DtVino
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function StockDataFromVino(ByVal data As StStockDataFromVino, ByVal services As ServiceProvider) As StockData
        Dim oSD As New StockData
        oSD.Articulo = data.DtVino.Rows(0)(_V.IDArticulo)
        oSD.Almacen = data.DtVino.Rows(0)(_V.IDAlmacen)
        If data.Cantidad > 0 Then
            oSD.TipoMovimiento = enumTipoMovimiento.tmEntFabrica
        Else
            oSD.TipoMovimiento = enumTipoMovimiento.tmSalFabrica
            data.Cantidad = -data.Cantidad
        End If
        oSD.Cantidad = data.Cantidad
        oSD.Traza = data.DtVino.Rows(0)(_V.IDVino)
        oSD.Lote = data.DtVino.Rows(0)(_V.Lote)
        oSD.Ubicacion = data.DtVino.Rows(0)(_V.IDDeposito)
        oSD.FechaDocumento = Date.Today
        Return oSD
    End Function

    <Serializable()> _
    Public Class StDeshaceEstructura
        Public IDVino As Guid
        Public NOperacion As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal NOperacion As String)
            Me.IDVino = IDVino
            Me.NOperacion = NOperacion
        End Sub
    End Class

    <Task()> Public Shared Sub DeshacerEstructura(ByVal data As StDeshaceEstructura, ByVal services As ServiceProvider)
        Dim StSel As New BdgVinoEstructura.StSelOnVinoOperacion(data.IDVino, data.NOperacion)
        Dim dtVE As DataTable = ProcessServer.ExecuteTask(Of BdgVinoEstructura.StSelOnVinoOperacion, DataTable)(AddressOf BdgVinoEstructura.SelOnVinoOperacion, StSel, services)
        For Each oRw As DataRow In dtVE.Rows
            oRw.Delete()
        Next
        BusinessHelper.UpdateTable(dtVE)
    End Sub

    <Serializable()> _
    Public Class VinoQ
        Public IDVino As Guid
        Public Cantidad As Double
    End Class

    <Serializable()> _
    Public Class StMensajeError
        Public IDVino As Guid
        Public Mensaje As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDVino As Guid, ByVal Mensaje As String)
            Me.IDVino = IDVino
            Me.Mensaje = Mensaje
        End Sub
    End Class

    <Task()> Public Shared Function MensajeErrorVino(ByVal data As StMensajeError, ByVal services As ServiceProvider) As String
        Dim stMensajeRdo As String
        Dim dt As DataTable = New BdgVino().SelOnPrimaryKey(data.IDVino)
        If dt.Rows.Count > 0 Then
            stMensajeRdo = String.Format(data.Mensaje & vbCrLf & vbCrLf _
            & "Depósito: {0}" & vbCrLf _
            & "Artículo: {1}" & vbCrLf _
            & "Lote: {2}" & vbCrLf _
            & "Almacén: {3}" & vbCrLf _
            & "Fecha Llenado: {4}" & vbCrLf, _
            dt.Rows(0)(_V.IDDeposito), dt.Rows(0)(_V.IDArticulo), dt.Rows(0)(_V.Lote), dt.Rows(0)(_V.IDAlmacen), dt.Rows(0)(_V.Fecha))
        Else : stMensajeRdo = data.Mensaje
        End If
        Return stMensajeRdo
    End Function

    <Task()> Public Shared Function GetIDVinoCantidad(ByVal data As ProcesoStocks.DataVinoQ, ByVal services As ServiceProvider) As ProcesoStocks.DataVinoQ
        Dim oFltr As New Filter
        oFltr.Add(New StringFilterItem(_V.IDArticulo, data.IDArticulo))
        oFltr.Add(New StringFilterItem(_V.IDDeposito, data.IDDeposito))
        oFltr.Add(New StringFilterItem(_V.Lote, data.Lote))
        oFltr.Add(New StringFilterItem(_V.IDAlmacen, data.IDAlmacen))

        Dim dtVinos As DataTable = AdminData.GetData("negBdgVinoDesdeStock", oFltr)
        If dtVinos.Rows.Count > 1 Then ApplicationService.GenerateError("Situación inesperada: hay mas de un vino con el mismo Artículo/Depósito/Lote/Almacén")
        If dtVinos.Rows.Count = 1 Then
            data.IDVino = dtVinos.Rows(0)(_V.IDVino)
            data.Cantidad = dtVinos.Rows(0)(_DV.Cantidad)
            Return data
        End If
    End Function

End Class


<Serializable()> _
Public Class arStockData
    Public Stocks(-1) As StockData

    Public Sub New()
    End Sub
End Class

'<Serializable()> _
'Public Class clsVinoQ
'    Public IDArticulo As String
'    Public IDDeposito As String
'    Public Lote As String
'    Public IDAlmacen As String

'    Public IDVino As Guid
'    Public Cantidad As Double

'    Public Sub New()
'    End Sub

'    Public Sub New(ByVal IDArticulo As String, ByVal IDDeposito As String, ByVal Lote As String, ByVal IDAlmacen As String)
'        Me.IDArticulo = IDArticulo
'        Me.IDDeposito = IDDeposito
'        Me.Lote = Lote
'        Me.IDAlmacen = IDAlmacen
'    End Sub

'End Class
