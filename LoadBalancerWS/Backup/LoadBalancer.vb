'//////////////////////////////////////////////////////////////////////////////
'//                                                                          //
'// LOAD BALANCER - WEBSERVICE CONNECTION   V.1                              //
'//                                                                          //
'// IVAN SALDIVAR RODRIGUEZ (R)2011                                          //
'//                                                                          //
'// TODOS LOS DERECHOS DE COPIA ESTÁN RESERVADOS A IVAN SALDIVAR RODRIGUEZ,  //
'// SU USO NO AUTORIZADO SERÁ SANCIONADO DE ACUERDO A LAS LEYES DE DERECHO   //
'// DE AUTOR VIGENTES EN EL PAÍS DONDE SE HAYA COMETIDO LA FALTA.            //
'// SE PROHIBE LA COPIA, REPRODUCCIÓN Y DISTRIBUCIÓN CON FINES COMERCIALES   //
'// SIN PREVIA AUTORIZACIÓN. EL ALGORITMO DE BALANCEO ES PROPIEDAD DEL AUTOR //
'// ,POR LO TANTO, ESTE FORMA PARTE INTEGRA DEL CÓDIGO EN EL CUAL SE HA      //
'// IMPLEMENTADO.                                                            //
'//                                                                          //
'// SANTIAGO - CHILE, 2011-01-01                                             //
'//////////////////////////////////////////////////////////////////////////////

Public Class LoadBalancerWS
    Dim strTitulo As String = ""

    Public Sub AsignaWEBSERVICE(ByVal userid As String, _
                                 ByRef classWS As Object, _
                                 ByRef instanciaWS As String, _
                                 ByRef registro_sistema As Boolean, _
                                 ByVal strTitulo As String)

        'BALANCEADOR DE CARGA EN WEBSERVICES DEFINIDOS PARA LA INSTALACIÓN.
        'APLICAR LA SIGUIENTE FORMULA PARA EL CÁLCULO DE LA CANTIDAD DE INSTANCIAS A
        'CREAR PARA EL BALANCEO DINÁMICO:
        '
        'k = CANTIDAD MÁXIMA DE USUARIOS (PICK)
        'N = CANTIDAD DE SERVICIOS WEB A INSTALAR (CLONES)
        '
        '              N = k / 100
        '
        'LO QUE ASEGURA QUE EN EL PICK DE UTILIZACIÓN NO HABRÁ MÁS DE 100 USUARIOS CONECTADOS
        'A UN MISMO SERVICIO WEB.
        '
        Dim Parametros As String = ""
        Dim DTS As Data.DataSet = Nothing
        Dim WServiceName As String = ""
        Dim WServiceNameInstancia As String = ""

        Try
            Parametros = "<PARAMS>" & _
                         "<PARAM nombre='@ID_USER' valor='" & userid & "'/>" & _
                         "</PARAMS>"

            DTS = Deserialize(classWS.CONSULTA_GENERICA_SP_SERIALIZADA(Serialize("LBWS_CONSULTA_CONEXIONES_WS", True), Serialize(Parametros, True)), "DATASET")

            If CONSULTA_ERROR(DTS, strTitulo) Then
                Exit Sub
            End If


            If DTS IsNot Nothing Then
                If DTS.Tables.Count > 0 Then
                    If DTS.Tables(0).Rows.Count > 0 Then
                        WServiceName = DTS.Tables(0).Rows(0).Item("NWEBSERVICE").ToString.Trim

                        classWS.Url = WServiceName

                        If DTS.Tables(0).Rows(0).Item("NWEBSERVICE_NOASIGNADO").ToString.Trim <> "" Then
                            instanciaWS = "INSTANCIA CONEXIÓN: " & DTS.Tables(0).Rows(0).Item("NWEBSERVICE_NOASIGNADO").ToString.Trim & " | "

                        ElseIf DTS.Tables(0).Rows(0).Item("NWEBSERVICE_MENOS_CARGA").ToString.Trim <> "" Then
                            instanciaWS = "Instancia conexión: " & DTS.Tables(0).Rows(0).Item("NWEBSERVICE_MENOS_CARGA").ToString.Trim & " | "

                        End If
                        'ISR_20120808
                        registro_sistema = True

                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox("Atención, intentó conectarse a la instancia de sesión y se produjo una excepción: " & Chr(13) & Chr(13) & ex.Message & Chr(13) & Chr(13) & "Por favor. contactar al Administrador.", MsgBoxStyle.Exclamation, strTitulo)

        End Try

    End Sub
    Public Sub DesAsignaWEBSERVICE(ByVal userid As String, _
                                    ByRef classWS As Object)
        'BALANCEADOR DE CARGA EN WEBSERVICES DEFINIDOS PARA LA INSTALACIÓN.
        Dim Parametros As String = ""
        Dim salida As String = ""

        Try
            Parametros = "<PARAMS>" & _
                         "<PARAM nombre='@ID_USER' valor='" & userid & "'/>" & _
                         "</PARAMS>"

            salida = Deserialize(classWS.CONSULTA_GENERICA_SP2_SERIALIZADA(Serialize("LBWS_ELIMINA_CONEXION_WS", True), Serialize(Parametros, True)), "STRING")

            If salida IsNot Nothing Then
                If InStr(salida, "OK") = 0 Then
                    MsgBox("Atención, al salir del sistema se intentó eliminar la referencia de conexión a la instancia", MsgBoxStyle.Information, strTitulo)
                End If
            End If

        Catch ex As Exception
            MsgBox("Atención, al salir del sistema se intentó eliminar la referencia de conexión a la instancia", MsgBoxStyle.Information, strTitulo)
        End Try

    End Sub
    Private Function CONSULTA_ERROR(ByVal DTS As DataSet, ByVal strtitulo As String) As Boolean
        Dim salida As Boolean

        salida = False
        If DTS IsNot Nothing Then
            If DTS.Tables.Count > 0 Then
                If DTS.Tables(0).Columns.Item(0).ColumnName = "ERROR_PROCEDURE_" Then


                    MsgBox("Se presentó el siguiente error:" & Chr(13) & Chr(13) & _
                           "PROCEDIMIENTO:" & DTS.Tables(0).Rows(0).Item(0).ToString & Chr(13) & _
                           "CÓDIGO DE ERROR:" & DTS.Tables(0).Rows(0).Item(1).ToString & Chr(13) & Chr(13) & _
                           "DESCRIPCIÓN:" & Chr(13) & _
                           DTS.Tables(0).Rows(0).Item(2).ToString & Chr(13) & _
                           "LÍNEA DE ERROR:" & DTS.Tables(0).Rows(0).Item(3).ToString & Chr(13), MsgBoxStyle.Critical, strtitulo)

                    salida = True

                End If

            End If
        End If

        Return salida

    End Function
    Private Function Deserialize(ByVal Obj As Byte(), Optional ByVal TIPO As String = "") As Object
        Dim objDeserialize As Object

        If Obj IsNot Nothing Then
            Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
            Dim ms As New IO.MemoryStream(Obj)
            objDeserialize = bf.Deserialize(ms)

            If objDeserialize.GetType.Name.Trim.ToUpper = TIPO Or TIPO = "" Then
                Return objDeserialize
            Else
                If objDeserialize.GetType.Name.Trim.ToUpper = "STRING" Then
                    MsgBox("Ha sucedido una excepción : " & Chr(13) & Chr(13) & objDeserialize.ToString, MsgBoxStyle.Exclamation, "")
                End If

                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Private Function Serialize(ByVal Obj As Object, ByVal AsByte As Boolean) As Byte()
        Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim ms As New IO.MemoryStream
        bf.Serialize(ms, Obj)
        Return ms.ToArray
    End Function
End Class
