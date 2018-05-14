'///////////////////////////////////////////////////////////////////////
'//                                                                   //
'// LOAD BALANCER - WEBSERVICE CONNECTION   V.1                       //
'//                                                                   //
'// IVAN SALDIVAR RODRIGUEZ (R)2011                                   //
'//                                                                   //
'//                                                                   //
'// SU USO ESTÁ AUTORIZADO BAJO LICENCIA                              //                                                          
'// Attribution 4.0 International (CC BY 4.0)                         //
'//                                                                   //
'// SANTIAGO - CHILE, 2012-01-01                                      //       
'//                                                                   //
'///////////////////////////////////////////////////////////////////////

Imports Microsoft.Win32

Public Class LoadBalancerWS
    Dim strTitulo As String = ""

    Public Function AsignaWEBSERVICE(ByVal Parametros As String,
                                 ByRef classWS As Object,
                                 ByRef instanciaWS As String,
                                 ByRef registro_sistema As Boolean,
                                 ByRef strURI_Balanceo As String,
                                 ByVal vConnectionString As String,
                                 ByRef statusLic As Boolean
                                 ) As String

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
        Dim DTS As Data.DataSet = Nothing
        Dim WServiceName As String = ""
        Dim WServiceNameInstancia As String = ""
        Dim estadoLicencia As Boolean
        Dim datecurrent As DateTime
        Dim outfunction As String = ""

        datecurrent = Now()
        estadoLicencia = GetLicense()

        statusLic = estadoLicencia

        Try

            DTS = Deserialize(classWS.CONSULTA_GENERICA_SP_SERIALIZADA(Serialize("LBWS_CONSULTA_CONEXIONES_WS", True), Serialize(Parametros, True), vConnectionString), "DATASET")

            If CONSULTA_ERROR(DTS, strTitulo) Then
                outfunction = strTitulo

            Else
                If estadoLicencia Then
                    If DTS IsNot Nothing Then
                        If DTS.Tables.Count > 0 Then
                            If DTS.Tables(0).Rows.Count > 0 Then
                                WServiceName = DTS.Tables(0).Rows(0).Item("NWEBSERVICE").ToString.Trim

                                'classWS.Url = WServiceName
                                strURI_Balanceo = WServiceName

                                If DTS.Tables(0).Rows(0).Item("NWEBSERVICE_NOASIGNADO").ToString.Trim <> "" Then
                                    instanciaWS = DTS.Tables(0).Rows(0).Item("NWEBSERVICE_NOASIGNADO").ToString.Trim

                                ElseIf DTS.Tables(0).Rows(0).Item("NWEBSERVICE_MENOS_CARGA").ToString.Trim <> "" Then
                                    instanciaWS = DTS.Tables(0).Rows(0).Item("NWEBSERVICE_MENOS_CARGA").ToString.Trim

                                End If
                                'ISR_20120808
                                registro_sistema = True

                            End If
                        End If
                    End If

                Else
                    'classWS.Url = ""
                    instanciaWS = "NO SE HA BALANCEADO LA CONEXION: EL COMPONENTE DE BALANCEO NO ESTA REGISTRADO O ESTA USANDO UNA COPIA NO AUTORIZADA. CONTACTE AL PROVEEDOR&NewLine;IVAN SALDIVAR RODRIGUEZ (R)(c) - " & datecurrent.Year.ToString() & "&NewLine;ivansaldivar@gmail.com"
                    registro_sistema = False
                    strURI_Balanceo = ""
                End If
            End If

        Catch ex As Exception
            outfunction = "Atención, intentó conectarse a la instancia de sesión y se produjo una excepción: " & Chr(13) & Chr(13) & ex.Message & Chr(13) & Chr(13) & "Por favor. contactar al Administrador."

        End Try

        AsignaWEBSERVICE = outfunction

    End Function

    Public Function DesAsignaWEBSERVICE(ByVal Parametros As String,
                                    ByRef classWS As Object,
                                    ByVal vConnectionString As String) As String
        'BALANCEADOR DE CARGA EN WEBSERVICES DEFINIDOS PARA LA INSTALACIÓN.
        Dim salida As String = ""

        Try

            salida = Deserialize(classWS.CONSULTA_GENERICA_SP2_SERIALIZADA(Serialize("LBWS_ELIMINA_CONEXION_WS", True), Serialize(Parametros, True), vConnectionString), "STRING")

            If salida IsNot Nothing Then
                If InStr(salida, "OK") = 0 Then
                    salida = "Atención, al salir del sistema se intentó eliminar la referencia de conexión a la instancia"
                End If
            End If

        Catch ex As Exception
            salida = "Atención, al salir del sistema se intentó eliminar la referencia de conexión a la instancia"
        End Try
        DesAsignaWEBSERVICE = salida

    End Function
    Private Function CONSULTA_ERROR(ByVal DTS As DataSet, ByRef strError As String) As Boolean
        Dim salida As Boolean

        salida = False
        If DTS IsNot Nothing Then
            If DTS.Tables.Count > 0 Then
                If DTS.Tables(0).Columns.Item(0).ColumnName = "ERROR_PROCEDURE_" Then


                    strError = "Se presentó el siguiente error:" & Chr(13) & Chr(13) &
                           "PROCEDIMIENTO:" & DTS.Tables(0).Rows(0).Item(0).ToString & Chr(13) &
                           "CÓDIGO DE ERROR:" & DTS.Tables(0).Rows(0).Item(1).ToString & Chr(13) & Chr(13) &
                           "DESCRIPCIÓN:" & Chr(13) &
                           DTS.Tables(0).Rows(0).Item(2).ToString & Chr(13) &
                           "LÍNEA DE ERROR:" & DTS.Tables(0).Rows(0).Item(3).ToString & Chr(13)

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

    Private Function GetLicense() As Boolean
        Dim valorlicencia As String = "criminalMind2013StormTrooperOneRoguelicenceoperationactive1968isr"
        Dim llavelicencia As String = "NTCX_LBWS_CODE_0001-20131021"
        Dim edgetrailfase As String = "20170825"

        Dim readValue = Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\LcaOperationSWValidate", llavelicencia, Nothing)
        Dim out As Boolean
        Dim datecurrent As DateTime
        Dim datecurrentStr As String
        Dim monthCurrent As String
        Dim dayCurrent As String

        datecurrent = Now()
        out = False
        If datecurrent.Month < 10 Then
            monthCurrent = "0" & datecurrent.Month.ToString
        Else
            monthCurrent = datecurrent.Month.ToString
        End If
        If datecurrent.Day < 10 Then
            dayCurrent = "0" & datecurrent.Day.ToString
        Else
            dayCurrent = datecurrent.Day.ToString
        End If
        datecurrentStr = datecurrent.Year.ToString & monthCurrent & dayCurrent



        If datecurrentStr < edgetrailfase Then
            out = True
        Else
            If readValue IsNot Nothing Then
                If readValue.ToString = valorlicencia Then
                    out = True
                End If
            End If
        End If

        GetLicense = out
    End Function

End Class
