Module mConexion

    ' Esta variable es global para toda la aplicación por lo que puede usarse en todos los programas
    ' que requieran esta cadena de conexión

    Public strConn As String
    Public strConnPRO As String

    Public Sub CreaCadenaConexion(ByVal Usuario As String, ByVal Password As String)
        strConn = "Server=SERVER-RAID; DataBase=2018ene; User ID=User_PRO; pwd=User_PRO2015"
        strConnPRO = "Server=SERVER-RAID; DataBase=2018ene; User ID=User_PRO; pwd=User_PRO2015" 'guarda la mezcla del mes en produccion
    End Sub

End Module
