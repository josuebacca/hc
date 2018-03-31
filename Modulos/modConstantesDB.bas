Attribute VB_Name = "modConstantesDB"
'declaro constantes para manejo de tipos de datos ADO
Public Const dbSqlVarchar = 200
Public Const dbSqlSmallint = 2
Public Const dbSqlNumeric = 131
Public Const dbSqlDate = 135
Public Const dbSqlInt = 3
Public Const dbSqlChar = 125

'declaro constantes para manejo de errores de transacciones ADO
Public Const dbSqlDuplicateKey = 2601
Public Const dbSqlLoginFailed = 4002
Public Const dbSqlDataSource = 0
Public Const dbSqlDefaultDrivers = 0
Public Const dbSqlPermission = 229

'declaro constantes con tipos de Motores de Base de Datos
Public Const dbEngineSQLServer = "SQLServer"
Public Const dbEngineInformix5 = "Informix 5"

'declaro sintaxis de SQL de la función de fecha del motor
Public Const dbDateSQLServer = "GetDate()"
Public Const dbDateInformix5 = "CURRENT"

