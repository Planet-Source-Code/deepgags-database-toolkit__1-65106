Attribute VB_Name = "ModMain"
Public obDAO As DAO.Workspace, obDB As DAO.Database
Public DBPass As String, strDBPath As String
Public ConnectionState As Boolean
Public TransferMode As Integer

Public Enum TranferTypes
    Import = 1
    Export = 2
End Enum

Sub main()
    ConnectionState = False
    frmMain.Show vbModal
End Sub

Sub OpenConnection()

    Set obDAO = DAO.DBEngine.Workspaces(0)
    Set obDB = obDAO.OpenDatabase(strDBPath, False, False, ";pwd=" & DBPass & "")
    ConnectionState = True
    
End Sub

Sub closeConnection()
    obDB.Close
    ConnectionState = False
'    Set obDB = Nothing
End Sub
 Function GetFldTypeName(fldType As Integer) As String
    Select Case fldType
      
    Case dbBoolean:
        GetFldTypeName = "Boolean"
    Case dbByte:
        GetFldTypeName = "Byte"
    Case dbInteger:
        GetFldTypeName = "Integer"
    Case dbLong:
        GetFldTypeName = "Long"
    Case dbCurrency:
        GetFldTypeName = "Currency"
    Case dbSingle:
        GetFldTypeName = "Single"
    Case dbDouble:
        GetFldTypeName = "Double"
    Case dbDate:
        GetFldTypeName = "Date"
    Case dbText:
        GetFldTypeName = "Text"
    Case dbLongBinary:
        GetFldTypeName = "LongBinary"
    Case dbMemo:
        GetFldTypeName = "Memo"
        
    End Select
          
End Function
