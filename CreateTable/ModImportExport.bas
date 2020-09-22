Attribute VB_Name = "ModImportExport"
'****************************************************************
'                    Create Query in database
'****************************************************************
 Sub TransferQuery(FromDb As DAO.Database, QueryName As String, ToDb As DAO.Database)
Dim tdfNew As DAO.QueryDef
Dim tdfOld As DAO.QueryDef
   
   For Each tdfOld In FromDb.QueryDefs
        On Error Resume Next
        If tdfOld.Name = QueryName Then
            Set tdfNew = ToDb.CreateQueryDef(QueryName)
            tdfNew.SQL = tdfOld.SQL
            ToDb.QueryDefs.Append tdfNew
            Exit Sub
        End If
    Next tdfOld
    Set tdfNew = Nothing
    
End Sub

'****************************************************************
'                    Create Table in database
'****************************************************************
 Sub TransferTable(FromDb As DAO.Database, TableName As String, ToDb As DAO.Database)
Dim tdfNew As DAO.TableDef
Dim tdfOld As DAO.TableDef
Dim FieldObj As DAO.Field
   For Each tdfOld In FromDb.TableDefs
        On Error Resume Next
        If tdfOld.Name = TableName Then
            Set tdfNew = ToDb.CreateTableDef(TableName)
            For Each FieldObj In tdfOld.Fields
                 AddTableField tdfNew, FieldObj.Name, FieldObj.Type, FieldObj.Size
            Next FieldObj
            ToDb.TableDefs.Append tdfNew
            Exit Sub
        End If
    Next tdfOld
    Set tdfNew = Nothing
End Sub
     

'****************************************************************
'                    Add Fields in Table
'****************************************************************
 Sub AddTableField(ptdfTableDef As DAO.TableDef, pstrFieldName As String, pintDatatype As Integer, Optional pintSize As Integer, Optional pblnFixedLength As Boolean, Optional pblnAutoIncrement As Boolean, Optional pblnAllowZeroLength As Boolean, Optional pblnRequired As Boolean, Optional pstrValidationText As String, Optional pstrValidationRule As String, Optional pvarDefaultValue As Variant)
    Dim fldX As DAO.Field
    Dim intSize As Integer

    Set fldX = ptdfTableDef.CreateField()

    fldX.OrdinalPosition = ptdfTableDef.Fields.Count

    fldX.Name = pstrFieldName
    fldX.Type = pintDatatype

    If Not IsMissing(pblnAllowZeroLength) Then
        If pintDatatype = dbText And pintSize > 0 Then
            fldX.Size = pintSize
        Else
            fldX.Size = GetFieldSize(pintDatatype)
        End If
    Else
        fldX.Size = GetFieldSize(pintDatatype)
    End If

    If fldX.Type = dbLong Then
        If pblnAutoIncrement = True Then
            fldX.Attributes = fldX.Attributes Or dbAutoIncrField
        End If
    End If

    If fldX.Type = dbText Or fldX.Type = dbMemo Then
            fldX.AllowZeroLength = True
    End If

    If Not IsMissing(pblnRequired) Then
        fldX.Required = pblnRequired
    End If

    If Not IsMissing(pstrValidationText) Then
        fldX.ValidationText = pstrValidationText
    End If

    If Not IsMissing(pstrValidationRule) Then
        fldX.ValidationRule = pstrValidationRule
    End If

    If Not IsMissing(pvarDefaultValue) Then
        fldX.DefaultValue = pvarDefaultValue
    End If

    ptdfTableDef.Fields.Append fldX

End Sub
'****************************************************************
'                    returns the size of the fields
'****************************************************************
Function GetFieldSize(pintDatatype As Integer) As Integer
  'return field length
  Select Case pintDatatype
    Case dbBoolean
      GetFieldSize = 1
    Case dbByte
      GetFieldSize = 1
    Case dbInteger
      GetFieldSize = 2
    Case dbLong
      GetFieldSize = 4
    Case dbCurrency
      GetFieldSize = 8
    Case dbSingle
      GetFieldSize = 4
    Case dbDouble
      GetFieldSize = 8
    Case dbDate
      GetFieldSize = 8
    Case dbText
      GetFieldSize = 50
    Case dbLongBinary
      GetFieldSize = 0
    Case dbMemo
      GetFieldSize = 0
  End Select
End Function

'****************************************************************
'                    adds the field from speciefied table
'****************************************************************

Public Sub AddField(TabName As String, FldName As String, ToDb As DAO.Database, FromDb As DAO.Database)
Dim tdfNew As DAO.TableDef
Dim tdfOld As DAO.TableDef
Dim FieldObj As DAO.Field

        Set FieldObj = ToDb.TableDefs(Trim(TabName)).CreateField()
        AddTableField ToDb.TableDefs(Trim(TabName)), FldName, FromDb.TableDefs(Trim(TabName)).Fields(Trim(FldName)).Type, FromDb.TableDefs(Trim(TabName)).Fields(Trim(FldName)).Size
       Set tdfNew = Nothing
End Sub

