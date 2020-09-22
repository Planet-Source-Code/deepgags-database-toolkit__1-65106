VERSION 5.00
Begin VB.Form frmDatabaseUpdate 
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   1530
      TabIndex        =   0
      Top             =   1170
      Width           =   1515
   End
End
Attribute VB_Name = "frmDatabaseUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TableName(5) As String
Dim QueryName(1) As String
Private Sub Form_Load()
    TableName(0) = "Appoint_New"
    TableName(1) = "Appointments"
    TableName(2) = "Fee_Rcpts"
    TableName(3) = "Fees_Master"
    TableName(4) = "Rcpt_Details"
    QueryName(0) = "PetInfo"
End Sub
Private Sub Command1_Click()
     Call CreateNewTables(App.Path & "\vet1.mdb", "mdecodc")
     Call CreateNewQueries(App.Path & "\vet1.mdb", "mdecodc")
End Sub
'****************************************************************
'                    Creates Connection with database
'****************************************************************
Public Function CreateNewTables(DBPath As String, Optional DBPass As String)
    Dim obDAO As DAO.Workspace, obDB As DAO.Database
    Dim i As Integer
    On Error GoTo Errhandler
'open database
    Set obDAO = DAO.DBEngine.Workspaces(0)
    Set obDB = obDAO.OpenDatabase(DBPath, False, False, ";pwd=" & DBPass & "")
    
'create common tables
    For i = 0 To 4
        If TableExists(obDB, TableName(i)) = False Then
            CreateNewTable i, obDB
        End If
    Next i
'close objects
    obDB.Close
    Set obDB = Nothing
    obDAO.Close
    Set obDAO = Nothing

    MsgBox "New database has been successfully created.", vbInformation + vbOKOnly, App.Title
    Exit Function

Errhandler:
    Set obDB = Nothing
    Set obDAO = Nothing
    MsgBox Err.Number & "" & Err.Description
    Resume Next
End Function
'****************************************************************
'                    Creates Connection with database
'****************************************************************
Public Function CreateNewQueries(DBPath As String, Optional DBPass As String)
    Dim obDAO As DAO.Workspace, obDB As DAO.Database
    Dim i As Integer
    On Error GoTo Errhandler
'open database
    Set obDAO = DAO.DBEngine.Workspaces(0)
    Set obDB = obDAO.OpenDatabase(DBPath, False, False, ";pwd=" & DBPass & "")
'create common tables
    For i = 0 To 0
        If QueryExists(obDB, QueryName(i)) = False Then
            CreateNewQuery i, obDB
        End If
    Next i
        
'close objects
    obDB.Close
    Set obDB = Nothing
    obDAO.Close
    Set obDAO = Nothing

'    MsgBox "New database has been successfully created.", vbInformation + vbOKOnly, App.Title
    Exit Function

Errhandler:
    Set obDB = Nothing
    Set obDAO = Nothing
    Resume Next
End Function

'****************************************************************
'                    Create Query in database
'****************************************************************
Private Sub CreateNewQuery(Index As Integer, pobDB As DAO.Database)
   Dim tdfNew As DAO.QueryDef
'create table object
    Set tdfNew = pobDB.CreateQueryDef(QueryName(Index))
     Select Case Index
        Case 0:
        tdfNew.SQL = "select p.clinic_code,p.clinic_code as Pat_Code,p.p_name,o.o_name,(o.Address +' '+ o.Address1+' ' +o.city) as Addr,o.phone1,o.phone2,o.mobile  from pet p,owner o where o.owner_code=p.owner_code UNION Select 'New' as clinic_code,pat_code,pat_name,Owner_Name,iif(isnull(Address),'---',Address) as Addr, phone1,'---' as phone2, mobile from Appoint_new"
    End Select
    pobDB.QueryDefs.Append tdfNew
    Set tdfNew = Nothing
End Sub
'****************************************************************
'                    Create Table in database
'****************************************************************
Private Sub CreateNewTable(Index As Integer, pobDB As DAO.Database)

    Dim tdfNew As DAO.TableDef
'create table object

    Set tdfNew = pobDB.CreateTableDef(TableName(Index))
    
    Select Case Index
        Case 0:
            'add table fields
                AddTableField tdfNew, "pat_code", dbText, 50
                AddTableField tdfNew, "pat_Name", dbText, 50
                AddTableField tdfNew, "owner_name", dbText, 100
                AddTableField tdfNew, "Address", dbText, 50
                AddTableField tdfNew, "Phone1", dbText, 50
                AddTableField tdfNew, "Mobile", dbText, 50
        Case 1:
                AddTableField tdfNew, "Appoint_Id", dbDouble
                AddTableField tdfNew, "pat_code", dbText, 250
                AddTableField tdfNew, "appoint_date", dbDate
                AddTableField tdfNew, "appoint_time", dbDate
                AddTableField tdfNew, "Purpose", dbText, 100
                AddTableField tdfNew, "Pat_type", dbText, 10
        Case 2:
                AddTableField tdfNew, "Receipt_No", dbInteger
                AddTableField tdfNew, "Clinic_code", dbText, 150
                AddTableField tdfNew, "Total", dbInteger
                AddTableField tdfNew, "RctDate", dbDate
                AddTableField tdfNew, "Pat_Type", dbText, 50
        Case 3:
                AddTableField tdfNew, "Receipt_Id", dbInteger
                AddTableField tdfNew, "Receipt_Type", dbText, 150
                AddTableField tdfNew, "Charges", dbInteger
        Case 4:
                AddTableField tdfNew, "Receipt_No", dbInteger
                AddTableField tdfNew, "SrNo", dbInteger
                AddTableField tdfNew, "Receipt_Id", dbInteger
                AddTableField tdfNew, "Rate", dbInteger
                AddTableField tdfNew, "Charged", dbInteger
    End Select
'append collection
    pobDB.TableDefs.Append tdfNew
    Set tdfNew = Nothing
End Sub
'****************************************************************
'                    Add Fields in Table
'****************************************************************
Private Sub AddTableField(ptdfTableDef As DAO.TableDef, pstrFieldName As String, pintDatatype As Integer, Optional pintSize As Integer, Optional pblnFixedLength As Boolean, Optional pblnAutoIncrement As Boolean, Optional pblnAllowZeroLength As Boolean, Optional pblnRequired As Boolean, Optional pstrValidationText As String, Optional pstrValidationRule As String, Optional pvarDefaultValue As Variant)
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
'                    Check Table Exists or Not
'****************************************************************
Private Function TableExists(DBObject As DAO.Database, TabName As String) As Boolean
    Dim TableObj As DAO.TableDef
    TableExists = False
    For Each TableObj In DBObject.TableDefs
        On Error Resume Next
        If TableObj.Name = TabName Then
            TableExists = True
            Exit Function
        End If
    Next TableObj
End Function
'****************************************************************
'                    Check Query Exists or Not
'****************************************************************
Private Function QueryExists(DBObject As DAO.Database, QueryName As String) As Boolean
    Dim QueryObj As DAO.QueryDef
    QueryExists = False
    For Each QueryObj In DBObject.QueryDefs
        On Error Resume Next
        If QueryObj.Name = QueryName Then
            QueryExists = True
            Exit Function
        End If
    Next QueryObj
End Function
'****************************************************************
'                    Check Query Properties
'****************************************************************
                    
'Private Sub chkQueryProperties(QueryObj As DAO.QueryDef)
'    Dim QryPropObj As DAO.Property
'    For Each QryPropObj In QueryObj.Properties
'        On Error Resume Next
'        If QryPropObj <> "" Then Debug.Print "    " & QryPropObj.Name & " = " & QryPropObj
'            On Error GoTo 0
'    Next QryPropObj
'End Sub
'****************************************************************
'                    Check Table Properties
'****************************************************************
'Private Sub ResetTableProperties(ptdfTableDef As DAO.TableDef)
'    Dim prpLoop As DAO.Property
'    For Each prpLoop In ptdfTableDef.Properties
'        On Error Resume Next
'        If prpLoop <> "" Then Debug.Print "    " & prpLoop.Name & " = " & prpLoop
'            On Error GoTo 0
'    Next prpLoop
'End Sub
