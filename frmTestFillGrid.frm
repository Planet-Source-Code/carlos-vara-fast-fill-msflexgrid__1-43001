VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTestFillGrid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TestFillGrid"
   ClientHeight    =   5280
   ClientLeft      =   1770
   ClientTop       =   2265
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6495
   Begin VB.CommandButton Command1 
      Caption         =   "Fill MSFlexGrid"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2235
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   60
      TabIndex        =   0
      Top             =   1260
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7011
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Tested under HP Pentium IV, take 2 to 3 sec. even less to populate the MSFlexGrid Control with 18,000 records."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   675
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   3915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Filling Time : "
      Height          =   195
      Left            =   4560
      TabIndex        =   2
      Top             =   1020
      Width           =   1815
   End
End
Attribute VB_Name = "frmTestFillGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cnn As ADODB.Connection

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Sub Command1_Click()
  
  Dim tmpRst As New ADODB.Recordset
  Dim sSQL   As String
  
  Dim sTime
  
  sSQL = "Select code, desc, quantity, unitprice, quantity*unitprice as Total From tblTest Order By Code;"
  tmpRst.Open sSQL, Cnn, adOpenKeyset, adLockOptimistic, adCmdText
  
  Screen.MousePointer = vbHourglass
  sTime = Time()
  'Fill the MSFlexGrid control
  FillGrid MSFlexGrid1, tmpRst, "^  Code  |<Description                                        |>  Quantity  |>  Unit Price  |>     Total    "
  Label1.Caption = "Filling Time : " & Format(sTime - Time, "HH:MM:SS")
  Screen.MousePointer = vbDefault
  
  tmpRst.Close
  Set tmpRst = Nothing
  
End Sub

Private Sub Form_Load()
  
  ConnectToDb
  MSFlexGrid1.Rows = 2
  MSFlexGrid1.FixedRows = 1
  MSFlexGrid1.FixedCols = 0
  
End Sub

Private Function FillGrid(Grd As MSFlexGrid, Rst As ADODB.Recordset, FormatHeadings As String) As Integer

  On Error GoTo ErrorFillingGrid
  Dim i      As Integer
  Dim r      As Integer
  Dim c      As Integer
  
  FillGrid = 0
  
  LockWindowUpdate MSFlexGrid1.hWnd
  
  Grd.SelectionMode = 1
  Grd.HighLight = 1
  Grd.FocusRect = 0

  'Define the Headers
  If Len(FormatHeadings) > 0 Then
    Grd.FormatString = FormatHeadings
  End If
  
  i = 0
   
  With Grd
  If Rst.EOF Then
    'If the Recordset is Empty then exit funtion an return 1 like error code
    .Rows = 1
    FillGrid = 1
  Else
    r = 1  'The starting row
    c = 0  'The starting column
    .Rows = Rst.RecordCount + 1  'The number of rows
    
    Do While Not Rst.EOF 'Loop trough the cols and rows to put the data.
      For c = 0 To .Cols - 1
        .TextMatrix(r, c) = IIf(IsNull(Rst(c)), vbNullString, Rst(c))
      Next c
      Rst.MoveNext
      r = r + 1
    Loop
  End If
  End With
  
  LockWindowUpdate False
  
  Exit Function
  
ErrorFillingGrid:
    MsgBox "Error No. " & Err & " " & Err.Source & vbCrLf & Err.Description, vbCritical, "Error FillGrid"
    FillGrid = 1
    Resume Next
  

End Function

Private Sub ConnectToDb()
  Dim sCnn          As String
  Dim sPath         As String
  Dim sDatabaseName As String

  sPath = App.Path & "\"
  sDatabaseName = "Test.mdb"
  
  sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;"
  sCnn = sCnn & "Data Source=" & sPath & sDatabaseName & ";"
  sCnn = sCnn & "Jet OLEDB:Database Password=JVG250870;"
  sCnn = sCnn & "Jet OLEDB:Engine Type=5"
  
  
  Set Cnn = New ADODB.Connection
  Cnn.CursorLocation = adUseClient
  Cnn.CommandTimeout = 300

  Cnn.Open sCnn

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cnn.Close
  Set Cnn = Nothing
End Sub
