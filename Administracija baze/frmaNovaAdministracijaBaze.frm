VERSION 5.00
Begin VB.Form frmaNovaAdministracijaBaze 
   Caption         =   "Administracija baze"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Shrink"
      Height          =   510
      Left            =   135
      TabIndex        =   3
      Top             =   1305
      Width           =   4380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CheckDB"
      Height          =   465
      Left            =   135
      TabIndex        =   2
      Top             =   720
      Width           =   4380
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   90
      TabIndex        =   1
      Top             =   1980
      Width           =   4380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reindex"
      Height          =   465
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   4380
   End
End
Attribute VB_Name = "frmaNovaAdministracijaBaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim cmd As New ADODB.Command

Private Sub Command1_Click()

Dim rsTablica As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim bOld As Boolean
con.Open
Set cmd.ActiveConnection = con
rs.Open "SELECT (@@MICROSOFTVERSION / POWER(2, 24))", con, adOpenKeyset, adLockOptimistic
List1.Clear
If rs.Fields(0) > 8 Then
    bOld = False
Else
    bOld = True
End If
rs.Close
rsTablica.Open "SELECT table_name as tableName FROM INFORMATION_SCHEMA.TABLES  WHERE table_type = 'BASE TABLE'", con, adOpenKeyset, adLockOptimistic
Do While Not rsTablica.EOF
    List1.AddItem Now & " - " & rsTablica.Fields(0), 0
    DoEvents
    If bOld Then
        cmd.CommandText = "DBCC DBREINDEX([" & rsTablica.Fields(0) & "],'',90)"
    Else
       cmd.CommandText = "ALTER INDEX ALL ON " & rsTablica.Fields(0) & " REBUILD WITH (FILLFACTOR = 90)"
    End If
    cmd.Execute
    rsTablica.MoveNext
Loop
rsTablica.Close
List1.AddItem Now & " - Gotovo reindeksiranje...", 0
con.Close
End Sub

Private Sub Command2_Click()
con.Open
List1.AddItem Now & " - CheckDB started...", 0
DoEvents
cmd.CommandTimeout = 3000
Set cmd.ActiveConnection = con
cmd.CommandText = "DBCC CHECKDB ('" & con.DefaultDatabase & "',NOINDEX ) WITH  NO_INFOMSGS"
cmd.Execute
con.Close
List1.AddItem Now & " - CheckDB end...", 0
DoEvents

End Sub

Private Sub Command3_Click()
con.Open
List1.AddItem Now & " - Shrinking started...", 0
DoEvents
cmd.CommandTimeout = 3000
Set cmd.ActiveConnection = con
cmd.CommandText = "DBCC SHRINKDATABASE  ('" & con.DefaultDatabase & "' ) WITH  NO_INFOMSGS"
cmd.Execute
con.Close
List1.AddItem Now & " - Shrinking ended...", 0
DoEvents
End Sub

Private Sub Form_Load()
con.ConnectionString = "FILE NAME=" & App.Path & "\KONEKCIJA.UDL"
End Sub
