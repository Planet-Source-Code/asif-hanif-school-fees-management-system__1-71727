Attribute VB_Name = "connection"
Public cn As New ADODB.connection
Public rsSTUDENT As New ADODB.Recordset
Public rssubjects As New ADODB.Recordset
Public rsfeepayment As New ADODB.Recordset
Public rsuser As New ADODB.Recordset
Public rsname As New ADODB.Recordset
Public rsuserlog As New ADODB.Recordset
Public schoolname As String
Public address1 As String
Public address2 As String
Public pubuserid As String
Public pubusername As String
Public pubpassword As String



Public Sub conopen()
Set cn = New ADODB.connection
' set database password
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=false;Data Source= " & App.Path & "\database\school.mdb;Jet OLEDB:Database Password="
' set database without password
'cn.Open "Provider=Microsoft.jet.oledb.4.0; Data Source=" & App.Path & "\database\school.mdb"
cn.CursorLocation = adUseClient
With rsSTUDENT
If .State = closed Then
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .ActiveConnection = cn
End If
End With
With rssubjects
If .State = closed Then
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .ActiveConnection = cn
End If
End With
With rsfeepayment
If .State = closed Then
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .ActiveConnection = cn
End If
End With
With rsuser
If .State = closed Then
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .ActiveConnection = cn
End If
End With
With rsname
If .State = closed Then
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .ActiveConnection = cn
End If
End With
End Sub
Public Sub studentopen()
rsSTUDENT.Open "select * from student order by grno", cn, adOpenStatic, adLockOptimistic
End Sub
Public Sub subjectsopen()
rssubjects.Open "select * from subjects order by levelcode", cn, adOpenStatic, adLockOptimistic
End Sub
Public Sub feepaymentopen()
rsfeepayment.Open "select * from feepayment", cn, adOpenStatic, adLockOptimistic
End Sub
Public Sub useropen()
rsuser.Open "select * from user_details", cn, adOpenStatic, adLockOptimistic
End Sub
Public Sub nameopen()
rsname.Open "select * from schoolname", cn, adOpenStatic, adLockOptimistic
End Sub
Public Sub userlogopen()
rsuserlog.Open "select * from userlog", cn, adOpenDynamic, adLockOptimistic
End Sub
