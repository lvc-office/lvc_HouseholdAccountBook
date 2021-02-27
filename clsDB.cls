VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event runEvent()
Private cn As ADODB.Connection   'ADOコネクションオブジェクト
Private rs As ADODB.Recordset    'ADOレコードセットオブジェクト

'外部のAccessファイルを指定して接続する場合
Public Function cnOpen(ByVal sProvider As String, ByVal sSource As String) As Boolean
On Error GoTo Step_Err
    
    Set cn = New ADODB.Connection 'ADOコネクションのインスタンス作成
    cn.Open sProvider & sSource 'Accessファイルを指定
    
    cnOpen = True
    
    Exit Function
    
Step_Err:
    cnOpen = False
    
End Function

'外部のAccessファイルを指定して切断する場合
Public Function cnClose() As Boolean
On Error GoTo Step_Err
    cn.Close: Set cn = Nothing
    
    cnClose = True
    
    Exit Function
    
Step_Err:
    cnClose = False
    
End Function

'外部のAccessファイルを指定してレコードセット接続する場合
Public Function rsOpen(ByVal sSQL As String) As Boolean
On Error GoTo Step_Err
    
    Set rs = New ADODB.Recordset 'ADOレコードセットのインスタンス作成
    rs.Open sSQL, cn, adOpenKeyset, adLockOptimistic 'レコード抽出
    
    rsOpen = True
    
    Exit Function
Step_Err:
    rsOpen = False
    
End Function

'外部のAccessファイルを指定してレコードセット切断する場合
Public Function rsClose() As Boolean
On Error GoTo Step_Err
    rs.Close: Set rs = Nothing
    
    rsClose = True
    
    Exit Function
    
Step_Err:
    rsClose = False
    
End Function

'該当のField単独データを取得する。
Public Function getData(ByVal sString As String, ByVal sField As String) As String
    rs.Filter = sString
    If rs.RecordCount = 1 Then
        getData = rs.Fields(sField)
    Else
        getData = ""
        Exit Function
    End If
    
End Function

'該当のFieldレコードナンバーによる取得を行う。
Public Function getNumData(ByVal sString, ByVal sField As String, ByVal iNum As Long) As String
On Error GoTo Step_Err
    rs.Filter = sString
    If rs.RecordCount >= 1 Then
        rs.AbsolutePosition = iNum
        getNumData = rs.Fields(sField)
    Else
        getNumData = ""
        Exit Function
    End If
    Exit Function
    
Step_Err:
    getNumData = ""
End Function

'該当のField単独データを追加する。
Public Function addData(ByVal sString As String, ByVal sField As String, ByVal sValue As String) As Boolean
    cn.BeginTrans   'トランザクション設定
    
    rs.Filter = sString
    If rs.RecordCount = 0 Then
        rs.AddNew sField, sValue
        rs.Update
        cn.CommitTrans  'トランザクションコミット
        addData = True
    Else
        cn.RollbackTrans    'トランザクションロールバック
        addData = False
        Exit Function
    End If
    
End Function

'該当のField複数データを追加する。
Public Function addDataF(ByVal sField As String, ByVal sValue As String) As Boolean
On Error GoTo Step_Err
    cn.BeginTrans   'トランザクション設定
    
    rs.AddNew sField, sValue
    rs.Update
    cn.CommitTrans  'トランザクションコミット
    addDataF = True
    Exit Function
        
Step_Err:
    cn.RollbackTrans    'トランザクションロールバック
    addDataF = False
    
End Function

'該当のField単独データを更新する。
Public Function crtData(ByVal sString As String, ByVal sField As String, ByVal sValue As String) As Boolean
    cn.BeginTrans   'トランザクション設定
    
    rs.Filter = sString
    If rs.RecordCount = 1 Then
        rs.Update sField, sValue
        cn.CommitTrans  'トランザクションコミット
        crtData = True
    Else
        cn.RollbackTrans    'トランザクションロールバック
        crtData = False
        Exit Function
    End If
    
End Function

'該当のField単独データを更新する。
Public Function crtDataF(ByVal sField As String, ByVal sValue As String) As Boolean
    cn.BeginTrans   'トランザクション設定
    
    rs.Update sField, sValue
    cn.CommitTrans  'トランザクションコミット
    crtDataF = True
    Exit Function

Step_Err:
    cn.RollbackTrans    'トランザクションロールバック
    crtDataF = False
    
End Function

'該当のField単独データを削除する。
Public Function delData(ByVal sString As String) As Boolean
    cn.BeginTrans   'トランザクション設定
    
    rs.Filter = sString
    If rs.RecordCount = 1 Then
        rs.Delete
        rs.Update
        cn.CommitTrans  'トランザクションコミット
        delData = True
    Else
        cn.RollbackTrans    'トランザクションロールバック
        delData = False
        Exit Function
    End If
    
End Function

'該当要求データの件数を返す。
Public Function moveCount(ByVal sString As String, ByVal sField As String) As Long
    rs.Filter = sString
    moveCount = rs.RecordCount

End Function

'SQL実行
Public Function runSQL(ByVal sSQL As String)
    cn.Execute sSQL
End Function
