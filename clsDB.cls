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
Private cn As ADODB.Connection   'ADO�R�l�N�V�����I�u�W�F�N�g
Private rs As ADODB.Recordset    'ADO���R�[�h�Z�b�g�I�u�W�F�N�g

'�O����Access�t�@�C�����w�肵�Đڑ�����ꍇ
Public Function cnOpen(ByVal sProvider As String, ByVal sSource As String) As Boolean
On Error GoTo Step_Err
    
    Set cn = New ADODB.Connection 'ADO�R�l�N�V�����̃C���X�^���X�쐬
    cn.Open sProvider & sSource 'Access�t�@�C�����w��
    
    cnOpen = True
    
    Exit Function
    
Step_Err:
    cnOpen = False
    
End Function

'�O����Access�t�@�C�����w�肵�Đؒf����ꍇ
Public Function cnClose() As Boolean
On Error GoTo Step_Err
    cn.Close: Set cn = Nothing
    
    cnClose = True
    
    Exit Function
    
Step_Err:
    cnClose = False
    
End Function

'�O����Access�t�@�C�����w�肵�ă��R�[�h�Z�b�g�ڑ�����ꍇ
Public Function rsOpen(ByVal sSQL As String) As Boolean
On Error GoTo Step_Err
    
    Set rs = New ADODB.Recordset 'ADO���R�[�h�Z�b�g�̃C���X�^���X�쐬
    rs.Open sSQL, cn, adOpenKeyset, adLockOptimistic '���R�[�h���o
    
    rsOpen = True
    
    Exit Function
Step_Err:
    rsOpen = False
    
End Function

'�O����Access�t�@�C�����w�肵�ă��R�[�h�Z�b�g�ؒf����ꍇ
Public Function rsClose() As Boolean
On Error GoTo Step_Err
    rs.Close: Set rs = Nothing
    
    rsClose = True
    
    Exit Function
    
Step_Err:
    rsClose = False
    
End Function

'�Y����Field�P�ƃf�[�^���擾����B
Public Function getData(ByVal sString As String, ByVal sField As String) As String
    rs.Filter = sString
    If rs.RecordCount = 1 Then
        getData = rs.Fields(sField)
    Else
        getData = ""
        Exit Function
    End If
    
End Function

'�Y����Field���R�[�h�i���o�[�ɂ��擾���s���B
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

'�Y����Field�P�ƃf�[�^��ǉ�����B
Public Function addData(ByVal sString As String, ByVal sField As String, ByVal sValue As String) As Boolean
    cn.BeginTrans   '�g�����U�N�V�����ݒ�
    
    rs.Filter = sString
    If rs.RecordCount = 0 Then
        rs.AddNew sField, sValue
        rs.Update
        cn.CommitTrans  '�g�����U�N�V�����R�~�b�g
        addData = True
    Else
        cn.RollbackTrans    '�g�����U�N�V�������[���o�b�N
        addData = False
        Exit Function
    End If
    
End Function

'�Y����Field�����f�[�^��ǉ�����B
Public Function addDataF(ByVal sField As String, ByVal sValue As String) As Boolean
On Error GoTo Step_Err
    cn.BeginTrans   '�g�����U�N�V�����ݒ�
    
    rs.AddNew sField, sValue
    rs.Update
    cn.CommitTrans  '�g�����U�N�V�����R�~�b�g
    addDataF = True
    Exit Function
        
Step_Err:
    cn.RollbackTrans    '�g�����U�N�V�������[���o�b�N
    addDataF = False
    
End Function

'�Y����Field�P�ƃf�[�^���X�V����B
Public Function crtData(ByVal sString As String, ByVal sField As String, ByVal sValue As String) As Boolean
    cn.BeginTrans   '�g�����U�N�V�����ݒ�
    
    rs.Filter = sString
    If rs.RecordCount = 1 Then
        rs.Update sField, sValue
        cn.CommitTrans  '�g�����U�N�V�����R�~�b�g
        crtData = True
    Else
        cn.RollbackTrans    '�g�����U�N�V�������[���o�b�N
        crtData = False
        Exit Function
    End If
    
End Function

'�Y����Field�P�ƃf�[�^���X�V����B
Public Function crtDataF(ByVal sField As String, ByVal sValue As String) As Boolean
    cn.BeginTrans   '�g�����U�N�V�����ݒ�
    
    rs.Update sField, sValue
    cn.CommitTrans  '�g�����U�N�V�����R�~�b�g
    crtDataF = True
    Exit Function

Step_Err:
    cn.RollbackTrans    '�g�����U�N�V�������[���o�b�N
    crtDataF = False
    
End Function

'�Y����Field�P�ƃf�[�^���폜����B
Public Function delData(ByVal sString As String) As Boolean
    cn.BeginTrans   '�g�����U�N�V�����ݒ�
    
    rs.Filter = sString
    If rs.RecordCount = 1 Then
        rs.Delete
        rs.Update
        cn.CommitTrans  '�g�����U�N�V�����R�~�b�g
        delData = True
    Else
        cn.RollbackTrans    '�g�����U�N�V�������[���o�b�N
        delData = False
        Exit Function
    End If
    
End Function

'�Y���v���f�[�^�̌�����Ԃ��B
Public Function moveCount(ByVal sString As String, ByVal sField As String) As Long
    rs.Filter = sString
    moveCount = rs.RecordCount

End Function

'SQL���s
Public Function runSQL(ByVal sSQL As String)
    cn.Execute sSQL
End Function