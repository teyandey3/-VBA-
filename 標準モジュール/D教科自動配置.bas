Attribute VB_Name = "d���Ȏ����z�u"
Option Explicit

Sub �{�^���C���Ȃ������ŐU�蕪����v���O����()


    Dim ���k�pws As Worksheet
    Set ���k�pws = Worksheets("���k�p������")

    Dim �\�� As a�\�쐬�p�ϐ��N���X
    Set �\�� = New a�\�쐬�p�ϐ��N���X
    Call �\��.�\�쐬�p�ϐ�������(���k�pws, "���k�pws")
    
    '���͂��ꂽ�e���Ȃ̍��v����ϐ��Ɋi�[
    Dim �G�R�}�����v As Integer: �G�R�}�����v = ���k�pws.Range("E2").Value - 1 '�z��0�Ԗڂ���n�܂邽��-1����
    
    '�e���Ȃ̃R�}���̍��v����]�\�ɂ���R�}���̍��v�ƈ�v���Ă��邩�m�F
    If Not ���k�pws.Range("E3").Value - 1 = �G�R�}�����v Then
        MsgBox "�G�ߍu�K�̊e���Ȃ̃R�}���̍��v�Ɗ�]�R�}���̍��v����v���Ă��܂���B" '������v���Ă��Ȃ���΁A�G���[���b�Z�[�W��\��
        Exit Sub
    ElseIf ���k�pws.Range("E3").Value = "" _
    Or ���k�pws.Range("E3").Value = 0 Then
        MsgBox "��L�̋G�ߍu�K�̃R�}���\�Ɋe���Ȃ̃R�}������͂��Ă��������B" '�������v��0�Ȃ�A�G���[���b�Z�[�W��\��
        Exit Sub
    End If '������v���Ă�����A�ȉ��̃v���O���������s
    
    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��
    
    '�V���b�t����̔z��"����"���Z���ɓ\��t��
    Dim �\�t As z���Ȏ����z�u�N���X
    Set �\�t = New z���Ȏ����z�u�N���X
    
    Call �\�t.���ȒT���ƃZ���\�t(���k�pws, �\��.�\�s�n, �\��.�\��n, �\��.�\�s�I, �\��.�\��I, �G�R�}�����v)
    
    '�I��͈͂�����
    ���k�pws.Range("E3").ClearContents
    
    '�֐���}���i�����t�������ŐF��t���������Ȃ𑝂₷�Ƃ��͂��̃R�[�h�𑝂₷�j
    ���k�pws.Range("E3").Formula = "=SUM(F3:J3)"               '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("F3").Formula = "=COUNTIF(17:1048576,R13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("G3").Formula = "=COUNTIF(17:1048576,S13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("H3").Formula = "=COUNTIF(17:1048576,T13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("I3").Formula = "=COUNTIF(17:1048576,U13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("J3").Formula = "=COUNTIF(17:1048576,V13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("K3").Formula = "=COUNTIF(17:1048576,W13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("L3").Formula = "=COUNTIF(17:1048576,X13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("M3").Formula = "=COUNTIF(17:1048576,Y13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("N3").Formula = "=COUNTIF(17:1048576,Z13)"  '�iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("O3").Formula = "=COUNTIF(17:1048576,AA13)" '�iVBA��ʂł̂ݕύX�\�j

    '���̍�Ƃ��w�����郁�b�Z�[�W�̕\��
    MsgBox "�u�m�F�p�v�̗��̐��������͂����R�}���Ɠ��������m�F����" + vbCrLf + "�{�^���C�������Ă��������B"

    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n


End Sub
