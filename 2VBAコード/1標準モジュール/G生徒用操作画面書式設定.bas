Attribute VB_Name = "G���k�p�����ʏ����ݒ�"
Option Explicit

Sub �{�^���F���k�p�����ʏ����ݒ�()


    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��

    Dim ���k�pws As Worksheet
    Set ���k�pws = Worksheets("���k�p������")
   
    Dim ���� As g���C�A�E�g�Ə����N���X
    Set ���� = New g���C�A�E�g�Ə����N���X

    Call ����.�y�[�W���C�A�E�g�ƕ�������(���k�pws, ���k�pws, "���k�pws")

    'K4�Z���̕�������
    ���k�pws.Range("K4").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    'K5�Z���̕�������
    ���k�pws.Range("K5").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    'Q3�Z���̕�������
    ���k�pws.Range("U4").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    'AF1�Z���̕�������
    ���k�pws.Range("AF1").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    'AF2�Z���̕�������
    ���k�pws.Range("AF2").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n
    

End Sub
