Attribute VB_Name = "L�u�t�p�����ʏ����ݒ�"
Option Explicit

Sub �{�^���M�u�t�p�����ʏ����ݒ�()


    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��

    Dim �u�t�pws As Worksheet
    Set �u�t�pws = Worksheets("�u�t�p������")
   
    Dim ���� As g���C�A�E�g�Ə����N���X
    Set ���� = New g���C�A�E�g�Ə����N���X

    Call ����.�y�[�W���C�A�E�g�ƕ�������(�u�t�pws, �u�t�pws, "�u�t�pws")

    'K5�Z���̕�������
    �u�t�pws.Range("K5").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    'Q3�Z���̕�������
    �u�t�pws.Range("U4").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    
    'AF1�Z���̕�������
    �u�t�pws.Range("AF1").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���
    'AF2�Z���̕�������
    �u�t�pws.Range("AF2").HorizontalAlignment = xlLeft '���������̕����z�u���������ɂ���

    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n


End Sub
