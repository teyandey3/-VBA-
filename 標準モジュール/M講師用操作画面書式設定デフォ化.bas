Attribute VB_Name = "M�u�t�p�����ʏ����ݒ�f�t�H��"
Option Explicit

Sub �{�^���N�u�t�p������_�\�̐ݒ�����ׂăf�t�H���g�ɂ���()


    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��

    Dim �u�t�pws As Worksheet
    Set �u�t�pws = Worksheets("�u�t�p������")

    Dim �f�t�H�� As h�����ݒ�f�t�H���g��
    Set �f�t�H�� = New h�����ݒ�f�t�H���g��
    Call �f�t�H��.�����ݒ�f�t�H���g��(�u�t�pws)

    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n

    
End Sub

