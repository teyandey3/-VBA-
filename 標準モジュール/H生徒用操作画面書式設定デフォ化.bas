Attribute VB_Name = "h���k�p�����ʏ����ݒ�f�t�H��"
Option Explicit

Sub �{�^���G���k�p������_�\�̐ݒ�����ׂăf�t�H���g�ɂ���()


    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��

    Dim ���k�pws As Worksheet
    Set ���k�pws = Worksheets("���k�p������")

    Dim �f�t�H�� As h�����ݒ�f�t�H���g��
    Set �f�t�H�� = New h�����ݒ�f�t�H���g��
    Call �f�t�H��.�����ݒ�f�t�H���g��(���k�pws)

    '���Ȗ��̏�����
    ���k�pws.Range("Q11:XFD11").ClearContents
    ���k�pws.Range("Q13:XFD13").ClearContents
    
    Dim �ʋ��Ȗ� As Variant: �ʋ��Ȗ� = Array("�p��", "���w", "����", "����", "�Љ�")
    Dim �G���Ȗ� As Variant: �G���Ȗ� = Array("�G�@�p��", "�G�@���w", "�G�@����", "�G�@����", "�G�@�Љ�")
    ���k�pws.Range("R11:V11").Value = �ʋ��Ȗ�
    ���k�pws.Range("R13:V13").Value = �G���Ȗ�
    
    '�G�߉Ȗڐ��̏�����
    ���k�pws.Range("Q11").Value = 5
    
    '�ʏ�Ȗڐ��̏�����
    ���k�pws.Range("Q13").Value = 5
    
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n

  
End Sub
