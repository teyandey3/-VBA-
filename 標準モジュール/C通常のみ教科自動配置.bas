Attribute VB_Name = "C�ʏ�̂݋��Ȏ����z�u"
Option Explicit

Sub �{�^���B�ʏ틳�Ȃ݂̂̕\���쐬����v���O����()


    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��

    Dim ���k�pws As Worksheet
    Set ���k�pws = Worksheets("���k�p������")

    Dim �\�� As a�\�쐬�p�ϐ��N���X
    Set �\�� = New a�\�쐬�p�ϐ��N���X
    Call �\��.�\�쐬�p�ϐ�������(���k�pws, "���k�pws")

    '�V���b�t����̔z��"����"���Z���ɓ\��t��
    Dim �\�t As z���Ȏ����z�u�N���X
    Set �\�t = New z���Ȏ����z�u�N���X
    
    '�ʏ틳�Ȃ̂ݓ��͂��邽�߁A�R�����v�ɂ�0��������B
    Call �\�t.���ȒT���ƃZ���\�t(���k�pws, �\��.�\�s�n, �\��.�\��n, �\��.�\�s�I, �\��.�\��I, 0)

    '���̍�Ƃ��w�����郁�b�Z�[�W�̕\��
    MsgBox "�{�^���D�������Ă��������B"

    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n


End Sub
