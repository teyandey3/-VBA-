Attribute VB_Name = "A���k�p�����ʏ�����"
Option Explicit

Sub �{�^���@���k�p�����ʏ������v���O����()


    Dim ���k�pws As Worksheet
    Set ���k�pws = Worksheets("���k�p������")

    Dim �\�� As a�\�쐬�p�ϐ��N���X
    Set �\�� = New a�\�쐬�p�ϐ��N���X
    Call �\��.�\�쐬�p�ϐ�������(���k�pws, "���k�pws")

    '�{�^���A�̓�x������h�����߂̃v���O����
    If ���k�pws.Range("K5").Value = "����}�[�J�[(�����Ȃ���)" Then 'K5�Z���iVBA��ʂł̂ݕύX�\�j�̕������擾���āA�����ԈႦ�ă{�^���A��A���œ�x��������A�G���[�R�[�h��\��
        MsgBox "�\�̕ҏW�r���Ń{�^���@�������Ă��܂��B�ҏW�r���ŏ������������ꍇ��" + vbCrLf + "K5�Z���́u����}�[�J�[(�����Ȃ���)�v����������Ƀ{�^���@�������Ă��������B"
        Exit Sub
    End If '�G���[���Ȃ���΁A�ȉ��̃v���O���������s

    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��

    '�\�̍폜
    Dim �\�� As b�폜�N���X
    Set �\�� = New b�폜�N���X
    Call �\��.�\�S�폜(���k�pws, �\��.�\�s�n, �\��.�\��n)

    ���k�pws.Range("B1") = ""             '�Z��B2�̓��e�������iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("E2:O2").ClearContents '�Z��E2����J2�̓��e�������iVBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("E3:O3").ClearContents '�Z��E3����J3�̓��e�������iVBA��ʂł̂ݕύX�\�j
    
    '�������o�쐬���쐬����N���X�̌Ăяo��
    Dim �����o As c�������o�쐬�N���X
    Set �����o = New c�������o�쐬�N���X
    Call �����o.�c�������o�쐬(���k�pws, "���k�pws", �\��.�\�s�n, �\��.�\��n, �\��.�\�s�I, �\��.�R�}��)

    '�r���������N���X�̌Ăяo��
    Dim �r�� As d�r�������N���X
    Set �r�� = New d�r�������N���X
    Call �r��.�r������(���k�pws, �\��.�\�s�n, �\��.�\��n, �\��.�\�s�I, �\��.�\��I, �\��.�R�}��)
    
    '�I��͈͂ɕ��������
    ���k�pws.Cells(�\��.�\�s�n, �\��.�J�n����).Value = "�J�n��"
    
    '���̍�Ƃ��w�����郁�b�Z�[�W�̕\��
    MsgBox "�u�K�J�n����B17�̃Z���ɓ��͂��Ă��������B"

    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n


End Sub

