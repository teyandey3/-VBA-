Attribute VB_Name = "E���k�p�����ʕ���"
Option Explicit

Sub �{�^���D���k�p�����ʃV�[�g�𕡐�����v���O����()
Attribute �{�^���D���k�p�����ʃV�[�g�𕡐�����v���O����.VB_ProcData.VB_Invoke_Func = " \n14"


    Dim ���k�pws As Worksheet
    Set ���k�pws = Worksheets("���k�p������")

    Dim �\�� As a�\�쐬�p�ϐ��N���X
    Set �\�� = New a�\�쐬�p�ϐ��N���X
    Call �\��.�\�쐬�p�ϐ�������(���k�pws, "���k�pws")
    
    '�����V�[�g���x�쐬���邱�Ƃ�h��
    Dim ���O As String: ���O = ���k�pws.Range("B1") 'Worksheets("2�����ʁi�ҏW���ցj")���琶�k�����R�s�[�iVBA��ʂł̂ݕύX�\�j
    Dim �V�[�g���� As Worksheet
    
    If ���O = "" Then '�������k�������͂���Ă��Ȃ��ꍇ�A�G���[�R�[�h��\��
        MsgBox "���k������͂��Ă��������B"
        Exit Sub
    End If '�G���[���Ȃ���΁A�ȉ��̃v���O���������s
    
    For Each �V�[�g���� In Sheets                               '�V�[�g�̒����瓯�����k���̃V�[�g���Ȃ���For Each���[�v�ŒT��
        If �V�[�g����.Name = ���O Then
            MsgBox "�������k�̊�]�����\���x�쐬���Ă��܂��B" '����΃G���[���b�Z�[�W��\��
            Exit Sub
        End If
    Next �V�[�g���� '�G���[���Ȃ���΁A�ȉ��̃v���O���������s
   
    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��
   
    'Worksheets("2�����ʁi�ҏW���ցj")�𕡐�
    ���k�pws.Copy after:=���k�pws
    ActiveSheet.Name = "����"  '���������V�[�g�ɉ���������
    
    '�V�����\�̒���
    Dim �V�\�� As a�\�쐬�p�ϐ��N���X
    Set �V�\�� = New a�\�쐬�p�ϐ��N���X
    Call �V�\��.�\�쐬�p�ϐ�������(���k�pws, "����")
    
    Dim �\���� As b�폜�N���X
    Set �\���� = New b�폜�N���X
    Call �\����.�V�V�[�g�\���H�p(Worksheets("����"), �V�\��.�\�s�n, �V�\��.�\��n, �\��.�\�s�n, �V�\��.�\��I) '�V�V�[�g�̕\�̎n�_�̍s����A���V�[�g�̕\�̎n�_�̈��̍s�܂ō폜
                                                                                   '(���j���̕����͋��V�[�g�̕\�̎n�_�ɃZ�b�g���邱��
    Dim �y���� As g���C�A�E�g�Ə����N���X
    Set �y���� = New g���C�A�E�g�Ə����N���X
    Call �y����.�y�[�W���C�A�E�g�ƕ�������(���k�pws, Worksheets("����"), "����")

    Dim ������ As y�����t�����ݒ�
    Set ������ = New y�����t�����ݒ�
    Call ������.�����t�����ݒ�(���k�pws, Worksheets("����"), �V�\��.�\�s�n, �V�\��.�\��n, �V�\��.�\�s�I, �V�\��.�\��I)
    
    '�쐬�������
    Worksheets("����").Range("B2").Value = Date
    
    '���������V�[�g�̖��O�𐶓k���ɕύX
    Worksheets("����").Name = ���O
    
    '�������ł���悤��K5�Z���̔���}�[�J�[���폜
    ���k�pws.Range("K5").ClearContents
    
    '�t�@�C����ۑ�
    ActiveWorkbook.Save
    
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n
    
    
End Sub
