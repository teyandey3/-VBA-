Attribute VB_Name = "B���k�p�����ʕ\���H"
Option Explicit

Sub �{�^���A���k�p�����ʐV�K�V�[�g������\�ɉ��H����v���O����()


    Dim ���k�pws As Worksheet
    Set ���k�pws = Worksheets("���k�p������")

    Dim �\�� As a�\�쐬�p�ϐ��N���X
    Set �\�� = New a�\�쐬�p�ϐ��N���X
    Call �\��.�\�쐬�p�ϐ�������(���k�pws, "���k�pws")

    '�{�^���A�̓�x������h�����߂̃v���O����
    If ���k�pws.Range("K5").Value = "����}�[�J�[(�����Ȃ���)" Then 'K5�Z���iVBA��ʂł̂ݕύX�\�j�̕������擾���āA�����ԈႦ�ă{�^���A��A���œ�x��������A�G���[�R�[�h��\��
        MsgBox "�{�^���A��A���œ�x�����Ă��܂��B���̋@�\�͌��ݎg�����Ƃ��ł��܂���B" + vbCrLf + "�{�^���B���������A�ŏ������Ƃ���蒼���Ă��������B"
        Exit Sub
    End If '�G���[���Ȃ���΁A�ȉ��̃v���O���������s
    
    '�J�n������͂��������肷�邽�߂̃v���O����
    If ���k�pws.Cells(�\��.�\�s�n, �\��.�J�n����).Value = "�J�n��" Then '�����u�K�J�n����B15�̃Z���ɓ��͂��Ă��Ȃ�������G���[�R�[�h��\��
        MsgBox "�u�K�J�n������͂��Ă��������B"
        Exit Sub                                                   '�v���O�����̏I��
    End If
    
    Application.ScreenUpdating = False '��ʕ`����~
    Application.Cursor = xlWait '�E�G�C�g�J�[�\��
    Application.EnableEvents = False '�C�x���g��}�~
    Application.DisplayAlerts = False '�m�F���b�Z�[�W��}�~
    Application.Calculation = xlCalculationManual '�v�Z���蓮��
    
    '���t���������͂���N���X�̌Ăяo��
    Dim ���t As e���t�������̓N���X
    Set ���t = New e���t�������̓N���X
    Call ���t.���t��������(���k�pws, �\��.�\�s�n, �\��.�\��n, �\��.�\�s�I, �\��.�\��I, �\��.�R�}��, �\��.���o����, �\��.�J�n����)

    '�y���ɐF��t��������t��������ݒ肷��N���X(�v���V�[�W��)�̌Ăяo��
    Dim �y���� As f�y�������t�����N���X
    Set �y���� = New f�y�������t�����N���X
    Call �y����.�y�������t����(���k�pws)
    
    '�I��͈͂�����
    Dim �\�� As b�폜�N���X
    Set �\�� = New b�폜�N���X
    Call �\��.�\���Ӎ폜(���k�pws, �\��.�\�s�n, �\��.�\��n, �\��.�\�s�I, �\��.�\��I)
    
    '��L�̃v���O�����ŏ����Ă��܂����\�̉��O�g��t���Ȃ���
    ���k�pws.Range(Cells(�\��.�\�s�I, �\��.�\��n), Cells(�\��.�\�s�I, �\��.�\��I)).Borders(xlEdgeBottom).Weight = xlThick
    
    '�֐���}��
    ���k�pws.Range("E2").Formula = "=SUM(F2:O2)"
    ���k�pws.Range("E3").Formula = "=COUNTIF(17:1048576,""=0"")" '1048576�̓G�N�Z���̍ŏI�s�Ȃ̂ŕύX�s�v�B�i"15"���ŏ��̍s��VBA��ʂł̂ݕύX�\�j
    ���k�pws.Range("K4").Formula = "=IF(E2=E3, """", ""�G���[: �e���Ȃ̃R�}���̍��v�Ɗ�]�����\�ɓ��͂��ꂽ�R�}���̍��v����v���܂���B"")"
        
    '�{�^���A��A���œ�x�������Ƃ�h�����߂̃G���[���ʃ}�[�J�[��}��
    ���k�pws.Range("K5") = "����}�[�J�[(�����Ȃ���)" '�iVBA��ʂł̂ݕύX�\�j
        
    '���̑�����w�����郁�b�Z�[�W�̕\��
    MsgBox "�e���Ȃ̃R�}������͂��Ă��������B"
    
    Application.Calculation = xlCalculationAutomatic '�v�Z��������
    Application.DisplayAlerts = True '�m�F���b�Z�[�W���J�n
    Application.EnableEvents = True '�C�x���g���J�n
    Application.Cursor = xlDefault '�W���J�[�\��
    Application.ScreenUpdating = True '��ʕ`����J�n
    
    
End Sub
