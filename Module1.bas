Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public str As String            '�����񃏁[�N
Public nname As String          '�R���s���[�^�[��
Public ret                      '���b�Z�[�W�@���^�[��
Public strSQL As String         'sql�p
Public s_date As String         'sql�p
Public sets As Integer
Public XNO As Integer
Public fnx As Integer
Public wno As Long
Public Psetn  As Variant
Public Qsetn  As Variant
Public Tnyset  As Integer

Public sinn As Single
Public dobn As Double
Public dobn1 As Double
Public dobn2 As Double
Public dobn3 As Double
Public dobn4 As Double
Public intn As Integer
Public lct As Integer

Public gds1 As String
Public gds2 As Integer

Public M1W As Integer
Public N1W As Integer
Public P1W As Single

Public STNO As String
Public STEN As Integer
Public STDT As Integer
Public STAD As Double
Public STTD As Double
Public STYD As Double
Public STPH As Double
Public STPW As Double

Public DBB As Database
Public RECS As Recordset
Public Kubun_CD As String
Public myDB As Database
Public Myset As Recordset
Public Myset2 As Recordset
Public W_Single As Single


Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Function get_uriage(Stp As String, uriage As String, ST As String) As String
'���㕪�ށ@�@�@stp:���i���� uriage:���㕪��  ST:ST(��Е���)
    
get_uriage = uriage

'���R
If Stp = "S3" Then
    If ST = "S" Then
        get_uriage = "D00"
    Else
        get_uriage = "Y00"
    End If
End If

'�}�X
If Stp = "S4" Then
    If ST = "S" Then
        get_uriage = "E00"
    Else
        get_uriage = "Z00"
    End If
End If

End Function
Function get_ri(YN As Integer, YM As Integer, CD As String) As Integer
'���x�b�g���v�Z�@�@�@yn:YN ym:YM cd:���i�^�C�vCD
    Dim i As Integer
    Dim K As Integer
    Dim Nmas As Integer
    Dim Mmas As Integer
    Dim Bai, Itiretume As Integer

    i = YN * 2
    K = YM * 2


    
    If CD = "S3" Or CD = "S3A" Then
        get_ri = K - 2 '���i�q
        
    ElseIf CD = "S1A" Or CD = "S2A" Then '�V�۽�i�q
        Select Case YM
            Case 1 To 2
                get_ri = i + K
            Case 3
                get_ri = i + K + YN - 1
            Case 4
                get_ri = i + K + YN
            Case 5
                get_ri = (i + K + YN) - 1
            Case 6
                get_ri = (i + K + YN + YN)
            Case 7
                get_ri = (i + K + YN + YN)
            Case 8
                get_ri = (i + K + YN + YN + YN)
            Case 9
                get_ri = (i + K + YN + YN + YN) - 1
            Case 10
                get_ri = (i + K + YN + YN + YN + YN)
            Case 11
                get_ri = (i + K + YN + YN + YN + YN)
            Case 12
                get_ri = (i + K + YN + YN + YN + YN + YN)
            Case 13
                get_ri = (i + K + YN + YN + YN + YN + YN) - 1
            Case Else
        End Select
    
    ElseIf CD = "P3" Then '��Ÿر���V�L
        get_ri = YN * 3 + 4
    ElseIf CD = "P3A" Then '��Ÿر���V�L
        get_ri = YN * 3 + 4

    ElseIf CD = "P4" Then '��Ÿر���V����
        get_ri = YN * 2
    ElseIf CD = "P4A" Then '��Ÿر���V����
        get_ri = YN * 2

    ElseIf CD = "P5" Then   '�e�i�q
                
        Nmas = YN - 1 '�c�i�q�{��
        Mmas = YM - 1 '���i�q�{��(�㉺�[�����j
        
        get_ri = Nmas * 2 '�㉺�[�p���x�b�g
        
        Bai = kirishute(Nmas / 2, 0)
        Itiretume = kirishute(Mmas / 2, 0)

        Select Case Nmas Mod 2
            Case 0
                get_ri = get_ri + Mmas * Bai
            Case 1
                get_ri = get_ri + Itiretume + Mmas * Bai
        End Select
    
    Else '���۽�i�q
        Select Case YM
            Case 1 To 3
                get_ri = i + K
            Case 4
                get_ri = i + K + YN
            Case 5
                get_ri = (i + K + YN) - 1
            Case 6
                get_ri = (i + K + YN + YN)
            Case 7
                get_ri = (i + K + YN + YN)
            Case 8
                get_ri = (i + K + YN + YN + YN)
            Case 9
                get_ri = (i + K + YN + YN + YN) - 1
            Case 10
                get_ri = (i + K + YN + YN + YN + YN)
            Case 11
                get_ri = (i + K + YN + YN + YN + YN)
            Case 12
                get_ri = (i + K + YN + YN + YN + YN + YN)
            Case 13
                get_ri = (i + K + YN + YN + YN + YN + YN) - 1
            Case Else
        End Select
    End If

'    If CD = "S2" Or CD = "S2A" Or CD = "S5A" Then
'        get_ri = get_ri * 2
'    End If


End Function

Function get_ln2(Stp As String, Hno As String, tyu As String) As String
    If Stp Like "SH" Then
        get_ln2 = "C"
    ElseIf Stp Like "HS3" Then
        get_ln2 = "A"
    ElseIf Stp Like "HS*" Then
        get_ln2 = "C"
    ElseIf Stp Like "S*" Then
        get_ln2 = "A"
    ElseIf Stp = "RW" Then
        get_ln2 = "A"
    ElseIf Stp Like "P*" Then
        get_ln2 = "A"
    ElseIf Stp = "XX" Then
        get_ln2 = "A"
    ElseIf Stp Like "M*" Then
        get_ln2 = "C"
    ElseIf Stp Like "GP*" Then
        get_ln2 = "C"
    ElseIf Stp Like "HT4" Then
        If Hno = "PD11" Then
            get_ln2 = "A"
        ElseIf Hno = "PD21" Then
            get_ln2 = "C"
        End If
    ElseIf Stp Like "HT*" Then
        get_ln2 = "C"
    ElseIf Stp = "KK4" Then
        get_ln2 = "A"
    ElseIf Stp Like "KK*" Then
        get_ln2 = "C"
    ElseIf Stp Like "NS*" Then
        get_ln2 = "C"
    ElseIf Stp Like "FT*" Then
        get_ln2 = "B"    '2018/02/02 B���C���ɕύX
    End If
    
    If tyu = "F" Then
        get_ln2 = "F"
    End If

End Function

Function get_ln(Stp As String, rbn As Integer, rym As Integer, Hno As String, MH As Integer, Uri As String, tyu As String) As String
'Function get_ln(Stp As String, Hno As String, tyu As String) As String
'���C���v�Z�@�@�@stp:���i����  rbn:��ޯĐ�  rym:YM  Hno:�ϐ��O���[�v    MH:MH     URI:���㕪��
    Dim i As Integer
        
    If Stp Like "SH" Then
        get_ln = "C"
    ElseIf Stp Like "HS3" Then
        get_ln = "A"
    ElseIf Stp Like "HS*" Then
        get_ln = "C"
    ElseIf Stp Like "S*" Then
        get_ln = "A"
    ElseIf Stp = "RW" Then
        get_ln = "A"
    ElseIf Stp Like "P*" Then
        get_ln = "A"
    ElseIf Stp = "XX" Then
        get_ln = "A"
    ElseIf Stp Like "M*" Then
        get_ln = "C"
    ElseIf Stp Like "GP*" Then
        get_ln = "C"
    ElseIf Stp Like "HT4" Then
        If Hno = "PD11" Then
            get_ln = "A"
        ElseIf Hno = "PD21" Then
            get_ln = "C"
        End If
    ElseIf Stp Like "HT*" Then
        get_ln = "C"
    ElseIf Stp = "KK4" Then
        get_ln = "A"
    ElseIf Stp Like "KK*" Then
        get_ln = "C"
    ElseIf Stp Like "NS*" Then
        get_ln = "C"
    ElseIf Stp Like "FT*" Then
'        get_ln = "C"
        get_ln = "B"    '2018/02/02 B���C���ɕύX
    End If
    
    If tyu = "F" Then
        get_ln = "F"
    End If

End Function

Function get_bcd(kbn As Variant, cbn As Variant, sck As Variant, did As Variant) As Variant
    '�ް���ޕҏW kbn:��Ћ敪,cbn:�����ԍ�,sck:�o�בq��,did:�ް�ID
    If did = "A0" Then
        '��
        If kbn = "T" Then
            '���R
            get_bcd = "ZZZ" & Mid(cbn, 1, 10) & "  "
        Else
            '�O��
            get_bcd = "ZZZ" & cbn
        End If
    Else
        '�K�i
        If Mid(sck, 1, 1) = "T" Then
            '���R
            get_bcd = "ZZZYB" & Mid(cbn, 1, 8) & "  "
        Else
            '�O��
            get_bcd = "ZZZ" & Mid(cbn, 1, 10) & "    "
        End If
    End If

End Function

Public Function GetMyComputerName() As String
    Dim strCmptrNameBuff As String * 16
    GetComputerName strCmptrNameBuff, Len(strCmptrNameBuff)
    GetMyComputerName = Left$(strCmptrNameBuff, InStr(strCmptrNameBuff, vbNullChar) - 1)
End Function

Function get_cno() As String
    get_cno = gds1
End Function

Function get_sno() As Integer
    get_sno = gds2
End Function
Function str_day() As String
    Dim dw As String
    dw = Format(Date, "yy/mm/dd")
    str_day = dw
End Function

Function get_NCNo(SType As String) As String
    If Mid(SType, 2, 1) = "2" Then
        get_NCNo = "99,100"
    ElseIf Mid(SType, 2, 1) = "3" Then
        If Mid(SType, 4, 1) = "0" Then
            get_NCNo = "114"
        ElseIf Mid(SType, 4, 1) = "1" Then
            get_NCNo = "116"
        ElseIf Mid(SType, 4, 1) = "2" Then
            get_NCNo = "117"
        End If
    End If
End Function

Function shisya(dbl�Ώ� As Double, lng���� As Long) As Double
    '�l�̌ܓ�
    '��̈ʂ�-1�A�\�̈ʂ�-2�A�S�̈ʂ�-3�A
    '�����_��P�ʂ�0�A��2�ʂ�1�A��3�ʂ�2�A
    Dim lng���l As Long
    lng���l = 10 ^ Abs(lng����)
    
    If lng���� > 0 Then
        shisya = Int(dbl�Ώ� * lng���l + 0.555555) / lng���l
    Else
        shisya = Int(dbl�Ώ� / lng���l + 0.555555) * lng���l
    End If
End Function

Function shisya2(dbl�Ώ� As Double, lng���� As Long) As Double
    '�l�̌ܓ�
    '��̈ʂ�-1�A�\�̈ʂ�-2�A�S�̈ʂ�-3�A
    '�����_��P�ʂ�0�A��2�ʂ�1�A��3�ʂ�2�A
    Dim lng���l As Long
    lng���l = 10 ^ Abs(lng����)
    
    If lng���� > 0 Then
        shisya2 = Int(dbl�Ώ� * lng���l + 0.5) / lng���l
    Else
        shisya2 = Int(dbl�Ώ� / lng���l + 0.5) * lng���l
    End If
End Function

Function kirishute(dbl�Ώ� As Double, lng���� As Long) As Double
    '�؂�̂�
    '��̈ʂ�-1�A�\�̈ʂ�-2�A�S�̈ʂ�-3�A
    '�����_��P�ʂ�0�A��2�ʂ�1�A��3�ʂ�2�A

    Dim lng���l As Long
    lng���l = 10 ^ Abs(lng����)

    If lng���� > 0 Then
        kirishute = Int(dbl�Ώ� * lng���l) / lng���l
    Else
        kirishute = Int(dbl�Ώ� / lng���l) * lng���l
    End If
End Function

Function kiriage(dbl�Ώ� As Double, lng���� As Long) As Double
    '�؂�グ
    '��̈ʂ�-1�A�\�̈ʂ�-2�A�S�̈ʂ�-3�A
    '�����_��P�ʂ�0�A��2�ʂ�1�A��3�ʂ�2�A

    Dim lng���l As Long
    lng���l = 10 ^ Abs(lng����)

    If lng���� > 0 Then
        
        If Int(dbl�Ώ� * lng���l) = (dbl�Ώ� * lng���l) Then
            kiriage = Int(dbl�Ώ� * lng���l) / lng���l
        Else
            kiriage = (Int(dbl�Ώ� * lng���l) + 1) / lng���l
        End If
    Else
        If Int(dbl�Ώ� / lng���l) = (dbl�Ώ� / lng���l) Then
            kiriage = Int(dbl�Ώ� / lng���l) * lng���l
        Else
            kiriage = (Int(dbl�Ώ� / lng���l) + 1) * lng���l
        End If
    End If
End Function

Function Odr_Add()
    '�����f�[�^�̒ǉ� �󒍃f�[�^�̑��݃`�F�b�N
    Dim DB As Database
    Dim rst As Recordset
    Dim rsd As Recordset

    Dim trc As Long
    Dim sds As String

    trc = DCount("[�����ԍ�]", "SAN")
    If trc = 0 Then
        Exit Function
    End If

    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("SAN�ݐ�", dbOpenDynaset)
    Set rsd = DB.OpenRecordset("SAN")

    rsd.MoveFirst
    Do Until rsd.EOF

        sds = rsd![�����ԍ�]
        rst.FindFirst "[�����ԍ�]='" & sds & "'"

        ' NoMatch �v���p�e�B�̒l�Ɋ�Â��Ė߂�l��ݒ肵�܂��B
        If rst.NoMatch = False Then
            rsd.Delete
        End If
        rsd.MoveNext
    Loop
    rsd.Close: Set rsd = Nothing
    rst.Close: Set rst = Nothing


    trc = DCount("[�����ԍ�]", "q_�p�i����")
    If trc = 0 Then
        Exit Function
    End If
    Set rst = DB.OpenRecordset("q_�p�i����", dbOpenDynaset)
    rst.MoveFirst
    Do Until rst.EOF
        rst.Edit
        '���@�l1
        If rst![���@�L��1] = "MW" And rst![��{�^�C�v�R�[�h] Like "M092D01*" Then
            rst![���@�l1] = rst![���@�l1] - 480
            rst![���@�L��1] = "W"
        ElseIf rst![���@�L��1] = "MW" And rst![��{�^�C�v�R�[�h] Like "M092D02*" Then
            rst![���@�l1] = rst![���@�l1] - 920
            rst![���@�L��1] = "W"
        ElseIf rst![���@�L��1] = "MW" Then
            rst![���@�l1] = rst![���@�l1] - 820
            rst![���@�L��1] = "W"
        ElseIf rst![���@�L��1] = "W" Then
            rst![���@�l1] = rst![���@�l1]

        ElseIf rst![���@�L��1] = "MH" And rst![��{�^�C�v�R�[�h] Like "M092D01*" Then
            rst![���@�l1] = rst![���@�l1] - 160
            rst![���@�L��1] = "H"
        ElseIf rst![���@�L��1] = "MH" And rst![��{�^�C�v�R�[�h] Like "M092D02*" Then
            rst![���@�l1] = rst![���@�l1] - 330
            rst![���@�L��1] = "H"
        ElseIf rst![���@�L��1] = "MH" Then
            rst![���@�l1] = rst![���@�l1] - 600
            rst![���@�L��1] = "H"
        ElseIf rst![���@�L��1] = "H" Then
            rst![���@�l1] = rst![���@�l1]
        Else
            rst![���@�l1] = 0
        End If

        '���@�l2
        If rst![���@�L��2] = "MW" And rst![��{�^�C�v�R�[�h] Like "M092D01*" Then
            rst![���@�l2] = rst![���@�l2] - 480
            rst![���@�L��2] = "W"
        ElseIf rst![���@�L��2] = "MW" And rst![��{�^�C�v�R�[�h] Like "M092D02*" Then
            rst![���@�l2] = rst![���@�l2] - 920
            rst![���@�L��2] = "W"
        ElseIf rst![���@�L��2] = "MW" Then
            rst![���@�l2] = rst![���@�l2] - 820
            rst![���@�L��2] = "W"
        ElseIf rst![���@�L��2] = "W" Then
            rst![���@�l2] = rst![���@�l2]

        ElseIf rst![���@�L��2] = "MH" And rst![��{�^�C�v�R�[�h] Like "M092D01*" Then
            rst![���@�l2] = rst![���@�l2] - 160
            rst![���@�L��2] = "H"
        ElseIf rst![���@�L��2] = "MH" And rst![��{�^�C�v�R�[�h] Like "M092D02*" Then
            rst![���@�l2] = rst![���@�l2] - 330
            rst![���@�L��2] = "H"
        ElseIf rst![���@�L��2] = "MH" Then
            rst![���@�l2] = rst![���@�l2] - 600
            rst![���@�L��2] = "H"
        ElseIf rst![���@�L��2] = "H" Then
            rst![���@�l2] = rst![���@�l2]
        Else
            rst![���@�l2] = 0
        End If

        rst.Update
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    DB.Close: Set DB = Nothing
End Function
Function Odr7_Add()
    
    '�����f�[�^�̒ǉ� �󒍃f�[�^�̑��݃`�F�b�N
    Dim DB As Database
    Dim rst As Recordset
    Dim rsd As Recordset
    
    Dim trc As Long
    Dim sds As String
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("SAN�ݐ�", dbOpenDynaset)
    Set rsd = DB.OpenRecordset("SANA7")
    
    trc = DCount("[�����ԍ�]", "SANA7")
    If trc = 0 Then
        GoTo nods
    End If
    
    rsd.MoveFirst
    Do Until rsd.EOF
                    
        sds = rsd![�����ԍ�]
        rst.FindFirst "[�����ԍ�]='" & sds & "'"
    
        ' NoMatch �v���p�e�B�̒l�Ɋ�Â��Ė߂�l��ݒ肵�܂��B
        If rst.NoMatch = False Then
            rsd.Delete
        End If
        rsd.MoveNext
    Loop
    
nods:
    rst.Close: Set rst = Nothing
    rsd.Close: Set rsd = Nothing
    DB.Close: Set DB = Nothing
    
End Function
Function Ka_Add() '�ǉ��p�@�i2013/1/29���݁@�g�p���Ă��Ȃ��j
    
    '����΂�f�[�^�쐬
    Dim DB As Database
    Dim rst As Recordset
    Dim rso As Recordset
    Dim rSM As Recordset
    Dim rsn As Recordset
    Dim rsk As Recordset
    Dim trc As Long
    Dim sds As String
    Dim hds As String
    Dim ksz As String
    Dim chk As Boolean
    Dim kigou(7) As Variant
    Dim sunpou(7) As Long
    
    'DoCmd.OpenQuery "d_YOTEIܰ�"
    'DoCmd.OpenQuery "d_�i�q��ܰ�"
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("j_YOTEI�ǉ�")
    Set rSM = DB.OpenRecordset("m_�ϐ�", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ܰ�no")
    Set rso = DB.OpenRecordset("YOTEIܰ�")
    Set rsk = DB.OpenRecordset("m_�H��}�ԍ�", dbOpenDynaset)
        
    Set DBB = CurrentDb
    Set RECS = DBB.OpenRecordset("�i�q��ܰ�")

    trc = DCount("[�Г��Ǘ��ԍ�1]", "j_YOTEI�ǉ�")
    If trc = 0 Then
        GoTo nodu
    End If
    
    rst.MoveFirst
    Do Until rst.EOF
            
            rst.Edit
                 
            rsn.MoveFirst
            wno = rsn![�̔�no]
            wno = wno + 1
            rsn.Edit
            rsn![�̔�no] = wno
            rsn.Update
        
            rso.AddNew
            rso![�����ԍ�] = rst![�Г��Ǘ��ԍ�1] & rst![�Г��Ǘ��ԍ�2]
            rso![�����}��] = 0
            rso![�̔�no] = wno
            rso![DT�敪] = "0"
            rso![�����敪] = ""
            rso![�󒍋敪] = ""
            rso![�Г��Ǘ��ԍ�1] = rst![�Г��Ǘ��ԍ�1]
            rso![�Г��Ǘ��ԍ�2] = rst![�Г��Ǘ��ԍ�2]
            rso![�[����] = rst![�[����]
            rso![�����\�辯Đ�] = rst![�����\�辯Đ�]
           
            If rst![T�g�p] = False Then
                rso![H���@�l] = rst![TH]
                rso![W���@�l] = rst![TW]
            Else
                rso![H���@�l] = rst![H��]
                rso![W���@�l] = rst![W��]
            End If
            
            rso![���i����] = rst![���i����]
            rso![COL] = rst![�F]
            rso![���i����CD] = rst![���i����CD]
            rso![�ϐ���ٰ��] = rst![�ϐ���ٰ��]
            rso![��g�H��}�ԍ�] = rst![��g�H��}�ԍ�]
            ksz = rst![��g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![��g�^��] = rsk![�^��]
            
            rso![���g�H��}�ԍ�] = rst![���g�H��}�ԍ�]
            ksz = rst![���g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![���g�^��] = rsk![�^��]
            
            rso![�G�g�H��}�ԍ�] = rst![�G�g�H��}�ԍ�]
            ksz = rst![�G�g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�G�g�^��] = rsk![�^��]
            
            rso![�i�q���H��}�ԍ�] = rst![�i�q���H��}�ԍ�]
            ksz = rst![�i�q���H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�i�q���^��] = rsk![�^��]
            
            hds = rst![�ϐ���ٰ��]
            rSM.FindFirst "[�ϐ���ٰ��]='" & hds & "'"
            
            If rSM![MW����] = True Then
                rso![YMW] = (rso![W���@�l] / 2) + rSM![MW�p�ϐ�]
            Else
                rso![YMW] = rso![W���@�l] + rSM![MW�p�ϐ�]
            End If
            
            rso![YMH] = rso![H���@�l] + rSM![MH�p�ϐ�]
            
            
            rso![�㉺�g�ؒf���@] = rso![YMW] + rSM![��g���g���@�p]
            rso![�G�g�ؒf���@] = rso![YMH] + rSM![�G�g���@�p]
            
            If rst![���i����CD] = "S3" Then
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
            End If
            If rst![���i����CD] = "S4" Then
                rso![�i�q�G�ؒf���@] = rso![YMH] + rSM![�i�q�G���@�p]
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
                STTD = rso![�i�q�G�ؒf���@]
            End If
            
            If rst![���i����CD] = "XX" Then
                rso![�i�q�G�ؒf���@] = rso![YMH] - 50.5
                STYD = rso![�i�q�G�ؒf���@]
            End If
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                If rst![���i����CD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![���i����CD] <> "S3" Then
               dobn1 = rso![YPW] * rso![YPW]
               dobn2 = rso![YPH] * rso![YPH]
               dobn3 = dobn1 + dobn2
               dobn4 = Sqr(dobn3)
               dobn = dobn4 / 2
               rso![YP1] = dobn
               P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![��g���g�i�q��t�߯�]
            If rst![���i����CD] = "S3" Or rst![���i����CD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![�G�g�i�q��t�߯�]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![�G�g�i�q��t�߯�]
            End If
            
            STNO = rso![�����ԍ�]
            STEN = rso![�����}��]
            STDT = rso![DT�敪]
            STAD = rSM![�i�qAD�l]
            
            If rst![���i����CD] = "S2" Then
                rso![FS�敪] = True
            End If
            
            If rst![���i����CD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If

            If rst![���i����CD] = "XX" Then
                If rso![�㉺�g�ؒf���@] = 1193 Then
                    rso![YN] = 9
                    rso![YPW] = 117
                    N1W = rso![YN]
                Else
                    rso![YN] = 14
                    rso![YPW] = 117
                    N1W = rso![YN]
                End If
            End If
            
            rso![�ǉ��׸�] = True
 
            rso![���x�b�g] = get_ri(rso![YN], rso![YM], rso![���i����CD])
            
            rso.Update
           
            '�i�q���f�[�^�̏o��
            Select Case rst![���i����CD]
                Case "S1"
                    Call S1_set     '�N���X
                Case "S2"
                    Call S1_set     '�N���X�@����
                Case "S3"
                    Call S3_set     '��
                Case "S4"
                    Call S4_set     '�e
                Case "XX"
                    Call XX_set     '��t���쏊����
            End Select
            
            If rst![���i����CD] = "S2" Then
        
            rst.Edit
                 
            rsn.MoveFirst
            wno = rsn![�̔�no]
            wno = wno + 1
            rsn.Edit
            rsn![�̔�no] = wno
            rsn.Update
        
            rso.AddNew
            rso![�����ԍ�] = rst![�Г��Ǘ��ԍ�1] & rst![�Г��Ǘ��ԍ�2]
            rso![�����}��] = 0
            rso![�̔�no] = wno
            rso![DT�敪] = "0"
            rso![�����敪] = ""
            rso![�󒍋敪] = ""
            rso![�Г��Ǘ��ԍ�1] = rst![�Г��Ǘ��ԍ�1]
            rso![�Г��Ǘ��ԍ�2] = rst![�Г��Ǘ��ԍ�2]
            rso![�[����] = rst![�[����]
            rso![�����\�辯Đ�] = rst![�����\�辯Đ�]
           
            If rst![T�g�p] = False Then
                rso![H���@�l] = rst![TH2]
                rso![W���@�l] = rst![TW2]
            Else
                rso![H���@�l] = rst![H��2]
                rso![W���@�l] = rst![W��2]
            End If
            
            rso![���i����] = rst![���i����]
            rso![COL] = rst![�F]
            rso![���i����CD] = rst![���i����CD]
            rso![�ϐ���ٰ��] = rst![�ϐ���ٰ��]
            rso![��g�H��}�ԍ�] = rst![��g�H��}�ԍ�]
            ksz = rst![��g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![��g�^��] = rsk![�^��]
            
            rso![���g�H��}�ԍ�] = rst![���g�H��}�ԍ�]
            ksz = rst![���g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![���g�^��] = rsk![�^��]
            
            rso![�G�g�H��}�ԍ�] = rst![�G�g�H��}�ԍ�]
            ksz = rst![�G�g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�G�g�^��] = rsk![�^��]
            
            rso![�i�q���H��}�ԍ�] = rst![�i�q���H��}�ԍ�]
            ksz = rst![�i�q���H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�i�q���^��] = rsk![�^��]
            
            hds = rst![�ϐ���ٰ��2]
            rSM.FindFirst "[�ϐ���ٰ��]='" & hds & "'"
            
            If rSM![MW����] = True Then
                rso![YMW] = (rso![W���@�l] / 2) + rSM![MW�p�ϐ�]
            Else
                rso![YMW] = rso![W���@�l] + rSM![MW�p�ϐ�]
            End If
            
            rso![YMH] = rso![H���@�l] + rSM![MH�p�ϐ�]
            
            rso![�㉺�g�ؒf���@] = rso![YMW] + rSM![��g���g���@�p]
            rso![�G�g�ؒf���@] = rso![YMH] + rSM![�G�g���@�p]

            
            If rst![���i����CD] = "S3" Then
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
            End If
            If rst![���i����CD] = "S4" Then
                rso![�i�q�G�ؒf���@] = rso![YMH] + rSM![�i�q�G���@�p]
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
                STTD = rso![�i�q�G�ؒf���@]
            End If
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                If rst![���i����CD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![���i����CD] <> "S3" Then
                dobn1 = rso![YPW] * rso![YPW]
                dobn2 = rso![YPH] * rso![YPH]
                dobn3 = dobn1 + dobn2
                dobn4 = Sqr(dobn3)
                dobn = dobn4 / 2
                rso![YP1] = dobn
                P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![��g���g�i�q��t�߯�]
            If rst![���i����CD] = "S3" Or rst![���i����CD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![�G�g�i�q��t�߯�]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![�G�g�i�q��t�߯�]
            End If
            
            STNO = rso![�����ԍ�]
            STEN = rso![�����}��]
            STDT = rso![DT�敪]
            STAD = rSM![�i�qAD�l]
            
            If rst![���i����CD] = "S2" Then
                rso![FS�敪] = False
            End If
            
            
            If rst![���i����CD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If
            
            rso![�ǉ��׸�] = True
 
            rso.Update
           
            Call S1_set     '�N���X�@����
            
            End If
            
'            rst.Edit
'            rst![�쐬�敪] = True
'            rst.Update
            rst.MoveNext
    Loop
    
nodu:
    rso.Close
    rSM.Close
    rsn.Close
    rst.Close
    rsk.Close
    RECS.Close
    
End Function
Function GP_Add()
    '����΂�f�[�^�쐬
    Dim DB As Database
    Dim rst As Recordset
    Dim rsd As Recordset
    Dim rso As Recordset
    Dim rSM As Recordset
    Dim rsn As Recordset
    Dim rsk As Recordset
    Dim trc As Long
    Dim sds As String
    Dim hds As String
    Dim ksz As String
    Dim chk As Boolean
    Dim kigou(7) As Variant
    Dim sunpou(7) As Long
    Dim kot As Variant
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("j_YOTEI_GP")
    Set rsd = DB.OpenRecordset("q_SAN�ݐ�", dbOpenDynaset)
    Set rSM = DB.OpenRecordset("m_�ϐ�", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ܰ�no")
    Set rso = DB.OpenRecordset("YOTEIܰ�")
    Set rsk = DB.OpenRecordset("m_�H��}�ԍ�", dbOpenDynaset)
        
    Set DBB = CurrentDb
    trc = DCount("[�����ԍ�]", "j_YOTEI_GP")
    If trc = 0 Then
        GoTo nodt
    End If

    rst.MoveFirst
    Do Until rst.EOF

        sds = rst![�����ԍ�]
        rsd.FindFirst "[�����ԍ�]='" & sds & "'"
    
        ' NoMatch �v���p�e�B�̒l�Ɋ�Â��Ė߂�l��ݒ肵�܂��B
        If rsd.NoMatch = False Then
        
            rsn.MoveFirst
            wno = rsn![�̔�no]
            wno = wno + 1
            rsn.Edit
            rsn![�̔�no] = wno
            rsn.Update
        
            rso.AddNew
            rso![�����ԍ�] = rst![�����ԍ�]
            rso![�����}��] = rst![�����}��]
            rso![COL] = rst![�F]
            
            rso![�̔�no] = wno
            rso![DT�敪] = "0"
            rso![�����敪] = rst![�����敪]
            rso![�󒍋敪] = rst![�󒍋敪]
            
            rso![�Г��Ǘ��ԍ�1] = rst![�Г��Ǘ��ԍ�1]
            kot = "0000" + rst![�Г��Ǘ��ԍ�2]
            rso![�Г��Ǘ��ԍ�2] = Right(kot, 4)
            rso![�[����] = rst![�[����]
            rso![�����\�辯Đ�] = rst![�����\�辯Đ�]
            rso![����] = rst![����]
           
            kigou(0) = rsd![���@�L��1]
            kigou(1) = rsd![���@�L��2]
            kigou(2) = rsd![���@�L��3]
            kigou(3) = rsd![���@�L��4]
            kigou(4) = rsd![���@�L��5]
            kigou(5) = rsd![���@�L��6]
            kigou(6) = rsd![���@�L��7]
            kigou(7) = rsd![���@�L��8]
            sunpou(0) = rsd![���@�l1]
            sunpou(1) = rsd![���@�l2]
            sunpou(2) = rsd![���@�l3]
            sunpou(3) = rsd![���@�l4]
            sunpou(4) = rsd![���@�l5]
            sunpou(5) = rsd![���@�l6]
            sunpou(6) = rsd![���@�l7]
            sunpou(7) = rsd![���@�l8]
            
            If rst![����ϰ�] = "1" Then
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H2" Then
                        rso![H���@�l] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
                
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W2" Then
                        rso![W���@�l] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
            Else
                rso![H���@�l] = rst![H��]
                rso![W���@�l] = rst![W��]
            End If
                
            
            rso![���i����] = rst![���i����]
            rso![���i����CD] = rst![���i����CD]
            rso![�ϐ���ٰ��] = rst![�ϐ���ٰ��]
            
            rso![��g�H��}�ԍ�] = rst![��g�H��}�ԍ�]
            ksz = rst![��g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![��g�^��] = rsk![�^��]
            
            rso![���g�H��}�ԍ�] = rst![���g�H��}�ԍ�]
            ksz = rst![���g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![���g�^��] = rsk![�^��]
            
            rso![�G�g�H��}�ԍ�] = rst![�G�g�H��}�ԍ�]
            ksz = rst![�G�g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�G�g�^��] = rsk![�^��]
            
            Select Case rst![���i����CD]
             Case "GP4"
                If rso![H���@�l] >= 2000 And rso![H���@�l] < 2245 Then
                    rso![YN] = 8
                ElseIf rso![H���@�l] >= 2245 And rso![H���@�l] < 2945 Then
                    rso![YN] = 10
                ElseIf rso![H���@�l] >= 2945 And rso![H���@�l] < 3645 Then
                    rso![YN] = 12
                ElseIf rso![H���@�l] >= 3645 And rso![H���@�l] < 3800 Then
                    rso![YN] = 14
                End If
                rso![YPH] = 700
             Case "GP7"
                rso![YN] = 4
                rso![YPH] = 700
            End Select
            
            rso.Update
        
'            rst.Edit
'            rst![�ް��쐬�敪] = True
'            rst.Update
        End If
        rst.MoveNext
    Loop
    
nodt:
    rso.Close
    rsd.Close
    rSM.Close
    rsn.Close
    rst.Close
    rsk.Close
End Function

Function Kb_Add()

    '����΂�f�[�^�쐬
    Dim DB As Database
    Dim rst As Recordset
    Dim rsd As Recordset
    Dim rso As Recordset
    Dim rSM As Recordset
    Dim rsn As Recordset
    Dim rsk As Recordset
    Dim trc As Long
    Dim sds As String
    Dim hds As String
    Dim ksz As String
    Dim chk As Boolean
    Dim kigou(7) As Variant
    Dim sunpou(7) As Long
    Dim kot As Variant
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("j_YOTEI")
    Set rsd = DB.OpenRecordset("q_SAN�ݐ�", dbOpenDynaset)
    Set rSM = DB.OpenRecordset("m_�ϐ�", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ܰ�no")
    Set rso = DB.OpenRecordset("YOTEIܰ�")
    Set rsk = DB.OpenRecordset("m_�H��}�ԍ�", dbOpenDynaset)
        
    Set DBB = CurrentDb
    Set RECS = DBB.OpenRecordset("�i�q��ܰ�")

    trc = DCount("[�����ԍ�]", "j_YOTEI")
    If trc = 0 Then
        GoTo nodt
    End If

    rst.MoveFirst
    Do Until rst.EOF

        sds = rst![�����ԍ�]
        rsd.FindFirst "[�����ԍ�]='" & sds & "'"
    
        ' NoMatch �v���p�e�B�̒l�Ɋ�Â��Ė߂�l��ݒ肵�܂��B
        If rsd.NoMatch = False Then
        
            rsn.MoveFirst
            wno = rsn![�̔�no]
            wno = wno + 1
            rsn.Edit
            rsn![�̔�no] = wno
            rsn.Update
        
            rso.AddNew
            rso![�����ԍ�] = rst![�����ԍ�]
            rso![�����}��] = rst![�����}��]
            
            If rst![���i�R�[�h] = "Sangutte" Then
                rso![COL] = rst![�����F]
            Else
                rso![COL] = rst![�F]
            End If
            
            rso![�̔�no] = wno
            rso![DT�敪] = "0"
            rso![�����敪] = rst![�����敪]
            rso![�󒍋敪] = rst![�󒍋敪]
            
            rso![�Г��Ǘ��ԍ�1] = rst![�Г��Ǘ��ԍ�1]
            kot = "0000" + rst![�Г��Ǘ��ԍ�2]
            rso![�Г��Ǘ��ԍ�2] = Right(kot, 4)
            
            rso![�[����] = rst![�[����]
            rso![�����\�辯Đ�] = rst![�����\�辯Đ�]
            rso![����] = rst![����]
           
            kigou(0) = rsd![���@�L��1]
            kigou(1) = rsd![���@�L��2]
            kigou(2) = rsd![���@�L��3]
            kigou(3) = rsd![���@�L��4]
            kigou(4) = rsd![���@�L��5]
            kigou(5) = rsd![���@�L��6]
            kigou(6) = rsd![���@�L��7]
            kigou(7) = rsd![���@�L��8]
            sunpou(0) = rsd![���@�l1]
            sunpou(1) = rsd![���@�l2]
            sunpou(2) = rsd![���@�l3]
            sunpou(3) = rsd![���@�l4]
            sunpou(4) = rsd![���@�l5]
            sunpou(5) = rsd![���@�l6]
            sunpou(6) = rsd![���@�l7]
            sunpou(7) = rsd![���@�l8]
            
            If rst![����ϰ�] = "1" Then
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H2" Then
                        rso![H���@�l] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
                
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W2" Then
                        rso![W���@�l] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
            Else
                rso![H���@�l] = rst![H��]
                rso![W���@�l] = rst![W��]
            End If
            
            rso![���i����] = rst![���i����]
            rso![���i����CD] = rst![���i����CD]
            rso![�ϐ���ٰ��] = rst![�ϐ���ٰ��]
            rso![��g�H��}�ԍ�] = rst![��g�H��}�ԍ�]
            ksz = rst![��g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![��g�^��] = rsk![�^��]
            
            rso![���g�H��}�ԍ�] = rst![���g�H��}�ԍ�]
            ksz = rst![���g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![���g�^��] = rsk![�^��]
            
            rso![�G�g�H��}�ԍ�] = rst![�G�g�H��}�ԍ�]
            ksz = rst![�G�g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�G�g�^��] = rsk![�^��]
            
            rso![�i�q���H��}�ԍ�] = rst![�i�q���H��}�ԍ�]
            ksz = rst![�i�q���H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�i�q���^��] = rsk![�^��]
            
            hds = rst![�ϐ���ٰ��]
            rSM.FindFirst "[�ϐ���ٰ��]='" & hds & "'"
            
            If rSM![MW����] = True Then
                rso![YMW] = (rso![W���@�l] / 2) + rSM![MW�p�ϐ�]
            Else
                rso![YMW] = rso![W���@�l] + rSM![MW�p�ϐ�]
            End If
            
            rso![YMH] = rso![H���@�l] + rSM![MH�p�ϐ�]
            
            rso![�㉺�g�ؒf���@] = rso![YMW] + rSM![��g���g���@�p]
            rso![�G�g�ؒf���@] = rso![YMH] + rSM![�G�g���@�p]

            
            If rst![���i����CD] = "S3" Then
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
            End If
            If rst![���i����CD] = "S4" Then
                rso![�i�q�G�ؒf���@] = rso![YMH] + rSM![�i�q�G���@�p]
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
                STTD = rso![�i�q�G�ؒf���@]
            End If
            If rst![���i����CD] = "XX" Then
                rso![�i�q�G�ؒf���@] = rso![YMH] - 50.5
                STYD = rso![�i�q�G�ؒf���@]
            End If
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                If rst![���i����CD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![���i����CD] <> "S3" Then
                dobn1 = rso![YPW] * rso![YPW]
                dobn2 = rso![YPH] * rso![YPH]
                dobn3 = dobn1 + dobn2
                dobn4 = Sqr(dobn3)
                dobn = dobn4 / 2
                rso![YP1] = dobn
                P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![��g���g�i�q��t�߯�]
            If rst![���i����CD] = "S3" Or rst![���i����CD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![�G�g�i�q��t�߯�]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![�G�g�i�q��t�߯�]
            End If
            
            If rst![���i����CD] = "S3" Then
                rso![FS�敪] = True
            End If
            
            STNO = rso![�����ԍ�]
            STEN = rso![�����}��]
            STDT = rso![DT�敪]
            STAD = rSM![�i�qAD�l]
            
            If rst![���i����CD] = "S2" Then
                rso![FS�敪] = True
            End If
            
            If rst![���i����CD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If
 
 
            If rst![���i����CD] = "XX" Then
                If rso![�㉺�g�ؒf���@] = 1193 Then
                    rso![YN] = 9
                    rso![YPW] = 117
                    N1W = rso![YN]
                Else
                    rso![YN] = 14
                    rso![YPW] = 117
                    N1W = rso![YN]
                End If
            End If
            
            rso![���x�b�g] = get_ri(rso![YN], rso![YM], rso![���i����CD]) * rst![��Đ�]
            
            rso.Update
           
            '�i�q���f�[�^�̏o��
            Select Case rst![���i����CD]
                Case "S1"
                    Call S1_set     '�N���X
                Case "S2"
                    Call S1_set     '�N���X�@����
                Case "S3"
                    Call S3_set     '��
                Case "S4"
                    Call S4_set     '�e
                Case "XX"
                    Call XX_set     '��t���쏊����
            End Select
            
'�N���X���葋�X�^�[�g
            If rst![���i����CD] = "S2" Then
            
            rsn.MoveFirst
            wno = rsn![�̔�no]
            wno = wno + 1
            rsn.Edit
            rsn![�̔�no] = wno
            rsn.Update
        
            rso.AddNew
            rso![�����ԍ�] = rst![�����ԍ�]
            rso![�����}��] = rst![�����}��]
            rso![COL] = rsd![�F]
            rso![�̔�no] = wno
            rso![DT�敪] = "0"
            rso![�����敪] = rst![�����敪]
            rso![�󒍋敪] = rst![�󒍋敪]
            rso![�Г��Ǘ��ԍ�1] = rst![�Г��Ǘ��ԍ�1]
            rso![�Г��Ǘ��ԍ�2] = rst![�Г��Ǘ��ԍ�2]
            rso![�[����] = rst![�[����]
            rso![�����\�辯Đ�] = rst![�����\�辯Đ�]
           
            kigou(0) = rsd![���@�L��1]
            kigou(1) = rsd![���@�L��2]
            kigou(2) = rsd![���@�L��3]
            kigou(3) = rsd![���@�L��4]
            kigou(4) = rsd![���@�L��5]
            kigou(5) = rsd![���@�L��6]
            kigou(6) = rsd![���@�L��7]
            kigou(7) = rsd![���@�L��8]
            sunpou(0) = rsd![���@�l1]
            sunpou(1) = rsd![���@�l2]
            sunpou(2) = rsd![���@�l3]
            sunpou(3) = rsd![���@�l4]
            sunpou(4) = rsd![���@�l5]
            sunpou(5) = rsd![���@�l6]
            sunpou(6) = rsd![���@�l7]
            sunpou(7) = rsd![���@�l8]
            
            If rst![����ϰ�] = "1" Then
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H1" Then
                        rso![H���@�l] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
                        
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W1" Then
                        rso![W���@�l] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
            Else
                rso![H���@�l] = rst![H��2]
                rso![W���@�l] = rst![W��2]
            End If
            
            rso![���i����] = rst![���i����]
            rso![���i����CD] = rst![���i����CD]
            rso![�ϐ���ٰ��] = rst![�ϐ���ٰ��2]
            rso![��g�H��}�ԍ�] = rst![��g�H��}�ԍ�]
            ksz = rst![��g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![��g�^��] = rsk![�^��]
            
            rso![���g�H��}�ԍ�] = rst![���g�H��}�ԍ�]
            ksz = rst![���g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![���g�^��] = rsk![�^��]
            
            rso![�G�g�H��}�ԍ�] = rst![�G�g�H��}�ԍ�]
            ksz = rst![�G�g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�G�g�^��] = rsk![�^��]
            
            rso![�i�q���H��}�ԍ�] = rst![�i�q���H��}�ԍ�]
            ksz = rst![�i�q���H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�i�q���^��] = rsk![�^��]
            
            hds = rst![�ϐ���ٰ��2]
            rSM.FindFirst "[�ϐ���ٰ��]='" & hds & "'"
            
            If rSM![MW����] = True Then
                rso![YMW] = (rso![W���@�l] / 2) + rSM![MW�p�ϐ�]
            Else
                rso![YMW] = rso![W���@�l] + rSM![MW�p�ϐ�]
            End If
            
            rso![YMH] = rso![H���@�l] + rSM![MH�p�ϐ�]
            
            rso![�㉺�g�ؒf���@] = rso![YMW] + rSM![��g���g���@�p]
            rso![�G�g�ؒf���@] = rso![YMH] + rSM![�G�g���@�p]

            
            If rst![���i����CD] = "S3" Then
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
            End If
            If rst![���i����CD] = "S4" Then
                rso![�i�q�G�ؒf���@] = rso![YMH] + rSM![�i�q�G���@�p]
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
                STTD = rso![�i�q�G�ؒf���@]
            End If
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                If rst![���i����CD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![���i����CD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            If rso![YM] = 0 Then
               MsgBox (rst![���i����])
               MsgBox (rso![YMH])
               MsgBox (rSM![H�����}�X�ڕ��q])
               MsgBox (rSM![H�����}�X�ڕ���])
            End If
            
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![���i����CD] <> "S3" Then
                 dobn1 = rso![YPW] * rso![YPW]
                 dobn2 = rso![YPH] * rso![YPH]
                 dobn3 = dobn1 + dobn2
                 dobn4 = Sqr(dobn3)
                 dobn = dobn4 / 2
                 rso![YP1] = dobn
                 P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![��g���g�i�q��t�߯�]
            If rst![���i����CD] = "S3" Or rst![���i����CD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![�G�g�i�q��t�߯�]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![�G�g�i�q��t�߯�]
            End If
            
            STNO = rso![�����ԍ�]
            STEN = rso![�����}��]
            STDT = rso![DT�敪]
            STAD = rSM![�i�qAD�l]
            
            If rst![���i����CD] = "S2" Then
                rso![FS�敪] = False
            End If
            
            If rst![���i����CD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If
            
            rso![���x�b�g] = get_ri(rso![YN], rso![YM], rso![���i����CD]) * rst![��Đ�2]
 
            rso.Update
                       
            Call S1_set     '�N���X�@����
            
            End If
        
'�N���X���葋�@�d�m�c

'�����E����E�_���{�[���X�^�[�g
    
        
'            rst.Edit
'            rst![�ް��쐬�敪] = True
'            rst.Update
        End If
        rst.MoveNext
    Loop
    
nodt:
    rso.Close
    rsd.Close
    rSM.Close
    rsn.Close
    rst.Close
    rsk.Close
    RECS.Close
    
End Function

Sub Kc_Add() '�V�s�b�`�ʊi�q 2010/04/30

    '����΂�f�[�^�쐬
    Dim DB As Database
    Dim rst As Recordset
    Dim rsd As Recordset
    Dim rso As Recordset
    Dim rSM As Recordset
    Dim rsn As Recordset
    Dim rsk As Recordset
    Dim trc As Long
    Dim sds As String
    Dim hds As String
    Dim ksz As String
    Dim chk As Boolean
    Dim kigou(7) As Variant
    Dim sunpou(7) As Long
    Dim kot As Variant
    
    If DCount("[�����ԍ�]", "j_YOTEI_A") = 0 Then
        Exit Sub
    End If
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("j_YOTEI_A")
    Set rsd = DB.OpenRecordset("q_SAN�ݐ�", dbOpenDynaset)
    Set rSM = DB.OpenRecordset("m_�ϐ�", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ܰ�no")
    Set rso = DB.OpenRecordset("YOTEIܰ�")
    Set rsk = DB.OpenRecordset("m_�H��}�ԍ�", dbOpenDynaset)
        
    Set DBB = CurrentDb
    Set RECS = DBB.OpenRecordset("�i�q��ܰ�")


    rst.MoveFirst
    Do Until rst.EOF

        sds = rst![�����ԍ�]
        rsd.FindFirst "[�����ԍ�]='" & sds & "'"
    
        ' NoMatch �v���p�e�B�̒l�Ɋ�Â��Ė߂�l��ݒ肵�܂��B
        If rsd.NoMatch = False Then
        
            rsn.MoveFirst
            wno = rsn![�̔�no]
            wno = wno + 1
            rsn.Edit
            rsn![�̔�no] = wno
            rsn.Update
        
            rso.AddNew
            rso![�����ԍ�] = rst![�����ԍ�]
            rso![�����}��] = rst![�����}��]
            rso![COL] = rst![�F]
            
            rso![�̔�no] = wno
            rso![DT�敪] = "0"
            rso![�����敪] = rst![�����敪]
            rso![�󒍋敪] = rst![�󒍋敪]
            
            rso![�Г��Ǘ��ԍ�1] = rst![�Г��Ǘ��ԍ�1]
            kot = "0000" + rst![�Г��Ǘ��ԍ�2]
            rso![�Г��Ǘ��ԍ�2] = Right(kot, 4)
            
            rso![�[����] = rst![�[����]
            rso![�����\�辯Đ�] = rst![�����\�辯Đ�]
            rso![����] = rst![����]
           
            kigou(0) = rsd![���@�L��1]
            kigou(1) = rsd![���@�L��2]
            kigou(2) = rsd![���@�L��3]
            kigou(3) = rsd![���@�L��4]
            kigou(4) = rsd![���@�L��5]
            kigou(5) = rsd![���@�L��6]
            kigou(6) = rsd![���@�L��7]
            kigou(7) = rsd![���@�L��8]
            sunpou(0) = rsd![���@�l1]
            sunpou(1) = rsd![���@�l2]
            sunpou(2) = rsd![���@�l3]
            sunpou(3) = rsd![���@�l4]
            sunpou(4) = rsd![���@�l5]
            sunpou(5) = rsd![���@�l6]
            sunpou(6) = rsd![���@�l7]
            sunpou(7) = rsd![���@�l8]
            
            If rst![����ϰ�] = "1" Then
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H2" Then
                        rso![H���@�l] = sunpou(lct) / 10
                        'Exit Do
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
                
                lct = 0        ' �ϐ������������܂��B
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' ������ True �ł����
                        chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                        Exit Do         ' ���[�v���甲���܂��B
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W2" Then
                        rso![W���@�l] = sunpou(lct) / 10
                        'Exit Do
                    End If
                    lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                Loop
            Else
                rso![H���@�l] = rst![H��]
                rso![W���@�l] = rst![W��]
            End If
            
            rso![���i����] = rst![���i����]
            rso![���i����CD] = rst![���i����CD]
            rso![�ϐ���ٰ��] = rst![�ϐ���ٰ��]
            rso![��g�H��}�ԍ�] = rst![��g�H��}�ԍ�]
            ksz = rst![��g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![��g�^��] = rsk![�^��]
            
            rso![���g�H��}�ԍ�] = rst![���g�H��}�ԍ�]
            ksz = rst![���g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![���g�^��] = rsk![�^��]
            
            rso![�G�g�H��}�ԍ�] = rst![�G�g�H��}�ԍ�]
            ksz = rst![�G�g�H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�G�g�^��] = rsk![�^��]
            
            rso![�i�q���H��}�ԍ�] = rst![�i�q���H��}�ԍ�]
            ksz = rst![�i�q���H��}�ԍ�]
            rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
            rso![�i�q���^��] = rsk![�^��]
            
            hds = rst![�ϐ���ٰ��]
            rSM.FindFirst "[�ϐ���ٰ��]='" & hds & "'"
            
            
            If rSM![MW����] = True Then
                rso![YMW] = (rso![W���@�l] / 2) + rSM![MW�p�ϐ�]
            Else
                rso![YMW] = rso![W���@�l] + rSM![MW�p�ϐ�]
            End If
            
            rso![YMH] = rso![H���@�l] + rSM![MH�p�ϐ�]
            
            rso![�㉺�g�ؒf���@] = rso![YMW] + rSM![��g���g���@�p]
            rso![�G�g�ؒf���@] = rso![YMH] + rSM![�G�g���@�p]

            '=============================�i�q
            Select Case rst![���i����CD]
             Case "S3A"
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
             Case "S4A", "S5A"
                rso![�i�q�G�ؒf���@] = rso![YMH] + rSM![�i�q�G���@�p]
                rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                STYD = rso![�i�q���ؒf���@]
                STTD = rso![�i�q�G�ؒf���@]
            End Select
            
            '---------M,N
            Select Case rst![���i����CD]
             Case "S1A", "S2A"
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                rso![YN] = kiriage(dobn, 0)
                dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
                rso![YM] = kiriage(dobn, 0)
             Case "S3A" '��
                dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
                rso![YM] = shisya2(dobn, 0)
             Case "S4A", "S5A" '�e
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                rso![YN] = kirishute(dobn, 0) + 2
                dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
                rso![YM] = kirishute(dobn, 0) + 2
            End Select
            N1W = rso![YN]
            M1W = rso![YM]
            
            
            '---------PW,PH �i�q�s�b�`
            If rst![���i����CD] <> "S3A" Then
                dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rso![YN]
                rso![YPW] = shisya2(dobn, 1)
            End If
            dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rso![YM]
            
            If rst![���i����CD] = "S4A" Or rst![���i����CD] = "S5A" Then
                rso![YPW] = 100
                rso![YPH] = 100
            Else
                rso![YPH] = shisya2(dobn, 1)
            End If
            
            If rst![���i����CD] <> "S3A" Then
                dobn1 = rso![YPW] * rso![YPW]
                dobn2 = rso![YPH] * rso![YPH]
                dobn3 = dobn1 + dobn2
                dobn4 = Sqr(dobn3)
                dobn = dobn4 / 2
                rso![YP1] = shisya2(dobn, 1)
                P1W = rso![YP1]
            End If
            
            '---------YA1,YA2 �gAB�l
            rso![YA1] = rso![YPW] / 2 + rSM![��g���g�i�q��t�߯�]
            If rst![���i����CD] = "S3A" Then
                rso![YA2] = rso![YPH] + rSM![�G�g�i�q��t�߯�]
            ElseIf rst![���i����CD] = "S4A" And rst![�ϐ���ٰ��] = "239" Then
                rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 50 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![���i����CD] = "S4A" And rst![�ϐ���ٰ��] = "283" Then '�X�}�[�W���p
                rso![YA1] = (rso![YMW] - 54 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 64 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![���i����CD] = "S4A" And rst![�ϐ���ٰ��] >= "266" Then
                rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![���i����CD] = "S5A" And rst![�ϐ���ٰ��] >= "266" Then
                rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![���i����CD] = "S4A" And rst![�ϐ���ٰ��] >= "264" Then
                rso![YA1] = (rso![YMW] - 54 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 64 - (100 * (rso![YM] - 2))) / 2
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![�G�g�i�q��t�߯�]
            End If
            
            
            If rst![���i����CD] = "S3A" Then
                rso![FS�敪] = True
            End If
            
            STNO = rso![�����ԍ�]
            STEN = rso![�����}��]
            STDT = rso![DT�敪]
            STAD = rSM![�i�qAD�l]
            
            If rst![���i����CD] = "S2A" Or rst![���i����CD] = "S5A" Then
                rso![FS�敪] = True
            End If
            
            If rst![���i����CD] = "S4A" Or rst![���i����CD] = "S5A" Then
                STPW = rso![YA1]
                STPH = rso![YA2]
            End If
 
 
            rso![���x�b�g] = get_ri(rso![YN], rso![YM], rso![���i����CD]) * rst![��Đ�]
            
            rso.Update
           
            '�i�q���f�[�^�̏o��
            Select Case rst![���i����CD]
                Case "S1A"
                    Call S1_set     '�N���X
                Case "S2A"
                    Call S1_set     '�N���X�@����
                Case "S3A"
                    Call S3_set     '��
                Case "S4A", "S5A"
                    Call S4A_set    '�e
            End Select
            
'�Q���`�[�X�^�[�g
            If rst![���i����CD] = "S2A" Or rst![���i����CD] = "S5A" Then
                rsn.MoveFirst
                wno = rsn![�̔�no]
                wno = wno + 1
                rsn.Edit
                rsn![�̔�no] = wno
                rsn.Update
        
                rso.AddNew
                rso![�����ԍ�] = rst![�����ԍ�]
                rso![�����}��] = rst![�����}��]
                rso![COL] = rst![�F]
                rso![�̔�no] = wno
                rso![DT�敪] = "0"
                rso![�����敪] = rst![�����敪]
                rso![�󒍋敪] = rst![�󒍋敪]
                rso![�Г��Ǘ��ԍ�1] = rst![�Г��Ǘ��ԍ�1]
                kot = "0000" + rst![�Г��Ǘ��ԍ�2]
                rso![�Г��Ǘ��ԍ�2] = Right(kot, 4)
                rso![�[����] = rst![�[����]
                rso![�����\�辯Đ�] = rst![�����\�辯Đ�]
                rso![����] = rst![����]
           
                kigou(0) = rsd![���@�L��1]
                kigou(1) = rsd![���@�L��2]
                kigou(2) = rsd![���@�L��3]
                kigou(3) = rsd![���@�L��4]
                kigou(4) = rsd![���@�L��5]
                kigou(5) = rsd![���@�L��6]
                kigou(6) = rsd![���@�L��7]
                kigou(7) = rsd![���@�L��8]
                sunpou(0) = rsd![���@�l1]
                sunpou(1) = rsd![���@�l2]
                sunpou(2) = rsd![���@�l3]
                sunpou(3) = rsd![���@�l4]
                sunpou(4) = rsd![���@�l5]
                sunpou(5) = rsd![���@�l6]
                sunpou(6) = rsd![���@�l7]
                sunpou(7) = rsd![���@�l8]
            
                If rst![����ϰ�] = "1" Then
                    lct = 0        ' �ϐ������������܂��B
                    chk = True
                    Do While lct < 8        '
                        If lct = 8 Then     ' ������ True �ł����
                            chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                            Exit Do         ' ���[�v���甲���܂��B
                        End If
                        If kigou(lct) = "H" Or kigou(lct) = "H1" Then
                            rso![H���@�l] = sunpou(lct) / 10
                            'Exit Do
                        End If
                        lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                    Loop
                
                    lct = 0        ' �ϐ������������܂��B
                    chk = True
                    Do While lct < 8        '
                        If lct = 8 Then     ' ������ True �ł����
                            chk = False     ' �t���O�̒l�� False �ɐݒ肵�܂��B
                            Exit Do         ' ���[�v���甲���܂��B
                        End If
                        If kigou(lct) = "W" Or kigou(lct) = "W1" Then
                            rso![W���@�l] = sunpou(lct) / 10
                            'Exit Do
                        End If
                        lct = lct + 1       ' �J�E���^�𑝂₵�܂��B
                    Loop
                Else
                    rso![H���@�l] = rst![H��2]
                    rso![W���@�l] = rst![W��2]
                End If
            
                rso![���i����] = rst![���i����]
                rso![���i����CD] = rst![���i����CD]
                rso![�ϐ���ٰ��] = rst![�ϐ���ٰ��2]
                rso![��g�H��}�ԍ�] = rst![��g�H��}�ԍ�]
                ksz = rst![��g�H��}�ԍ�]
                rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
                rso![��g�^��] = rsk![�^��]
            
                rso![���g�H��}�ԍ�] = rst![���g�H��}�ԍ�]
                ksz = rst![���g�H��}�ԍ�]
                rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
                rso![���g�^��] = rsk![�^��]
            
                rso![�G�g�H��}�ԍ�] = rst![�G�g�H��}�ԍ�]
                ksz = rst![�G�g�H��}�ԍ�]
                rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
                rso![�G�g�^��] = rsk![�^��]
            
                rso![�i�q���H��}�ԍ�] = rst![�i�q���H��}�ԍ�]
                ksz = rst![�i�q���H��}�ԍ�]
                rsk.FindFirst "[�H��}�ԍ�]='" & ksz & "'"
                rso![�i�q���^��] = rsk![�^��]
                
                hds = rst![�ϐ���ٰ��2]
                rSM.FindFirst "[�ϐ���ٰ��]='" & hds & "'"
            
            
                If rSM![MW����] = True Then
                    rso![YMW] = (rso![W���@�l] / 2) + rSM![MW�p�ϐ�]
                Else
                    rso![YMW] = rso![W���@�l] + rSM![MW�p�ϐ�]
                End If
                
                rso![YMH] = rso![H���@�l] + rSM![MH�p�ϐ�]
                
                rso![�㉺�g�ؒf���@] = rso![YMW] + rSM![��g���g���@�p]
                rso![�G�g�ؒf���@] = rso![YMH] + rSM![�G�g���@�p]

                '=============================�i�q
                Select Case rst![���i����CD]
                 Case "S3A"
                    rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                    STYD = rso![�i�q���ؒf���@]
                 Case "S4A", "S5A"
                    rso![�i�q�G�ؒf���@] = rso![YMH] + rSM![�i�q�G���@�p]
                    rso![�i�q���ؒf���@] = rso![YMW] + rSM![�i�q�����@�p]
                    STYD = rso![�i�q���ؒf���@]
                    STTD = rso![�i�q�G�ؒf���@]
                End Select
            
                '---------M,N
                Select Case rst![���i����CD]
                 Case "S1A", "S2A"
                    dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                    rso![YN] = kiriage(dobn, 0)
                    dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
                    rso![YM] = kiriage(dobn, 0)
                 Case "S3A" '��
                    dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
                    rso![YM] = shisya2(dobn, 0)
                 Case "S4A", "S5A" '�e
                    dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rSM![W�����}�X�ڕ���]
                    rso![YN] = kirishute(dobn, 0) + 2
                    dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rSM![H�����}�X�ڕ���]
                    rso![YM] = kirishute(dobn, 0) + 2
                End Select
                N1W = rso![YN]
                M1W = rso![YM]
            
            
                '---------PW,PH �i�q�s�b�`
                If rst![���i����CD] <> "S3A" Then
                    dobn = (rso![YMW] + rSM![W�����}�X�ڕ��q]) / rso![YN]
                    rso![YPW] = shisya2(dobn, 1)
                End If
                dobn = (rso![YMH] + rSM![H�����}�X�ڕ��q]) / rso![YM]
                
                If rst![���i����CD] = "S4A" Or rst![���i����CD] = "S5A" Then
                    rso![YPW] = 100
                    rso![YPH] = 100
                Else
                    rso![YPH] = shisya2(dobn, 1)
                End If
            
                If rst![���i����CD] <> "S3A" Then
                    dobn1 = rso![YPW] * rso![YPW]
                    dobn2 = rso![YPH] * rso![YPH]
                    dobn3 = dobn1 + dobn2
                    dobn4 = Sqr(dobn3)
                    dobn = dobn4 / 2
                    rso![YP1] = shisya2(dobn, 1)
                    P1W = rso![YP1]
                End If
            
                '---------YA1,YA2 �gAB�l
                rso![YA1] = rso![YPW] / 2 + rSM![��g���g�i�q��t�߯�]
                If rst![���i����CD] = "S3A" Then
                    rso![YA2] = rso![YPH] + rSM![�G�g�i�q��t�߯�]
                ElseIf rst![���i����CD] = "S4A" And rst![�ϐ���ٰ��] = "239" Then
                    rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 50 - (100 * (rso![YM] - 2))) / 2
                ElseIf rst![���i����CD] = "S4A" And rst![�ϐ���ٰ��] >= "266" Then
                    rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
                ElseIf rst![���i����CD] = "S5A" And rst![�ϐ���ٰ��] >= "266" Then
                    rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
                ElseIf rst![���i����CD] = "S4A" And rst![�ϐ���ٰ��] >= "264" Then
                    rso![YA1] = (rso![YMW] - 54 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 64 - (100 * (rso![YM] - 2))) / 2
                Else
                    rso![YA2] = rso![YPH] / 2 + rSM![�G�g�i�q��t�߯�]
                End If
            
            
                STNO = rso![�����ԍ�]
                STEN = rso![�����}��]
                STDT = rso![DT�敪]
                STAD = rSM![�i�qAD�l]
            
                If rst![���i����CD] = "S2A" Or rst![���i����CD] = "S5A" Then
                    rso![FS�敪] = False
                End If
            
                If rst![���i����CD] = "S4A" Or rst![���i����CD] = "S5A" Then
                    STPW = rso![YA1]
                    STPH = rso![YA2]
                End If
 
 
                rso![���x�b�g] = get_ri(rso![YN], rso![YM], rso![���i����CD]) * rst![��Đ�2]
                
                rso.Update
           
                '�i�q���f�[�^�̏o��
                Select Case rst![���i����CD]
                    Case "S2A"
                        Call S1_set     '�N���X�@����
                    Case "S5A"
                        Call S4A_set    '�e
                End Select
            End If
            
        
'�Q���`�[�@�d�m�c

'�����E����E�_���{�[���X�^�[�g
    
        
'            rst.Edit
'            rst![�ް��쐬�敪] = True
'            rst.Update
        End If
        rst.MoveNext
    Loop
    
nodt:
    rso.Close
    rsd.Close
    rSM.Close
    rsn.Close
    rst.Close
    rsk.Close
    RECS.Close
    
End Sub

Public Sub S1_set()

    Dim epchk As Integer
    
    If M1W <= 3 Then                ' M <= 3
        Call S1_M3
    Else                            ' M >= 4
        epchk = M1W Mod 2
        If epchk = 0 Then
            Call S1_NM3             ' M = ����
        Else                        ' M = �
            If N1W = M1W Then
                Call S1_NS35          ' N = M
            End If
            If N1W > M1W Then
                Call S1_NS36          ' N > M
            End If
            If N1W < M1W Then
                Call S1_NS37          ' N < M
            End If
        End If
    End If
    
End Sub

Public Sub XX_set()

    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = 1
    RECS![�ؒf���@] = STYD
    RECS![�ؒf�{��] = N1W
    RECS![���@1] = STAD
    RECS.Update
    
End Sub
Public Sub S3_set()

    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = 1
    RECS![�ؒf���@] = STYD
    RECS![�ؒf�{��] = M1W - 1
    RECS![���@1] = STAD
    RECS.Update
    
End Sub
Public Sub S4_set()

    Dim intA As Integer
    Dim M1WW As Integer
    Dim GHON As Integer
    Dim KHON As Integer
    Dim RECNO As Integer
    
    '�G�i�q
    RECNO = 1
    
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STTD
    RECS![���@1] = 6.5
    RECS![�i�q�敪] = "1"
    
    If N1W = 2 Then
        RECS![�ؒf�{��] = 1
        
        dobn = (M1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![����] = M1WW
        
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![���@2] = 6.5 + STPH * lct * 2
            Case 2
                RECS![���@3] = 6.5 + STPH * lct * 2
            Case 3
                RECS![���@4] = 6.5 + STPH * lct * 2
            Case 4
                RECS![���@5] = 6.5 + STPH * lct * 2
            Case 5
                RECS![���@6] = 6.5 + STPH * lct * 2
            Case 6
                RECS![���@7] = 6.5 + STPH * lct * 2
            Case 7
                RECS![���@8] = 6.5 + STPH * lct * 2
            Case 8
                RECS![���@9] = 6.5 + STPH * lct * 2
            Case 9
                RECS![���@10] = 6.5 + STPH * lct * 2
            Case 10
                RECS![���@11] = 6.5 + STPH * lct * 2
            Case 11
                RECS![���@12] = 6.5 + STPH * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (M1W Mod 2) <> 0 Then
            'RECS![�ؒf�{��] = N1W      ���ԕ��C��
            RECS![�ؒf�{��] = N1W - 1
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = 6.5 + STPH * lct * 2
                Case 2
                    RECS![���@3] = 6.5 + STPH * lct * 2
                Case 3
                    RECS![���@4] = 6.5 + STPH * lct * 2
                Case 4
                    RECS![���@5] = 6.5 + STPH * lct * 2
                Case 5
                    RECS![���@6] = 6.5 + STPH * lct * 2
                Case 6
                    RECS![���@7] = 6.5 + STPH * lct * 2
                Case 7
                    RECS![���@8] = 6.5 + STPH * lct * 2
                Case 8
                    RECS![���@9] = 6.5 + STPH * lct * 2
                Case 9
                    RECS![���@10] = 6.5 + STPH * lct * 2
                Case 10
                    RECS![���@11] = 6.5 + STPH * lct * 2
                Case 11
                    RECS![���@12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (N1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![�ؒf�{��] = KHON
            
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = 6.5 + STPH * lct * 2
                Case 2
                    RECS![���@3] = 6.5 + STPH * lct * 2
                Case 3
                    RECS![���@4] = 6.5 + STPH * lct * 2
                Case 4
                    RECS![���@5] = 6.5 + STPH * lct * 2
                Case 5
                    RECS![���@6] = 6.5 + STPH * lct * 2
                Case 6
                    RECS![���@7] = 6.5 + STPH * lct * 2
                Case 7
                    RECS![���@8] = 6.5 + STPH * lct * 2
                Case 8
                    RECS![���@9] = 6.5 + STPH * lct * 2
                Case 9
                    RECS![���@10] = 6.5 + STPH * lct * 2
                Case 10
                    RECS![���@11] = 6.5 + STPH * lct * 2
                Case 11
                    RECS![���@12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = RECNO
            RECS![�ؒf���@] = STTD
            RECS![���@1] = 6.5
            RECS![�i�q�敪] = "1"
            
            RECS![�ؒf�{��] = N1W - 1 - KHON
            
            dobn = M1W / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                Case 1
                    RECS![���@2] = 6.5 + STPH
                Case 2
                    RECS![���@3] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 3
                    RECS![���@4] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 4
                    RECS![���@5] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 5
                    RECS![���@6] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 6
                    RECS![���@7] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 7
                    RECS![���@8] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 8
                    RECS![���@9] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 9
                    RECS![���@10] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 10
                    RECS![���@11] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 11
                    RECS![���@12] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
    
    '���i�q�@���ԕ�
    RECNO = RECNO + 1
    
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STYD
    RECS![���@1] = STAD
    RECS![�i�q�敪] = "2"
    
    If M1W = 2 Then
        RECS![�ؒf�{��] = 1
        dobn = (N1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![����] = M1WW
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![���@2] = STAD + STPW * lct * 2
            Case 2
                RECS![���@3] = STAD + STPW * lct * 2
            Case 3
                RECS![���@4] = STAD + STPW * lct * 2
            Case 4
                RECS![���@5] = STAD + STPW * lct * 2
            Case 5
                RECS![���@6] = STAD + STPW * lct * 2
            Case 6
                RECS![���@7] = STAD + STPW * lct * 2
            Case 7
                RECS![���@8] = STAD + STPW * lct * 2
            Case 8
                RECS![���@9] = STAD + STPW * lct * 2
            Case 9
                RECS![���@10] = STAD + STPW * lct * 2
            Case 10
                RECS![���@11] = STAD + STPW * lct * 2
            Case 11
                RECS![���@12] = STAD + STPW * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (N1W Mod 2) <> 0 Then
            RECS![�ؒf�{��] = M1W - 1
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = STAD + STPW * lct * 2
                Case 2
                    RECS![���@3] = STAD + STPW * lct * 2
                Case 3
                    RECS![���@4] = STAD + STPW * lct * 2
                Case 4
                    RECS![���@5] = STAD + STPW * lct * 2
                Case 5
                    RECS![���@6] = STAD + STPW * lct * 2
                Case 6
                    RECS![���@7] = STAD + STPW * lct * 2
                Case 7
                    RECS![���@8] = STAD + STPW * lct * 2
                Case 8
                    RECS![���@9] = STAD + STPW * lct * 2
                Case 9
                    RECS![���@10] = STAD + STPW * lct * 2
                Case 10
                    RECS![���@11] = STAD + STPW * lct * 2
                Case 11
                    RECS![���@12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (M1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![�ؒf�{��] = KHON
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = STAD + STPW * lct * 2
                Case 2
                    RECS![���@3] = STAD + STPW * lct * 2
                Case 3
                    RECS![���@4] = STAD + STPW * lct * 2
                Case 4
                    RECS![���@5] = STAD + STPW * lct * 2
                Case 5
                    RECS![���@6] = STAD + STPW * lct * 2
                Case 6
                    RECS![���@7] = STAD + STPW * lct * 2
                Case 7
                    RECS![���@8] = STAD + STPW * lct * 2
                Case 8
                    RECS![���@9] = STAD + STPW * lct * 2
                Case 9
                    RECS![���@10] = STAD + STPW * lct * 2
                Case 10
                    RECS![���@11] = STAD + STPW * lct * 2
                Case 11
                    RECS![���@12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = RECNO
            RECS![�ؒf���@] = STYD
            RECS![���@1] = STAD
            RECS![�i�q�敪] = "2"
            
            RECS![�ؒf�{��] = M1W - 1 - KHON
            
            dobn = N1W / 2
            M1WW = kirishute(dobn, 0)
            
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = STAD + STPW
                Case 2
                    RECS![���@3] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 3
                    RECS![���@4] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 4
                    RECS![���@5] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 5
                    RECS![���@6] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 6
                    RECS![���@7] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 7
                    RECS![���@8] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 8
                    RECS![���@9] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 9
                    RECS![���@10] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 10
                    RECS![���@11] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 11
                    RECS![���@12] = STAD + STPW + (STPW * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
         
    '���i�q �㉺�[��
    RECNO = RECNO + 1
         
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STYD
    RECS![���@1] = STAD
    RECS![�i�q�敪] = "3"
    RECS![�ؒf�{��] = 2
         
    If (N1W Mod 2) = 0 Then
        M1WW = N1W
        RECS![����] = M1WW - 1
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![���@2] = STAD + STPW * lct
            Case 2
                RECS![���@3] = STAD + STPW * lct
            Case 3
                RECS![���@4] = STAD + STPW * lct
            Case 4
                RECS![���@5] = STAD + STPW * lct
            Case 5
                RECS![���@6] = STAD + STPW * lct
            Case 6
                RECS![���@7] = STAD + STPW * lct
            Case 7
                RECS![���@8] = STAD + STPW * lct
            Case 8
                RECS![���@9] = STAD + STPW * lct
            Case 9
                RECS![���@10] = STAD + STPW * lct
            Case 10
                RECS![���@11] = STAD + STPW * lct
            Case 11
                RECS![���@12] = STAD + STPW * lct
            End Select
        Next lct
        RECS.Update
    Else
        M1WW = N1W
        RECS![����] = M1WW - 1
        intA = M1WW / 2
        For lct = 1 To intA
            Select Case lct
            Case 1
                If lct = intA Then
                    RECS![���@2] = STYD / 2
                Else
                    RECS![���@2] = STAD + STPW * lct
                End If
            Case 2
                If lct = intA Then
                    RECS![���@3] = STYD / 2
                Else
                    RECS![���@3] = STAD + STPW * lct
                End If
            Case 3
                If lct = intA Then
                    RECS![���@4] = STYD / 2
                Else
                    RECS![���@4] = STAD + STPW * lct
                End If
            Case 4
                If lct = intA Then
                    RECS![���@5] = STYD / 2
                Else
                    RECS![���@5] = STAD + STPW * lct
                End If
            Case 5
                If lct = intA Then
                    RECS![���@6] = STYD / 2
                Else
                    RECS![���@6] = STAD + STPW * lct
                End If
            Case 6
                If lct = intA Then
                    RECS![���@7] = STYD / 2
                Else
                    RECS![���@7] = STAD + STPW * lct
                End If
            Case 7
                If lct = intA Then
                    RECS![���@8] = STYD / 2
                Else
                    RECS![���@8] = STAD + STPW * lct
                End If
            Case 8
                If lct = intA Then
                    RECS![���@9] = STYD / 2
                Else
                    RECS![���@9] = STAD + STPW * lct
                End If
            Case 9
                If lct = intA Then
                    RECS![���@10] = STYD / 2
                Else
                    RECS![���@10] = STAD + STPW * lct
                End If
            Case 10
                If lct = intA Then
                    RECS![���@11] = STYD / 2
                Else
                    RECS![���@11] = STAD + STPW * lct
                End If
            Case 11
                If lct = intA Then
                    RECS![���@12] = STYD / 2
                Else
                    RECS![���@12] = STAD + STPW * lct
                End If
            End Select
        Next lct
        
        For lct = intA To M1WW - 1
            Select Case lct
            Case 1
                    RECS![���@2] = STAD + STPW * lct
            Case 2
                    RECS![���@3] = STAD + STPW * lct
            Case 3
                    RECS![���@4] = STAD + STPW * lct
            Case 4
                    RECS![���@5] = STAD + STPW * lct
            Case 5
                    RECS![���@6] = STAD + STPW * lct
            Case 6
                    RECS![���@7] = STAD + STPW * lct
            Case 7
                    RECS![���@8] = STAD + STPW * lct
            Case 8
                    RECS![���@9] = STAD + STPW * lct
            Case 9
                    RECS![���@10] = STAD + STPW * lct
            Case 10
                    RECS![���@11] = STAD + STPW * lct
            Case 11
                    RECS![���@12] = STAD + STPW * lct
            End Select
        Next lct
        RECS.Update
    End If
     
End Sub

Public Sub S4A_set()
    Dim intA As Integer
    Dim M1WW As Integer
    Dim GHON As Integer
    Dim KHON As Integer
    Dim RECNO As Integer
    
    '�G�i�q
    RECNO = 1
    
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STTD
    RECS![���@1] = 6.5
    RECS![�i�q�敪] = "1"
    
    If N1W = 2 Then 'N1W=YN OK!
        RECS![�ؒf�{��] = 1

        dobn = (M1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![����] = M1WW

        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![���@2] = 6.5 + STPH + 100
            Case 2
                RECS![���@3] = RECS![���@2] + 200
            Case 3
                RECS![���@4] = RECS![���@3] + 200
            Case 4
                RECS![���@5] = RECS![���@4] + 200
            Case 5
                RECS![���@6] = RECS![���@5] + 200
            Case 6
                RECS![���@7] = RECS![���@6] + 200
            Case 7
                RECS![���@8] = RECS![���@7] + 200
            Case 8
                RECS![���@9] = RECS![���@8] + 200
            Case 9
                RECS![���@10] = RECS![���@9] + 200
            Case 10
                RECS![���@11] = RECS![���@10] + 200
            Case 11
                RECS![���@12] = RECS![���@11] + 200
            End Select
        Next lct
        RECS.Update
    Else
        If (M1W Mod 2) <> 0 Then '�c�����e�ڐ����(���i�q�{��������)�Ȃ�P�p�^�[���łn�j N=6,n=5
            RECS![�ؒf�{��] = N1W - 1
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![���@3] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![���@4] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![���@5] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![���@6] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![���@7] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![���@8] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![���@9] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![���@10] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![���@11] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![���@12] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
        Else '�c�����e�ڐ����(���i�q�{��������)�Ȃ�i�q���H�p�^�[���͂Q�p�^�[�� N=7,n=6
            dobn = (N1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![�ؒf�{��] = KHON
            
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![���@3] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![���@4] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![���@5] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![���@6] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![���@7] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![���@8] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![���@9] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![���@10] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![���@11] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![���@12] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = RECNO
            RECS![�ؒf���@] = STTD
            RECS![���@1] = 6.5
            RECS![�i�q�敪] = "1"
            
            RECS![�ؒf�{��] = N1W - 1 - KHON
            
            dobn = M1W / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                Case 1
                    RECS![���@2] = 6.5 + STPH
                Case 2
                    RECS![���@3] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 3
                    RECS![���@4] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 4
                    RECS![���@5] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 5
                    RECS![���@6] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 6
                    RECS![���@7] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 7
                    RECS![���@8] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 8
                    RECS![���@9] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 9
                    RECS![���@10] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 10
                    RECS![���@11] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 11
                    RECS![���@12] = 6.5 + STPH + (lct - 1) * 2 * 100
                End Select
            Next lct
            RECS.Update
        End If
    End If
    
    '���i�q�@���ԕ�
    RECNO = RECNO + 1
    
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STYD
    RECS![���@1] = STAD
    RECS![�i�q�敪] = "2"
    
    If M1W = 2 Then
        RECS![�ؒf�{��] = 1
        dobn = (N1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![����] = M1WW
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![���@2] = STAD + STPW + 100
            Case 2
                RECS![���@3] = STAD + STPW + (lct - 1) * 2 * 100
            Case 3
                RECS![���@4] = STAD + STPW + (lct - 1) * 2 * 100
            Case 4
                RECS![���@5] = STAD + STPW + (lct - 1) * 2 * 100
            Case 5
                RECS![���@6] = STAD + STPW + (lct - 1) * 2 * 100
            Case 6
                RECS![���@7] = STAD + STPW + (lct - 1) * 2 * 100
            Case 7
                RECS![���@8] = STAD + STPW + (lct - 1) * 2 * 100
            Case 8
                RECS![���@9] = STAD + STPW + (lct - 1) * 2 * 100
            Case 9
                RECS![���@10] = STAD + STPW + (lct - 1) * 2 * 100
            Case 10
                RECS![���@11] = STAD + STPW + (lct - 1) * 2 * 100
            Case 11
                RECS![���@12] = STAD + STPW + (lct - 1) * 2 * 100
            End Select
        Next lct
        RECS.Update
    Else
        If (N1W Mod 2) <> 0 Then  '�������e�ڐ����(�G�i�q�{��������)�Ȃ�i�q���H�p�^�[���͂P�p�^�[�� M=5,m=4
            RECS![�ؒf�{��] = M1W - 1
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![���@3] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![���@4] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![���@5] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![���@6] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![���@7] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![���@8] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![���@9] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![���@10] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![���@11] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![���@12] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
        Else '''''�c�����e�ڐ����(���i�q�{��������)�Ȃ�i�q���H�p�^�[���͂Q�p�^�[�� N=7,n=6
            dobn = (M1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![�ؒf�{��] = KHON
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![���@3] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![���@4] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![���@5] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![���@6] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![���@7] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![���@8] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![���@9] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![���@10] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![���@11] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![���@12] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = RECNO
            RECS![�ؒf���@] = STYD
            RECS![���@1] = STAD
            RECS![�i�q�敪] = "2"
            
            RECS![�ؒf�{��] = M1W - 1 - KHON
            
            dobn = N1W / 2
            M1WW = kirishute(dobn, 0)
            
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![���@2] = STAD + STPW
                Case 2
                    RECS![���@3] = STAD + STPW + (lct - 1) * 2 * 100
                Case 3
                    RECS![���@4] = STAD + STPW + (lct - 1) * 2 * 100
                Case 4
                    RECS![���@5] = STAD + STPW + (lct - 1) * 2 * 100
                Case 5
                    RECS![���@6] = STAD + STPW + (lct - 1) * 2 * 100
                Case 6
                    RECS![���@7] = STAD + STPW + (lct - 1) * 2 * 100
                Case 7
                    RECS![���@8] = STAD + STPW + (lct - 1) * 2 * 100
                Case 8
                    RECS![���@9] = STAD + STPW + (lct - 1) * 2 * 100
                Case 9
                    RECS![���@10] = STAD + STPW + (lct - 1) * 2 * 100
                Case 10
                    RECS![���@11] = STAD + STPW + (lct - 1) * 2 * 100
                Case 11
                    RECS![���@12] = STAD + STPW + (lct - 1) * 2 * 100
                End Select
            Next lct
            RECS.Update
        End If
    End If
         
    '���i�q �㉺�[��
    RECNO = RECNO + 1
         
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STYD
    RECS![���@1] = STAD
    RECS![�i�q�敪] = "3"
    RECS![�ؒf�{��] = 2
         
    M1WW = N1W
    RECS![����] = M1WW - 1
    For lct = 1 To M1WW - 1
        Select Case lct
        Case 1
            RECS![���@2] = STAD + STPW
        Case 2
            RECS![���@3] = STAD + STPW + (lct - 1) * 100
        Case 3
            RECS![���@4] = STAD + STPW + (lct - 1) * 100
        Case 4
            RECS![���@5] = STAD + STPW + (lct - 1) * 100
        Case 5
            RECS![���@6] = STAD + STPW + (lct - 1) * 100
        Case 6
            RECS![���@7] = STAD + STPW + (lct - 1) * 100
        Case 7
            RECS![���@8] = STAD + STPW + (lct - 1) * 100
        Case 8
            RECS![���@9] = STAD + STPW + (lct - 1) * 100
        Case 9
            RECS![���@10] = STAD + STPW + (lct - 1) * 100
        Case 10
            RECS![���@11] = STAD + STPW + (lct - 1) * 100
        Case 11
            RECS![���@12] = STAD + STPW + (lct - 1) * 100
        End Select
    Next lct
    RECS.Update
     
End Sub

Public Sub P5_set() '�i�p�i �䌅�ʊi�q�j�i�q��ܰ��e�[�u���ɒl���

    Dim intA As Integer
    Dim M1WW As Integer
    Dim GHON As Integer
    Dim KHON As Integer
    Dim RECNO As Integer
    
    '�G�i�q
    RECNO = 1
    
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STTD
    RECS![���@1] = 6.5
    RECS![�i�q�敪] = "1"
    
    If N1W = 2 Then  'N1W �F�G�i�q�{�� YN
        RECS![�ؒf�{��] = 1
        
        dobn = (M1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![����] = M1WW
        
        For lct = 1 To M1WW 'M1WW �F
            Select Case lct
            Case 1
                RECS![���@2] = 6.5 + STPH * lct * 2 'STPH�F�i�q�Ԋu YPH
            Case 2
                RECS![���@3] = 6.5 + STPH * lct * 2
            Case 3
                RECS![���@4] = 6.5 + STPH * lct * 2
            Case 4
                RECS![���@5] = 6.5 + STPH * lct * 2
            Case 5
                RECS![���@6] = 6.5 + STPH * lct * 2
            Case 6
                RECS![���@7] = 6.5 + STPH * lct * 2
            Case 7
                RECS![���@8] = 6.5 + STPH * lct * 2
            Case 8
                RECS![���@9] = 6.5 + STPH * lct * 2
            Case 9
                RECS![���@10] = 6.5 + STPH * lct * 2
            Case 10
                RECS![���@11] = 6.5 + STPH * lct * 2
            Case 11
                RECS![���@12] = 6.5 + STPH * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (M1W Mod 2) <> 0 Then
            'RECS![�ؒf�{��] = N1W      ���ԕ��C��
            RECS![�ؒf�{��] = N1W - 1
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![���@2] = 6.5 + STPH * lct * 2
                    Case 2
                        RECS![���@3] = 6.5 + STPH * lct * 2
                    Case 3
                        RECS![���@4] = 6.5 + STPH * lct * 2
                    Case 4
                        RECS![���@5] = 6.5 + STPH * lct * 2
                    Case 5
                        RECS![���@6] = 6.5 + STPH * lct * 2
                    Case 6
                        RECS![���@7] = 6.5 + STPH * lct * 2
                    Case 7
                        RECS![���@8] = 6.5 + STPH * lct * 2
                    Case 8
                        RECS![���@9] = 6.5 + STPH * lct * 2
                    Case 9
                        RECS![���@10] = 6.5 + STPH * lct * 2
                    Case 10
                        RECS![���@11] = 6.5 + STPH * lct * 2
                    Case 11
                        RECS![���@12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (N1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![�ؒf�{��] = KHON
            
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![���@2] = 6.5 + STPH * lct * 2
                    Case 2
                        RECS![���@3] = 6.5 + STPH * lct * 2
                    Case 3
                        RECS![���@4] = 6.5 + STPH * lct * 2
                    Case 4
                        RECS![���@5] = 6.5 + STPH * lct * 2
                    Case 5
                        RECS![���@6] = 6.5 + STPH * lct * 2
                    Case 6
                        RECS![���@7] = 6.5 + STPH * lct * 2
                    Case 7
                        RECS![���@8] = 6.5 + STPH * lct * 2
                    Case 8
                        RECS![���@9] = 6.5 + STPH * lct * 2
                    Case 9
                        RECS![���@10] = 6.5 + STPH * lct * 2
                    Case 10
                        RECS![���@11] = 6.5 + STPH * lct * 2
                    Case 11
                        RECS![���@12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = RECNO
            RECS![�ؒf���@] = STTD
            RECS![���@1] = 6.5
            RECS![�i�q�敪] = "1"
            
            RECS![�ؒf�{��] = N1W - 1 - KHON
            
            dobn = M1W / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![���@2] = 6.5 + STPH
                    Case 2
                        RECS![���@3] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 3
                        RECS![���@4] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 4
                        RECS![���@5] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 5
                        RECS![���@6] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 6
                        RECS![���@7] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 7
                        RECS![���@8] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 8
                        RECS![���@9] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 9
                        RECS![���@10] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 10
                        RECS![���@11] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 11
                        RECS![���@12] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
    
    '���i�q�@���ԕ�
    RECNO = RECNO + 1
    
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STYD
    RECS![���@1] = STAD
    RECS![�i�q�敪] = "2"
    
    If M1W = 2 Then
        RECS![�ؒf�{��] = 1
        dobn = (N1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![����] = M1WW
        For lct = 1 To M1WW
            Select Case lct
                Case 1
                    RECS![���@2] = STAD + STPW * lct * 2
                Case 2
                    RECS![���@3] = STAD + STPW * lct * 2
                Case 3
                    RECS![���@4] = STAD + STPW * lct * 2
                Case 4
                    RECS![���@5] = STAD + STPW * lct * 2
                Case 5
                    RECS![���@6] = STAD + STPW * lct * 2
                Case 6
                    RECS![���@7] = STAD + STPW * lct * 2
                Case 7
                    RECS![���@8] = STAD + STPW * lct * 2
                Case 8
                    RECS![���@9] = STAD + STPW * lct * 2
                Case 9
                    RECS![���@10] = STAD + STPW * lct * 2
                Case 10
                    RECS![���@11] = STAD + STPW * lct * 2
                Case 11
                    RECS![���@12] = STAD + STPW * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (N1W Mod 2) <> 0 Then
            RECS![�ؒf�{��] = M1W - 1
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![���@2] = STAD + STPW * lct * 2
                    Case 2
                        RECS![���@3] = STAD + STPW * lct * 2
                    Case 3
                        RECS![���@4] = STAD + STPW * lct * 2
                    Case 4
                        RECS![���@5] = STAD + STPW * lct * 2
                    Case 5
                        RECS![���@6] = STAD + STPW * lct * 2
                    Case 6
                        RECS![���@7] = STAD + STPW * lct * 2
                    Case 7
                        RECS![���@8] = STAD + STPW * lct * 2
                    Case 8
                        RECS![���@9] = STAD + STPW * lct * 2
                    Case 9
                        RECS![���@10] = STAD + STPW * lct * 2
                    Case 10
                        RECS![���@11] = STAD + STPW * lct * 2
                    Case 11
                        RECS![���@12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (M1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![�ؒf�{��] = KHON
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![���@2] = STAD + STPW * lct * 2
                    Case 2
                        RECS![���@3] = STAD + STPW * lct * 2
                    Case 3
                        RECS![���@4] = STAD + STPW * lct * 2
                    Case 4
                        RECS![���@5] = STAD + STPW * lct * 2
                    Case 5
                        RECS![���@6] = STAD + STPW * lct * 2
                    Case 6
                        RECS![���@7] = STAD + STPW * lct * 2
                    Case 7
                        RECS![���@8] = STAD + STPW * lct * 2
                    Case 8
                        RECS![���@9] = STAD + STPW * lct * 2
                    Case 9
                        RECS![���@10] = STAD + STPW * lct * 2
                    Case 10
                        RECS![���@11] = STAD + STPW * lct * 2
                    Case 11
                        RECS![���@12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = RECNO
            RECS![�ؒf���@] = STYD
            RECS![���@1] = STAD
            RECS![�i�q�敪] = "2"
            
            RECS![�ؒf�{��] = M1W - 1 - KHON
            
            dobn = N1W / 2
            M1WW = kirishute(dobn, 0)
            
            RECS![����] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![���@2] = STAD + STPW
                    Case 2
                        RECS![���@3] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 3
                        RECS![���@4] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 4
                        RECS![���@5] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 5
                        RECS![���@6] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 6
                        RECS![���@7] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 7
                        RECS![���@8] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 8
                        RECS![���@9] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 9
                        RECS![���@10] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 10
                        RECS![���@11] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 11
                        RECS![���@12] = STAD + STPW + (STPW * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
         
    '���i�q �㉺�[��
    RECNO = RECNO + 1
         
    RECS.AddNew
    RECS![�����ԍ�] = STNO
    RECS![�̔�no] = wno
    RECS![�����}��] = STEN
    RECS![DT�敪] = STDT
    RECS![�i�q�ԍ�] = RECNO
    RECS![�ؒf���@] = STYD - 19 'MW-31 STYD(=MW-12)-19
    RECS![�i�q�敪] = "3"
    RECS![�ؒf�{��] = 2
    
    RECS![���@1] = 0
    STAD = STPW - 5.5
    RECS![���@2] = STAD     '�i�qAD�l
         
    If (N1W Mod 2) = 0 Then  'N1W �FW�����s�b�`�i�}�X�ځj�� YN
        M1WW = N1W
        RECS![����] = M1WW - 1
        For lct = 1 To M1WW - 2
            Select Case lct
                Case 1
                    RECS![���@3] = STAD + STPW * lct
                Case 2
                    RECS![���@4] = STAD + STPW * lct
                Case 3
                    RECS![���@5] = STAD + STPW * lct
                Case 4
                    RECS![���@6] = STAD + STPW * lct
                Case 5
                    RECS![���@7] = STAD + STPW * lct
                Case 6
                    RECS![���@8] = STAD + STPW * lct
                Case 7
                    RECS![���@9] = STAD + STPW * lct
                Case 8
                    RECS![���@10] = STAD + STPW * lct
                Case 9
                    RECS![���@11] = STAD + STPW * lct
                Case 10
                    RECS![���@12] = STAD + STPW * lct
            End Select
        Next lct
        RECS.Update
    Else
        M1WW = N1W
        RECS![����] = M1WW - 1
        intA = M1WW / 2
        For lct = 1 To intA
            Select Case lct
            Case 1
                If lct = intA Then
                    RECS![���@3] = STYD / 2
                Else
                    RECS![���@3] = STAD + STPW * lct
                End If
            Case 2
                If lct = intA Then
                    RECS![���@4] = STYD / 2
                Else
                    RECS![���@4] = STAD + STPW * lct
                End If
            Case 3
                If lct = intA Then
                    RECS![���@5] = STYD / 2
                Else
                    RECS![���@5] = STAD + STPW * lct
                End If
            Case 4
                If lct = intA Then
                    RECS![���@6] = STYD / 2
                Else
                    RECS![���@6] = STAD + STPW * lct
                End If
            Case 5
                If lct = intA Then
                    RECS![���@7] = STYD / 2
                Else
                    RECS![���@7] = STAD + STPW * lct
                End If
            Case 6
                If lct = intA Then
                    RECS![���@8] = STYD / 2
                Else
                    RECS![���@8] = STAD + STPW * lct
                End If
            Case 7
                If lct = intA Then
                    RECS![���@9] = STYD / 2
                Else
                    RECS![���@9] = STAD + STPW * lct
                End If
            Case 8
                If lct = intA Then
                    RECS![���@10] = STYD / 2
                Else
                    RECS![���@10] = STAD + STPW * lct
                End If
            Case 9
                If lct = intA Then
                    RECS![���@11] = STYD / 2
                Else
                    RECS![���@11] = STAD + STPW * lct
                End If
            Case 10
                If lct = intA Then
                    RECS![���@12] = STYD / 2
                Else
                    RECS![���@12] = STAD + STPW * lct
                End If
            End Select
        Next lct
        
        For lct = intA To M1WW - 2
            Select Case lct
            Case 1
                    RECS![���@3] = STAD + STPW * lct
            Case 2
                    RECS![���@4] = STAD + STPW * lct
            Case 3
                    RECS![���@5] = STAD + STPW * lct
            Case 4
                    RECS![���@6] = STAD + STPW * lct
            Case 5
                    RECS![���@7] = STAD + STPW * lct
            Case 6
                    RECS![���@8] = STAD + STPW * lct
            Case 7
                    RECS![���@9] = STAD + STPW * lct
            Case 8
                    RECS![���@10] = STAD + STPW * lct
            Case 9
                    RECS![���@11] = STAD + STPW * lct
            Case 10
                    RECS![���@12] = STAD + STPW * lct
            End Select
        Next lct
        RECS.Update
    End If
End Sub



Public Sub Sps_Set1()
    For lct = 1 To RECS![����]
        Select Case lct
            Case 1
                RECS![���@2] = STAD + (P1W * 4) * lct
            Case 2
                RECS![���@3] = STAD + (P1W * 4) * lct
            Case 3
                RECS![���@4] = STAD + (P1W * 4) * lct
            Case 4
                RECS![���@5] = STAD + (P1W * 4) * lct
            Case 5
                RECS![���@6] = STAD + (P1W * 4) * lct
            Case 6
                RECS![���@7] = STAD + (P1W * 4) * lct
            Case 7
                RECS![���@8] = STAD + (P1W * 4) * lct
            Case 8
                RECS![���@9] = STAD + (P1W * 4) * lct
            Case 9
                RECS![���@10] = STAD + (P1W * 4) * lct
            Case 10
                RECS![���@11] = STAD + (P1W * 4) * lct
            Case 11
                RECS![���@12] = STAD + (P1W * 4) * lct
        End Select
    Next lct
End Sub
Public Sub Sps_Set2()
    For lct = 1 To RECS![����]
        Select Case lct
            Case 1
                RECS![���@2] = STAD + (P1W * 3)
            Case 2
                RECS![���@3] = STAD + (P1W * 4) + (P1W * 3)
            Case 3
                RECS![���@4] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 4
                RECS![���@5] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 5
                RECS![���@6] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 6
                RECS![���@7] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 7
                RECS![���@8] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 8
                RECS![���@9] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 9
                RECS![���@10] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 10
                RECS![���@11] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 11
                RECS![���@12] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
        End Select
    Next lct
End Sub
Public Sub Sps_Set3()
    For lct = 1 To RECS![����]
        Select Case lct
            Case 1
                RECS![���@2] = STAD + (P1W * 1)
            Case 2
                RECS![���@3] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 3
                RECS![���@4] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 4
                RECS![���@5] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 5
                RECS![���@6] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 6
                RECS![���@7] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 7
                RECS![���@8] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 8
                RECS![���@9] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 9
                RECS![���@10] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 10
                RECS![���@11] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 11
                RECS![���@12] = STAD + (P1W * 4) * (lct - 1) + P1W
        End Select
    Next lct
End Sub
Public Sub S1_M3()

    If N1W = M1W Then
        XNO = M1W
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�����}��] = STEN
            RECS![�̔�no] = wno
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = fnx
            RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
            RECS![�ؒf�{��] = 2 * 2
            RECS![���@1] = STAD
            RECS.Update
        Next fnx
    End If
    
    If N1W > M1W Then
        XNO = M1W + 1
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = fnx
            If fnx <= M1W Then
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![�ؒf�{��] = 2 * 2
            Else
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * M1W)
                RECS![�ؒf�{��] = (N1W - M1W) * 2
            End If
            RECS![���@1] = STAD
            RECS.Update
        Next fnx
    End If
    
    If N1W < M1W Then
        XNO = N1W + 1
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = fnx
            If fnx <= N1W Then
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![�ؒf�{��] = 2 * 2
            Else
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                RECS![�ؒf�{��] = (M1W - N1W) * 2
            End If
            RECS![���@1] = STAD
            RECS.Update
        Next fnx
    End If
    
End Sub
Public Sub S1_NM3()

    If N1W = M1W Then
        XNO = M1W
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = fnx
            RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
            RECS![�ؒf�{��] = 2 * 2
            If fnx >= 3 Then
                dobn = fnx / 2 - 1
                RECS![����] = kiriage(dobn, 0)
            End If
            RECS![���@1] = STAD
            If RECS![����] >= 1 Then
                Call Sps_Set1
            End If
            RECS.Update
        Next fnx
    End If
    
    If N1W > M1W Then
        XNO = M1W + 1
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![�����ԍ�] = STNO
            RECS![�̔�no] = wno
            RECS![�����}��] = STEN
            RECS![DT�敪] = STDT
            RECS![�i�q�ԍ�] = fnx
            If fnx <= M1W Then
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
            Else
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * M1W)
            End If
            
            If fnx <= M1W Then
                RECS![�ؒf�{��] = 2 * 2
            Else
                RECS![�ؒf�{��] = (N1W - M1W) * 2
            End If
            
            If fnx >= 3 And fnx <= M1W Then
                dobn = fnx / 2 - 1
                RECS![����] = kiriage(dobn, 0)
            End If
            If fnx > M1W Then
                RECS![����] = M1W / 2 - 1
            End If
            
            RECS![���@1] = STAD
            If RECS![����] >= 1 Then
                Call Sps_Set1
            End If
            
            RECS.Update
        Next fnx
    End If
    
    If N1W < M1W Then
        If N1W = 2 Then
            XNO = 3
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                If fnx <= 2 Then
                    RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                Else
                    RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                End If
                
                If fnx <= 2 Then
                    RECS![�ؒf�{��] = 2 * 2
                Else
                    RECS![�ؒf�{��] = (M1W - N1W) * 2
                End If
                
                If fnx = 3 Then
                    RECS![����] = 1
                End If
                RECS![���@1] = STAD
                If RECS![����] >= 1 Then
                    RECS![���@2] = STAD + (P1W * 3)
                End If
                
                RECS.Update
            Next fnx
        Else
            If ((N1W Mod 2) = 0) Or (M1W - N1W = 1) Then
                XNO = N1W + 1
                For fnx = 1 To XNO
                    RECS.AddNew
                    RECS![�����ԍ�] = STNO
                    RECS![�̔�no] = wno
                    RECS![�����}��] = STEN
                    RECS![DT�敪] = STDT
                    RECS![�i�q�ԍ�] = fnx
                    
                    If fnx <= N1W Then
                        RECS![�ؒf�{��] = 2 * 2
                    Else
                        RECS![�ؒf�{��] = (M1W - N1W) * 2
                    End If
            
                    RECS![���@1] = STAD
                    If fnx < 3 Then
                        RECS![����] = 0
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                    End If
                    If fnx >= 3 And fnx <= M1W Then
                        dobn = fnx / 2 - 1
                        RECS![����] = kiriage(dobn, 0)
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        If RECS![����] >= 1 Then
                            Call Sps_Set1
                        End If
                    End If
                    If fnx > N1W Then
                        RECS![����] = fnx / 2 - 1
                        RECS![����] = kiriage(dobn, 0)
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        If RECS![����] >= 1 Then
                            Call Sps_Set2
                        End If
                    End If
                    
                    RECS.Update
                Next fnx
            Else
                XNO = N1W + 2
                For fnx = 1 To XNO
                    RECS.AddNew
                    RECS![�����ԍ�] = STNO
                    RECS![�̔�no] = wno
                    RECS![�����}��] = STEN
                    RECS![DT�敪] = STDT
                    RECS![�i�q�ԍ�] = fnx
                    
                    If fnx <= N1W Then
                        RECS![�ؒf�{��] = 2 * 2
                    End If
                    If fnx = N1W + 1 Then
                        dobn = (M1W - N1W) / 2
                        RECS![�ؒf�{��] = kiriage(dobn, 0) * 2
                    End If
                    If fnx = N1W + 2 Then
                        dobn = (M1W - N1W) / 2
                        RECS![�ؒf�{��] = kirishute(dobn, 0) * 2
                    End If
                    
                    If fnx <= N1W Then
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                    Else
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                    End If
                    
                    RECS![���@1] = STAD
                    
                    If fnx >= 3 Then
                        dobn = fnx / 2 - 1
                        RECS![����] = kiriage(dobn, 0)
                    End If
                    
                    If RECS![����] >= 1 Then
                        If fnx >= 3 And fnx <= N1W Then
                            Call Sps_Set1
                        End If
                        If fnx = (N1W + 1) Then
                            Call Sps_Set2
                        End If
                        If fnx = (N1W + 2) Then
                            Call Sps_Set3
                        End If
                    End If
                    
                    RECS.Update
                Next fnx
            End If
        End If
          
    End If

End Sub
Public Sub S1_NS35()
    Select Case M1W
        Case 5
            XNO = M1W
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![�ؒf�{��] = 2 * 2
                If fnx >= 4 Then
                    RECS![����] = 1
                End If
                RECS![���@1] = STAD
                If RECS![����] >= 1 Then
                    RECS![���@2] = STAD + (P1W * 5)
                End If
                RECS.Update
            Next fnx
        Case 7
            XNO = M1W
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![�ؒf�{��] = 2 * 2
                If fnx >= 3 And fnx <= 5 Then
                    RECS![����] = 1
                End If
                If fnx >= 6 Then
                    RECS![����] = 2
                End If
                RECS![���@1] = STAD
                If RECS![����] >= 1 Then
                    Select Case RECS![����]
                        Case 1
                            RECS![���@2] = STAD + (P1W * 4)
                        Case 2
                            RECS![���@2] = STAD + (P1W * 4)
                            RECS![���@3] = STAD + (P1W * 4) + (P1W * 6)
                    End Select
                End If
                RECS.Update
            Next fnx
        Case 9
            XNO = M1W
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![�ؒf�{��] = 2 * 2
                RECS![���@1] = STAD
                Select Case fnx
                    Case 3 To 5
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8, 9
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 11
            XNO = M1W
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![�ؒf�{��] = 2 * 2
                RECS![���@1] = STAD
                Select Case fnx
                    Case 3, 4
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                   Case 8, 9
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                   Case 10, 11
                        RECS![����] = 4
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![���@5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
        
    End Select
End Sub
Public Sub S1_NS36()
    Select Case M1W
        Case 5
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 4, 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 5)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * M1W)
                        RECS![�ؒf�{��] = (N1W - M1W) * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 7
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 6)
                    Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * M1W)
                        RECS![�ؒf�{��] = (N1W - M1W) * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
        Case 9
            XNO = 10
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8, 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                    Case 10
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * M1W)
                        RECS![�ؒf�{��] = (N1W - M1W) * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx

        Case 11
            XNO = 12
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8, 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 10, 11
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 4
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![���@5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                   Case 12
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * M1W)
                        RECS![�ؒf�{��] = (N1W - M1W) * 2
                        RECS![����] = 4
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![���@5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
       
    End Select
End Sub
Public Sub S1_NS37()
    Select Case M1W
        Case 5
            Call S1_NS37_M5
        Case 7
            Call S1_NS37_M7
        Case 9
            Call S1_NS37_M9
        Case 11
            Call S1_NS37_M11
    End Select
End Sub
Public Sub S1_NS37_M5()
    Select Case N1W
        Case 2
            XNO = 4
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 2)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 4
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 5
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 5)
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
            
    End Select
End Sub
Public Sub S1_NS37_M7()
    Select Case N1W
        Case 2
            XNO = 4
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 4 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 5
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + P1W
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3, 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                   Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + (P1W * 7)     '04/02/01
                End Select
                
                RECS.Update
            Next fnx
        Case 5
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
        Case 6
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 6)
                    Case 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
            
    End Select
End Sub
Public Sub S1_NS37_M9()
    Select Case N1W
        Case 2
            XNO = 5
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 4 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 2)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + P1W
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3, 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 5)
                    Case 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
               End Select
                
                RECS.Update
            Next fnx
        Case 5
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 5)
                    Case 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 6
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 5)
                   Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 5)
                        RECS![���@4] = STAD + P1W + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 7
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�����}��] = STEN
                RECS![�̔�no] = wno
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 5)
                        RECS![���@4] = STAD + (P1W * 3) + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 8
            XNO = 9
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                    Case 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 5)
                        RECS![���@4] = STAD + (P1W * 3) + (P1W * 5) + (P1W * 5)
               End Select
                
                RECS.Update
            Next fnx
            
    End Select
End Sub
Public Sub S1_NS37_M11()
    Select Case N1W
        Case 2
            XNO = 4
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 8 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 4 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 4)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + P1W
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3, 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 4 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 4)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 3)
                   Case 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 5
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 4)
                    Case 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 4)
                    Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
        Case 6
            XNO = 9
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5, 6
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 4)
                    Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 4)
                        RECS![���@4] = STAD + P1W + (P1W * 4) + (P1W * 6)
                    Case 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 7
            XNO = 9
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                    Case 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 4)
                        RECS![���@4] = STAD + P1W + (P1W * 4) + (P1W * 6)
                End Select
           
                RECS.Update
            Next fnx
        Case 8
            XNO = 10
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                    Case 10
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 4
                        RECS![���@2] = STAD + P1W
                        RECS![���@3] = STAD + P1W + (P1W * 4)
                        RECS![���@4] = STAD + P1W + (P1W * 4) + (P1W * 6)
                        RECS![���@5] = STAD + P1W + (P1W * 4) + (P1W * 6) + (P1W * 4)
               End Select
                
                RECS.Update
            Next fnx
            
        Case 9
            XNO = 10
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8 To 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 10
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 4
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                        RECS![���@5] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                End Select
           
                RECS.Update
            Next fnx
        Case 10
            XNO = 11
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![�����ԍ�] = STNO
                RECS![�̔�no] = wno
                RECS![�����}��] = STEN
                RECS![DT�敪] = STDT
                RECS![�i�q�ԍ�] = fnx
                
                RECS![���@1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                    Case 3 To 4
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 1
                        RECS![���@2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 2
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8, 9
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 3
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 10
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![�ؒf�{��] = 2 * 2
                        RECS![����] = 4
                        RECS![���@2] = STAD + (P1W * 4)
                        RECS![���@3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![���@5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                    Case 11
                        RECS![�ؒf���@] = STAD * 2 + P1W * (2 * N1W)
                        RECS![�ؒf�{��] = 1 * 2
                        RECS![����] = 4
                        RECS![���@2] = STAD + (P1W * 3)
                        RECS![���@3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![���@4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                        RECS![���@5] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
            
    End Select
End Sub


Function get_sho(Sho As String, hen) '�󒍗���p�@�i2015/7/16
    Select Case Sho
     Case "XX"
      get_sho = "A�C�i�o"
     Case "P1", "P1A"
      get_sho = "A�N���X"
     Case "P2"
      get_sho = "A�T��"
     Case "P3", "P3A"
      get_sho = "A�N���A"
     Case "P4", "P4A"
      get_sho = "A�N���A"
     Case "P5"
      get_sho = "A�䌅"
     Case "HS3"
      get_sho = "A�x��"
     Case "HT4"
        If hen = "PD11" Then
          get_sho = "AHIT�x��"
        ElseIf hen = "PD21" Then
          get_sho = "CHIT�x��"
        End If
     Case "SM1"
      get_sho = "AS�^"

     Case "HS1"
      get_sho = "CHS�N"
     Case "HS2"
      get_sho = "CHS�G"
     Case "HS4"
      get_sho = "CHWP"
     Case "HS5"
      get_sho = "CHW��"
     Case "HS6"
      get_sho = "CHW��"
     Case "HS7"
      get_sho = "CHS��"
     Case "HT1"
      get_sho = "CHIT�N"
     Case "HT2"
      get_sho = "CHIT�G"
     Case "HT3"
      get_sho = "CHIT��"
     Case "KK1"
      get_sho = "C���ǃN"
     Case "KK2"
      get_sho = "C���ǒG"
     Case "KK3"
      get_sho = "C���ǉ�"
     Case "KK4"
      get_sho = "A���ǎx"
     Case "NS1"
      get_sho = "C2�ӃN"
     Case "NS2"
      get_sho = "C2�ӎx"
     Case "FT1"
      get_sho = "BEX��"

     Case Else
        If Sho Like "S*" Then
            get_sho = "A��"
        ElseIf Sho Like "M*" Then
            get_sho = "C���C�N"
        ElseIf Sho Like "G*" Then
            get_sho = "C��"
        End If
    End Select

End Function

Function get_sho2(Sho As String, hen As Variant, �����敪 As String)   '�󒍏󋵉����p�@�i2018/03/02)
    If �����敪 = "A" Then
        Select Case Sho
         Case "XX"
          get_sho2 = "A�C�i�o"
         Case "P1", "P1A"
          get_sho2 = "A�N���X"
         Case "P2"
          get_sho2 = "A�T��"
         Case "P3", "P3A"
          get_sho2 = "A�N���A"
         Case "P4", "P4A"
          get_sho2 = "A�N���A"
         Case "P5"
          get_sho2 = "A�䌅"
         Case "HS3"
          get_sho2 = "A�x��"
         Case "HT4"
            If hen = "PD11" Then
              get_sho2 = "AHIT�x��"
            ElseIf hen = "PD21" Then
              get_sho2 = "CHIT�x��"
            End If
         Case "SM1"
          get_sho2 = "AS�^"
    
         Case "HS1"
          get_sho2 = "CHS�N"
         Case "HS2"
          get_sho2 = "CHS�G"
         Case "HS4"
          get_sho2 = "CHWP"
         Case "HS5"
          get_sho2 = "CHW��"
         Case "HS6"
          get_sho2 = "CHW��"
         Case "HS7"
          get_sho2 = "CHS��"
         Case "HT1"
          get_sho2 = "CHIT�N"
         Case "HT2"
          get_sho2 = "CHIT�G"
         Case "HT3"
          get_sho2 = "CHIT��"
         Case "KK1"
          get_sho2 = "C���ǃN"
         Case "KK2"
          get_sho2 = "C���ǒG"
         Case "KK3"
          get_sho2 = "C���ǉ�"
         Case "KK4"
          get_sho2 = "A���ǎx"
         Case "NS1"
          get_sho2 = "C2�ӃN"
         Case "NS2"
          get_sho2 = "C2�ӎx"
         Case "FT1"
          get_sho2 = "BEX��"
    
         Case Else
            If Sho Like "S*" Then
                get_sho2 = "A��"
            ElseIf Sho Like "M*" Then
                get_sho2 = "C���C�N"
            ElseIf Sho Like "G*" Then
                get_sho2 = "C��"
            End If
        End Select
   Else
        Select Case Sho
         Case "XX"
          get_sho2 = "A�K�i"
         Case "P1", "P1A"
          get_sho2 = "A�K�i"
         Case "P2"
          get_sho2 = "A�K�i"
         Case "P3", "P3A"
          get_sho2 = "A�K�i"
         Case "P4", "P4A"
          get_sho2 = "A�K�i"
         Case "P5"
          get_sho2 = "A�K�i"
         Case "HS3"
          get_sho2 = "A�K�i"
         Case "HT4"
            If hen = "PD11" Then
              get_sho2 = "A�K�i"
            ElseIf hen = "PD21" Then
              get_sho2 = "C�K�i"
            End If
         Case "SM1"
          get_sho2 = "A�K�i"
    
         Case "HS1"
          get_sho2 = "C�K�i"
         Case "HS2"
          get_sho2 = "C�K�i"
         Case "HS4"
          get_sho2 = "C�K�i"
         Case "HS5"
          get_sho2 = "C�K�i"
         Case "HS6"
          get_sho2 = "C�K�i"
         Case "HS7"
          get_sho2 = "C�K�i"
         Case "HT1"
          get_sho2 = "C�K�i"
         Case "HT2"
          get_sho2 = "C�K�i"
         Case "HT3"
          get_sho2 = "C�K�i"
         Case "KK1"
          get_sho2 = "C�K�i"
         Case "KK2"
          get_sho2 = "C�K�i"
         Case "KK3"
          get_sho2 = "C�K�i"
         Case "KK4"
          get_sho2 = "A�K�i"
         Case "NS1"
          get_sho2 = "C�K�i"
         Case "NS2"
          get_sho2 = "C�K�i"
         Case "FT1"
          get_sho2 = "BEX��"
    
         Case Else
            If Sho Like "S*" Then
                get_sho2 = "A�K�i"
            ElseIf Sho Like "M*" Then
                get_sho2 = "C�K�i"
            ElseIf Sho Like "G*" Then
                get_sho2 = "C�K�i"
            End If
        End Select
    End If
End Function

