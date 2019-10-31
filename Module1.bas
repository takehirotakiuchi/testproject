Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public str As String            '文字列ワーク
Public nname As String          'コンピューター名
Public ret                      'メッセージ　リターン
Public strSQL As String         'sql用
Public s_date As String         'sql用
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
'売上分類　　　stp:商品ﾀｲﾌﾟ uriage:売上分類  ST:ST(会社分類)
    
get_uriage = uriage

'ヨコ
If Stp = "S3" Then
    If ST = "S" Then
        get_uriage = "D00"
    Else
        get_uriage = "Y00"
    End If
End If

'マス
If Stp = "S4" Then
    If ST = "S" Then
        get_uriage = "E00"
    Else
        get_uriage = "Z00"
    End If
End If

End Function
Function get_ri(YN As Integer, YM As Integer, CD As String) As Integer
'リベット数計算　　　yn:YN ym:YM cd:商品タイプCD
    Dim i As Integer
    Dim K As Integer
    Dim Nmas As Integer
    Dim Mmas As Integer
    Dim Bai, Itiretume As Integer

    i = YN * 2
    K = YM * 2


    
    If CD = "S3" Or CD = "S3A" Then
        get_ri = K - 2 '横格子
        
    ElseIf CD = "S1A" Or CD = "S2A" Then '新ｸﾛｽ格子
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
    
    ElseIf CD = "P3" Then 'ﾊﾟﾅｸﾘｱ中桟有
        get_ri = YN * 3 + 4
    ElseIf CD = "P3A" Then 'ﾊﾟﾅｸﾘｱ中桟有
        get_ri = YN * 3 + 4

    ElseIf CD = "P4" Then 'ﾊﾟﾅｸﾘｱ中桟無し
        get_ri = YN * 2
    ElseIf CD = "P4A" Then 'ﾊﾟﾅｸﾘｱ中桟無し
        get_ri = YN * 2

    ElseIf CD = "P5" Then   '枡格子
                
        Nmas = YN - 1 '縦格子本数
        Mmas = YM - 1 '横格子本数(上下端除く）
        
        get_ri = Nmas * 2 '上下端用リベット
        
        Bai = kirishute(Nmas / 2, 0)
        Itiretume = kirishute(Mmas / 2, 0)

        Select Case Nmas Mod 2
            Case 0
                get_ri = get_ri + Mmas * Bai
            Case 1
                get_ri = get_ri + Itiretume + Mmas * Bai
        End Select
    
    Else '旧ｸﾛｽ格子
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
        get_ln2 = "B"    '2018/02/02 Bラインに変更
    End If
    
    If tyu = "F" Then
        get_ln2 = "F"
    End If

End Function

Function get_ln(Stp As String, rbn As Integer, rym As Integer, Hno As String, MH As Integer, Uri As String, tyu As String) As String
'Function get_ln(Stp As String, Hno As String, tyu As String) As String
'ライン計算　　　stp:商品ﾀｲﾌﾟ  rbn:ﾘﾍﾞｯﾄ数  rym:YM  Hno:変数グループ    MH:MH     URI:売上分類
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
        get_ln = "B"    '2018/02/02 Bラインに変更
    End If
    
    If tyu = "F" Then
        get_ln = "F"
    End If

End Function

Function get_bcd(kbn As Variant, cbn As Variant, sck As Variant, did As Variant) As Variant
    'ﾊﾞｰｺｰﾄﾞ編集 kbn:会社区分,cbn:注文番号,sck:出荷倉庫,did:ﾃﾞｰﾀID
    If did = "A0" Then
        '受注
        If kbn = "T" Then
            '立山
            get_bcd = "ZZZ" & Mid(cbn, 1, 10) & "  "
        Else
            '三協
            get_bcd = "ZZZ" & cbn
        End If
    Else
        '規格
        If Mid(sck, 1, 1) = "T" Then
            '立山
            get_bcd = "ZZZYB" & Mid(cbn, 1, 8) & "  "
        Else
            '三協
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

Function shisya(dbl対象 As Double, lng桁数 As Long) As Double
    '四捨五入
    '一の位は-1、十の位は-2、百の位は-3、
    '小数点第１位は0、第2位は1、第3位は2、
    Dim lng数値 As Long
    lng数値 = 10 ^ Abs(lng桁数)
    
    If lng桁数 > 0 Then
        shisya = Int(dbl対象 * lng数値 + 0.555555) / lng数値
    Else
        shisya = Int(dbl対象 / lng数値 + 0.555555) * lng数値
    End If
End Function

Function shisya2(dbl対象 As Double, lng桁数 As Long) As Double
    '四捨五入
    '一の位は-1、十の位は-2、百の位は-3、
    '小数点第１位は0、第2位は1、第3位は2、
    Dim lng数値 As Long
    lng数値 = 10 ^ Abs(lng桁数)
    
    If lng桁数 > 0 Then
        shisya2 = Int(dbl対象 * lng数値 + 0.5) / lng数値
    Else
        shisya2 = Int(dbl対象 / lng数値 + 0.5) * lng数値
    End If
End Function

Function kirishute(dbl対象 As Double, lng桁数 As Long) As Double
    '切り捨て
    '一の位は-1、十の位は-2、百の位は-3、
    '小数点第１位は0、第2位は1、第3位は2、

    Dim lng数値 As Long
    lng数値 = 10 ^ Abs(lng桁数)

    If lng桁数 > 0 Then
        kirishute = Int(dbl対象 * lng数値) / lng数値
    Else
        kirishute = Int(dbl対象 / lng数値) * lng数値
    End If
End Function

Function kiriage(dbl対象 As Double, lng桁数 As Long) As Double
    '切り上げ
    '一の位は-1、十の位は-2、百の位は-3、
    '小数点第１位は0、第2位は1、第3位は2、

    Dim lng数値 As Long
    lng数値 = 10 ^ Abs(lng桁数)

    If lng桁数 > 0 Then
        
        If Int(dbl対象 * lng数値) = (dbl対象 * lng数値) Then
            kiriage = Int(dbl対象 * lng数値) / lng数値
        Else
            kiriage = (Int(dbl対象 * lng数値) + 1) / lng数値
        End If
    Else
        If Int(dbl対象 / lng数値) = (dbl対象 / lng数値) Then
            kiriage = Int(dbl対象 / lng数値) * lng数値
        Else
            kiriage = (Int(dbl対象 / lng数値) + 1) * lng数値
        End If
    End If
End Function

Function Odr_Add()
    '発注データの追加 受注データの存在チェック
    Dim DB As Database
    Dim rst As Recordset
    Dim rsd As Recordset

    Dim trc As Long
    Dim sds As String

    trc = DCount("[注文番号]", "SAN")
    If trc = 0 Then
        Exit Function
    End If

    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("SAN累積", dbOpenDynaset)
    Set rsd = DB.OpenRecordset("SAN")

    rsd.MoveFirst
    Do Until rsd.EOF

        sds = rsd![注文番号]
        rst.FindFirst "[注文番号]='" & sds & "'"

        ' NoMatch プロパティの値に基づいて戻り値を設定します。
        If rst.NoMatch = False Then
            rsd.Delete
        End If
        rsd.MoveNext
    Loop
    rsd.Close: Set rsd = Nothing
    rst.Close: Set rst = Nothing


    trc = DCount("[注文番号]", "q_パナ特注")
    If trc = 0 Then
        Exit Function
    End If
    Set rst = DB.OpenRecordset("q_パナ特注", dbOpenDynaset)
    rst.MoveFirst
    Do Until rst.EOF
        rst.Edit
        '寸法値1
        If rst![寸法記号1] = "MW" And rst![基本タイプコード] Like "M092D01*" Then
            rst![寸法値1] = rst![寸法値1] - 480
            rst![寸法記号1] = "W"
        ElseIf rst![寸法記号1] = "MW" And rst![基本タイプコード] Like "M092D02*" Then
            rst![寸法値1] = rst![寸法値1] - 920
            rst![寸法記号1] = "W"
        ElseIf rst![寸法記号1] = "MW" Then
            rst![寸法値1] = rst![寸法値1] - 820
            rst![寸法記号1] = "W"
        ElseIf rst![寸法記号1] = "W" Then
            rst![寸法値1] = rst![寸法値1]

        ElseIf rst![寸法記号1] = "MH" And rst![基本タイプコード] Like "M092D01*" Then
            rst![寸法値1] = rst![寸法値1] - 160
            rst![寸法記号1] = "H"
        ElseIf rst![寸法記号1] = "MH" And rst![基本タイプコード] Like "M092D02*" Then
            rst![寸法値1] = rst![寸法値1] - 330
            rst![寸法記号1] = "H"
        ElseIf rst![寸法記号1] = "MH" Then
            rst![寸法値1] = rst![寸法値1] - 600
            rst![寸法記号1] = "H"
        ElseIf rst![寸法記号1] = "H" Then
            rst![寸法値1] = rst![寸法値1]
        Else
            rst![寸法値1] = 0
        End If

        '寸法値2
        If rst![寸法記号2] = "MW" And rst![基本タイプコード] Like "M092D01*" Then
            rst![寸法値2] = rst![寸法値2] - 480
            rst![寸法記号2] = "W"
        ElseIf rst![寸法記号2] = "MW" And rst![基本タイプコード] Like "M092D02*" Then
            rst![寸法値2] = rst![寸法値2] - 920
            rst![寸法記号2] = "W"
        ElseIf rst![寸法記号2] = "MW" Then
            rst![寸法値2] = rst![寸法値2] - 820
            rst![寸法記号2] = "W"
        ElseIf rst![寸法記号2] = "W" Then
            rst![寸法値2] = rst![寸法値2]

        ElseIf rst![寸法記号2] = "MH" And rst![基本タイプコード] Like "M092D01*" Then
            rst![寸法値2] = rst![寸法値2] - 160
            rst![寸法記号2] = "H"
        ElseIf rst![寸法記号2] = "MH" And rst![基本タイプコード] Like "M092D02*" Then
            rst![寸法値2] = rst![寸法値2] - 330
            rst![寸法記号2] = "H"
        ElseIf rst![寸法記号2] = "MH" Then
            rst![寸法値2] = rst![寸法値2] - 600
            rst![寸法記号2] = "H"
        ElseIf rst![寸法記号2] = "H" Then
            rst![寸法値2] = rst![寸法値2]
        Else
            rst![寸法値2] = 0
        End If

        rst.Update
        rst.MoveNext
    Loop
    rst.Close: Set rst = Nothing
    DB.Close: Set DB = Nothing
End Function
Function Odr7_Add()
    
    '発注データの追加 受注データの存在チェック
    Dim DB As Database
    Dim rst As Recordset
    Dim rsd As Recordset
    
    Dim trc As Long
    Dim sds As String
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("SAN累積", dbOpenDynaset)
    Set rsd = DB.OpenRecordset("SANA7")
    
    trc = DCount("[注文番号]", "SANA7")
    If trc = 0 Then
        GoTo nods
    End If
    
    rsd.MoveFirst
    Do Until rsd.EOF
                    
        sds = rsd![注文番号]
        rst.FindFirst "[注文番号]='" & sds & "'"
    
        ' NoMatch プロパティの値に基づいて戻り値を設定します。
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
Function Ka_Add() '追加用　（2013/1/29現在　使用していない）
    
    'かんばんデータ作成
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
    
    'DoCmd.OpenQuery "d_YOTEIﾜｰｸ"
    'DoCmd.OpenQuery "d_格子部ﾜｰｸ"
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("j_YOTEI追加")
    Set rSM = DB.OpenRecordset("m_変数", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ﾜｰｸno")
    Set rso = DB.OpenRecordset("YOTEIﾜｰｸ")
    Set rsk = DB.OpenRecordset("m_工作図番号", dbOpenDynaset)
        
    Set DBB = CurrentDb
    Set RECS = DBB.OpenRecordset("格子部ﾜｰｸ")

    trc = DCount("[社内管理番号1]", "j_YOTEI追加")
    If trc = 0 Then
        GoTo nodu
    End If
    
    rst.MoveFirst
    Do Until rst.EOF
            
            rst.Edit
                 
            rsn.MoveFirst
            wno = rsn![採番no]
            wno = wno + 1
            rsn.Edit
            rsn![採番no] = wno
            rsn.Update
        
            rso.AddNew
            rso![注文番号] = rst![社内管理番号1] & rst![社内管理番号2]
            rso![注文枝番] = 0
            rso![採番no] = wno
            rso![DT区分] = "0"
            rso![注文区分] = ""
            rso![受注区分] = ""
            rso![社内管理番号1] = rst![社内管理番号1]
            rso![社内管理番号2] = rst![社内管理番号2]
            rso![納期日] = rst![納期日]
            rso![投入予定ｾｯﾄ数] = rst![投入予定ｾｯﾄ数]
           
            If rst![T使用] = False Then
                rso![H寸法値] = rst![TH]
                rso![W寸法値] = rst![TW]
            Else
                rso![H寸法値] = rst![H寸]
                rso![W寸法値] = rst![W寸]
            End If
            
            rso![商品ｺｰﾄﾞ] = rst![商品ｺｰﾄﾞ]
            rso![COL] = rst![色]
            rso![商品ﾀｲﾌﾟCD] = rst![商品ﾀｲﾌﾟCD]
            rso![変数ｸﾞﾙｰﾌﾟ] = rst![変数ｸﾞﾙｰﾌﾟ]
            rso![上枠工作図番号] = rst![上枠工作図番号]
            ksz = rst![上枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![上枠型番] = rsk![型番]
            
            rso![下枠工作図番号] = rst![下枠工作図番号]
            ksz = rst![下枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![下枠型番] = rsk![型番]
            
            rso![竪枠工作図番号] = rst![竪枠工作図番号]
            ksz = rst![竪枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![竪枠型番] = rsk![型番]
            
            rso![格子部工作図番号] = rst![格子部工作図番号]
            ksz = rst![格子部工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![格子部型番] = rsk![型番]
            
            hds = rst![変数ｸﾞﾙｰﾌﾟ]
            rSM.FindFirst "[変数ｸﾞﾙｰﾌﾟ]='" & hds & "'"
            
            If rSM![MW分割] = True Then
                rso![YMW] = (rso![W寸法値] / 2) + rSM![MW用変数]
            Else
                rso![YMW] = rso![W寸法値] + rSM![MW用変数]
            End If
            
            rso![YMH] = rso![H寸法値] + rSM![MH用変数]
            
            
            rso![上下枠切断寸法] = rso![YMW] + rSM![上枠下枠寸法用]
            rso![竪枠切断寸法] = rso![YMH] + rSM![竪枠寸法用]
            
            If rst![商品ﾀｲﾌﾟCD] = "S3" Then
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
            End If
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![格子竪切断寸法] = rso![YMH] + rSM![格子竪寸法用]
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
                STTD = rso![格子竪切断寸法]
            End If
            
            If rst![商品ﾀｲﾌﾟCD] = "XX" Then
                rso![格子竪切断寸法] = rso![YMH] - 50.5
                STYD = rso![格子竪切断寸法]
            End If
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
               dobn1 = rso![YPW] * rso![YPW]
               dobn2 = rso![YPH] * rso![YPH]
               dobn3 = dobn1 + dobn2
               dobn4 = Sqr(dobn3)
               dobn = dobn4 / 2
               rso![YP1] = dobn
               P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![上枠下枠格子取付ﾋﾟｯﾁ]
            If rst![商品ﾀｲﾌﾟCD] = "S3" Or rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![竪枠格子取付ﾋﾟｯﾁ]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![竪枠格子取付ﾋﾟｯﾁ]
            End If
            
            STNO = rso![注文番号]
            STEN = rso![注文枝番]
            STDT = rso![DT区分]
            STAD = rSM![格子AD値]
            
            If rst![商品ﾀｲﾌﾟCD] = "S2" Then
                rso![FS区分] = True
            End If
            
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If

            If rst![商品ﾀｲﾌﾟCD] = "XX" Then
                If rso![上下枠切断寸法] = 1193 Then
                    rso![YN] = 9
                    rso![YPW] = 117
                    N1W = rso![YN]
                Else
                    rso![YN] = 14
                    rso![YPW] = 117
                    N1W = rso![YN]
                End If
            End If
            
            rso![追加ﾌﾗｸﾞ] = True
 
            rso![リベット] = get_ri(rso![YN], rso![YM], rso![商品ﾀｲﾌﾟCD])
            
            rso.Update
           
            '格子部データの出力
            Select Case rst![商品ﾀｲﾌﾟCD]
                Case "S1"
                    Call S1_set     'クロス
                Case "S2"
                    Call S1_set     'クロス　飾窓
                Case "S3"
                    Call S3_set     '横
                Case "S4"
                    Call S4_set     '枡
                Case "XX"
                    Call XX_set     '稲葉製作所向け
            End Select
            
            If rst![商品ﾀｲﾌﾟCD] = "S2" Then
        
            rst.Edit
                 
            rsn.MoveFirst
            wno = rsn![採番no]
            wno = wno + 1
            rsn.Edit
            rsn![採番no] = wno
            rsn.Update
        
            rso.AddNew
            rso![注文番号] = rst![社内管理番号1] & rst![社内管理番号2]
            rso![注文枝番] = 0
            rso![採番no] = wno
            rso![DT区分] = "0"
            rso![注文区分] = ""
            rso![受注区分] = ""
            rso![社内管理番号1] = rst![社内管理番号1]
            rso![社内管理番号2] = rst![社内管理番号2]
            rso![納期日] = rst![納期日]
            rso![投入予定ｾｯﾄ数] = rst![投入予定ｾｯﾄ数]
           
            If rst![T使用] = False Then
                rso![H寸法値] = rst![TH2]
                rso![W寸法値] = rst![TW2]
            Else
                rso![H寸法値] = rst![H寸2]
                rso![W寸法値] = rst![W寸2]
            End If
            
            rso![商品ｺｰﾄﾞ] = rst![商品ｺｰﾄﾞ]
            rso![COL] = rst![色]
            rso![商品ﾀｲﾌﾟCD] = rst![商品ﾀｲﾌﾟCD]
            rso![変数ｸﾞﾙｰﾌﾟ] = rst![変数ｸﾞﾙｰﾌﾟ]
            rso![上枠工作図番号] = rst![上枠工作図番号]
            ksz = rst![上枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![上枠型番] = rsk![型番]
            
            rso![下枠工作図番号] = rst![下枠工作図番号]
            ksz = rst![下枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![下枠型番] = rsk![型番]
            
            rso![竪枠工作図番号] = rst![竪枠工作図番号]
            ksz = rst![竪枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![竪枠型番] = rsk![型番]
            
            rso![格子部工作図番号] = rst![格子部工作図番号]
            ksz = rst![格子部工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![格子部型番] = rsk![型番]
            
            hds = rst![変数ｸﾞﾙｰﾌﾟ2]
            rSM.FindFirst "[変数ｸﾞﾙｰﾌﾟ]='" & hds & "'"
            
            If rSM![MW分割] = True Then
                rso![YMW] = (rso![W寸法値] / 2) + rSM![MW用変数]
            Else
                rso![YMW] = rso![W寸法値] + rSM![MW用変数]
            End If
            
            rso![YMH] = rso![H寸法値] + rSM![MH用変数]
            
            rso![上下枠切断寸法] = rso![YMW] + rSM![上枠下枠寸法用]
            rso![竪枠切断寸法] = rso![YMH] + rSM![竪枠寸法用]

            
            If rst![商品ﾀｲﾌﾟCD] = "S3" Then
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
            End If
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![格子竪切断寸法] = rso![YMH] + rSM![格子竪寸法用]
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
                STTD = rso![格子竪切断寸法]
            End If
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn1 = rso![YPW] * rso![YPW]
                dobn2 = rso![YPH] * rso![YPH]
                dobn3 = dobn1 + dobn2
                dobn4 = Sqr(dobn3)
                dobn = dobn4 / 2
                rso![YP1] = dobn
                P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![上枠下枠格子取付ﾋﾟｯﾁ]
            If rst![商品ﾀｲﾌﾟCD] = "S3" Or rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![竪枠格子取付ﾋﾟｯﾁ]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![竪枠格子取付ﾋﾟｯﾁ]
            End If
            
            STNO = rso![注文番号]
            STEN = rso![注文枝番]
            STDT = rso![DT区分]
            STAD = rSM![格子AD値]
            
            If rst![商品ﾀｲﾌﾟCD] = "S2" Then
                rso![FS区分] = False
            End If
            
            
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If
            
            rso![追加ﾌﾗｸﾞ] = True
 
            rso.Update
           
            Call S1_set     'クロス　飾窓
            
            End If
            
'            rst.Edit
'            rst![作成区分] = True
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
    'かんばんデータ作成
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
    Set rsd = DB.OpenRecordset("q_SAN累積", dbOpenDynaset)
    Set rSM = DB.OpenRecordset("m_変数", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ﾜｰｸno")
    Set rso = DB.OpenRecordset("YOTEIﾜｰｸ")
    Set rsk = DB.OpenRecordset("m_工作図番号", dbOpenDynaset)
        
    Set DBB = CurrentDb
    trc = DCount("[注文番号]", "j_YOTEI_GP")
    If trc = 0 Then
        GoTo nodt
    End If

    rst.MoveFirst
    Do Until rst.EOF

        sds = rst![注文番号]
        rsd.FindFirst "[注文番号]='" & sds & "'"
    
        ' NoMatch プロパティの値に基づいて戻り値を設定します。
        If rsd.NoMatch = False Then
        
            rsn.MoveFirst
            wno = rsn![採番no]
            wno = wno + 1
            rsn.Edit
            rsn![採番no] = wno
            rsn.Update
        
            rso.AddNew
            rso![注文番号] = rst![注文番号]
            rso![注文枝番] = rst![注文枝番]
            rso![COL] = rst![色]
            
            rso![採番no] = wno
            rso![DT区分] = "0"
            rso![注文区分] = rst![注文区分]
            rso![受注区分] = rst![受注区分]
            
            rso![社内管理番号1] = rst![社内管理番号1]
            kot = "0000" + rst![社内管理番号2]
            rso![社内管理番号2] = Right(kot, 4)
            rso![納期日] = rst![納期日]
            rso![投入予定ｾｯﾄ数] = rst![投入予定ｾｯﾄ数]
            rso![数量] = rst![数量]
           
            kigou(0) = rsd![寸法記号1]
            kigou(1) = rsd![寸法記号2]
            kigou(2) = rsd![寸法記号3]
            kigou(3) = rsd![寸法記号4]
            kigou(4) = rsd![寸法記号5]
            kigou(5) = rsd![寸法記号6]
            kigou(6) = rsd![寸法記号7]
            kigou(7) = rsd![寸法記号8]
            sunpou(0) = rsd![寸法値1]
            sunpou(1) = rsd![寸法値2]
            sunpou(2) = rsd![寸法値3]
            sunpou(3) = rsd![寸法値4]
            sunpou(4) = rsd![寸法値5]
            sunpou(5) = rsd![寸法値6]
            sunpou(6) = rsd![寸法値7]
            sunpou(7) = rsd![寸法値8]
            
            If rst![特寸ﾏｰｸ] = "1" Then
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H2" Then
                        rso![H寸法値] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
                
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W2" Then
                        rso![W寸法値] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
            Else
                rso![H寸法値] = rst![H寸]
                rso![W寸法値] = rst![W寸]
            End If
                
            
            rso![商品ｺｰﾄﾞ] = rst![商品ｺｰﾄﾞ]
            rso![商品ﾀｲﾌﾟCD] = rst![商品ﾀｲﾌﾟCD]
            rso![変数ｸﾞﾙｰﾌﾟ] = rst![変数ｸﾞﾙｰﾌﾟ]
            
            rso![上枠工作図番号] = rst![上枠工作図番号]
            ksz = rst![上枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![上枠型番] = rsk![型番]
            
            rso![下枠工作図番号] = rst![下枠工作図番号]
            ksz = rst![下枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![下枠型番] = rsk![型番]
            
            rso![竪枠工作図番号] = rst![竪枠工作図番号]
            ksz = rst![竪枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![竪枠型番] = rsk![型番]
            
            Select Case rst![商品ﾀｲﾌﾟCD]
             Case "GP4"
                If rso![H寸法値] >= 2000 And rso![H寸法値] < 2245 Then
                    rso![YN] = 8
                ElseIf rso![H寸法値] >= 2245 And rso![H寸法値] < 2945 Then
                    rso![YN] = 10
                ElseIf rso![H寸法値] >= 2945 And rso![H寸法値] < 3645 Then
                    rso![YN] = 12
                ElseIf rso![H寸法値] >= 3645 And rso![H寸法値] < 3800 Then
                    rso![YN] = 14
                End If
                rso![YPH] = 700
             Case "GP7"
                rso![YN] = 4
                rso![YPH] = 700
            End Select
            
            rso.Update
        
'            rst.Edit
'            rst![ﾃﾞｰﾀ作成区分] = True
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

    'かんばんデータ作成
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
    Set rsd = DB.OpenRecordset("q_SAN累積", dbOpenDynaset)
    Set rSM = DB.OpenRecordset("m_変数", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ﾜｰｸno")
    Set rso = DB.OpenRecordset("YOTEIﾜｰｸ")
    Set rsk = DB.OpenRecordset("m_工作図番号", dbOpenDynaset)
        
    Set DBB = CurrentDb
    Set RECS = DBB.OpenRecordset("格子部ﾜｰｸ")

    trc = DCount("[注文番号]", "j_YOTEI")
    If trc = 0 Then
        GoTo nodt
    End If

    rst.MoveFirst
    Do Until rst.EOF

        sds = rst![注文番号]
        rsd.FindFirst "[注文番号]='" & sds & "'"
    
        ' NoMatch プロパティの値に基づいて戻り値を設定します。
        If rsd.NoMatch = False Then
        
            rsn.MoveFirst
            wno = rsn![採番no]
            wno = wno + 1
            rsn.Edit
            rsn![採番no] = wno
            rsn.Update
        
            rso.AddNew
            rso![注文番号] = rst![注文番号]
            rso![注文枝番] = rst![注文枝番]
            
            If rst![商品コード] = "Sangutte" Then
                rso![COL] = rst![特寸色]
            Else
                rso![COL] = rst![色]
            End If
            
            rso![採番no] = wno
            rso![DT区分] = "0"
            rso![注文区分] = rst![注文区分]
            rso![受注区分] = rst![受注区分]
            
            rso![社内管理番号1] = rst![社内管理番号1]
            kot = "0000" + rst![社内管理番号2]
            rso![社内管理番号2] = Right(kot, 4)
            
            rso![納期日] = rst![納期日]
            rso![投入予定ｾｯﾄ数] = rst![投入予定ｾｯﾄ数]
            rso![数量] = rst![数量]
           
            kigou(0) = rsd![寸法記号1]
            kigou(1) = rsd![寸法記号2]
            kigou(2) = rsd![寸法記号3]
            kigou(3) = rsd![寸法記号4]
            kigou(4) = rsd![寸法記号5]
            kigou(5) = rsd![寸法記号6]
            kigou(6) = rsd![寸法記号7]
            kigou(7) = rsd![寸法記号8]
            sunpou(0) = rsd![寸法値1]
            sunpou(1) = rsd![寸法値2]
            sunpou(2) = rsd![寸法値3]
            sunpou(3) = rsd![寸法値4]
            sunpou(4) = rsd![寸法値5]
            sunpou(5) = rsd![寸法値6]
            sunpou(6) = rsd![寸法値7]
            sunpou(7) = rsd![寸法値8]
            
            If rst![特寸ﾏｰｸ] = "1" Then
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H2" Then
                        rso![H寸法値] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
                
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W2" Then
                        rso![W寸法値] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
            Else
                rso![H寸法値] = rst![H寸]
                rso![W寸法値] = rst![W寸]
            End If
            
            rso![商品ｺｰﾄﾞ] = rst![商品ｺｰﾄﾞ]
            rso![商品ﾀｲﾌﾟCD] = rst![商品ﾀｲﾌﾟCD]
            rso![変数ｸﾞﾙｰﾌﾟ] = rst![変数ｸﾞﾙｰﾌﾟ]
            rso![上枠工作図番号] = rst![上枠工作図番号]
            ksz = rst![上枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![上枠型番] = rsk![型番]
            
            rso![下枠工作図番号] = rst![下枠工作図番号]
            ksz = rst![下枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![下枠型番] = rsk![型番]
            
            rso![竪枠工作図番号] = rst![竪枠工作図番号]
            ksz = rst![竪枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![竪枠型番] = rsk![型番]
            
            rso![格子部工作図番号] = rst![格子部工作図番号]
            ksz = rst![格子部工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![格子部型番] = rsk![型番]
            
            hds = rst![変数ｸﾞﾙｰﾌﾟ]
            rSM.FindFirst "[変数ｸﾞﾙｰﾌﾟ]='" & hds & "'"
            
            If rSM![MW分割] = True Then
                rso![YMW] = (rso![W寸法値] / 2) + rSM![MW用変数]
            Else
                rso![YMW] = rso![W寸法値] + rSM![MW用変数]
            End If
            
            rso![YMH] = rso![H寸法値] + rSM![MH用変数]
            
            rso![上下枠切断寸法] = rso![YMW] + rSM![上枠下枠寸法用]
            rso![竪枠切断寸法] = rso![YMH] + rSM![竪枠寸法用]

            
            If rst![商品ﾀｲﾌﾟCD] = "S3" Then
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
            End If
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![格子竪切断寸法] = rso![YMH] + rSM![格子竪寸法用]
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
                STTD = rso![格子竪切断寸法]
            End If
            If rst![商品ﾀｲﾌﾟCD] = "XX" Then
                rso![格子竪切断寸法] = rso![YMH] - 50.5
                STYD = rso![格子竪切断寸法]
            End If
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn1 = rso![YPW] * rso![YPW]
                dobn2 = rso![YPH] * rso![YPH]
                dobn3 = dobn1 + dobn2
                dobn4 = Sqr(dobn3)
                dobn = dobn4 / 2
                rso![YP1] = dobn
                P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![上枠下枠格子取付ﾋﾟｯﾁ]
            If rst![商品ﾀｲﾌﾟCD] = "S3" Or rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![竪枠格子取付ﾋﾟｯﾁ]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![竪枠格子取付ﾋﾟｯﾁ]
            End If
            
            If rst![商品ﾀｲﾌﾟCD] = "S3" Then
                rso![FS区分] = True
            End If
            
            STNO = rso![注文番号]
            STEN = rso![注文枝番]
            STDT = rso![DT区分]
            STAD = rSM![格子AD値]
            
            If rst![商品ﾀｲﾌﾟCD] = "S2" Then
                rso![FS区分] = True
            End If
            
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If
 
 
            If rst![商品ﾀｲﾌﾟCD] = "XX" Then
                If rso![上下枠切断寸法] = 1193 Then
                    rso![YN] = 9
                    rso![YPW] = 117
                    N1W = rso![YN]
                Else
                    rso![YN] = 14
                    rso![YPW] = 117
                    N1W = rso![YN]
                End If
            End If
            
            rso![リベット] = get_ri(rso![YN], rso![YM], rso![商品ﾀｲﾌﾟCD]) * rst![ｾｯﾄ数]
            
            rso.Update
           
            '格子部データの出力
            Select Case rst![商品ﾀｲﾌﾟCD]
                Case "S1"
                    Call S1_set     'クロス
                Case "S2"
                    Call S1_set     'クロス　飾窓
                Case "S3"
                    Call S3_set     '横
                Case "S4"
                    Call S4_set     '枡
                Case "XX"
                    Call XX_set     '稲葉製作所向け
            End Select
            
'クロス飾り窓スタート
            If rst![商品ﾀｲﾌﾟCD] = "S2" Then
            
            rsn.MoveFirst
            wno = rsn![採番no]
            wno = wno + 1
            rsn.Edit
            rsn![採番no] = wno
            rsn.Update
        
            rso.AddNew
            rso![注文番号] = rst![注文番号]
            rso![注文枝番] = rst![注文枝番]
            rso![COL] = rsd![色]
            rso![採番no] = wno
            rso![DT区分] = "0"
            rso![注文区分] = rst![注文区分]
            rso![受注区分] = rst![受注区分]
            rso![社内管理番号1] = rst![社内管理番号1]
            rso![社内管理番号2] = rst![社内管理番号2]
            rso![納期日] = rst![納期日]
            rso![投入予定ｾｯﾄ数] = rst![投入予定ｾｯﾄ数]
           
            kigou(0) = rsd![寸法記号1]
            kigou(1) = rsd![寸法記号2]
            kigou(2) = rsd![寸法記号3]
            kigou(3) = rsd![寸法記号4]
            kigou(4) = rsd![寸法記号5]
            kigou(5) = rsd![寸法記号6]
            kigou(6) = rsd![寸法記号7]
            kigou(7) = rsd![寸法記号8]
            sunpou(0) = rsd![寸法値1]
            sunpou(1) = rsd![寸法値2]
            sunpou(2) = rsd![寸法値3]
            sunpou(3) = rsd![寸法値4]
            sunpou(4) = rsd![寸法値5]
            sunpou(5) = rsd![寸法値6]
            sunpou(6) = rsd![寸法値7]
            sunpou(7) = rsd![寸法値8]
            
            If rst![特寸ﾏｰｸ] = "1" Then
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H1" Then
                        rso![H寸法値] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
                        
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W1" Then
                        rso![W寸法値] = sunpou(lct) / 10
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
            Else
                rso![H寸法値] = rst![H寸2]
                rso![W寸法値] = rst![W寸2]
            End If
            
            rso![商品ｺｰﾄﾞ] = rst![商品ｺｰﾄﾞ]
            rso![商品ﾀｲﾌﾟCD] = rst![商品ﾀｲﾌﾟCD]
            rso![変数ｸﾞﾙｰﾌﾟ] = rst![変数ｸﾞﾙｰﾌﾟ2]
            rso![上枠工作図番号] = rst![上枠工作図番号]
            ksz = rst![上枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![上枠型番] = rsk![型番]
            
            rso![下枠工作図番号] = rst![下枠工作図番号]
            ksz = rst![下枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![下枠型番] = rsk![型番]
            
            rso![竪枠工作図番号] = rst![竪枠工作図番号]
            ksz = rst![竪枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![竪枠型番] = rsk![型番]
            
            rso![格子部工作図番号] = rst![格子部工作図番号]
            ksz = rst![格子部工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![格子部型番] = rsk![型番]
            
            hds = rst![変数ｸﾞﾙｰﾌﾟ2]
            rSM.FindFirst "[変数ｸﾞﾙｰﾌﾟ]='" & hds & "'"
            
            If rSM![MW分割] = True Then
                rso![YMW] = (rso![W寸法値] / 2) + rSM![MW用変数]
            Else
                rso![YMW] = rso![W寸法値] + rSM![MW用変数]
            End If
            
            rso![YMH] = rso![H寸法値] + rSM![MH用変数]
            
            rso![上下枠切断寸法] = rso![YMW] + rSM![上枠下枠寸法用]
            rso![竪枠切断寸法] = rso![YMH] + rSM![竪枠寸法用]

            
            If rst![商品ﾀｲﾌﾟCD] = "S3" Then
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
            End If
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![格子竪切断寸法] = rso![YMH] + rSM![格子竪寸法用]
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
                STTD = rso![格子竪切断寸法]
            End If
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                    rso![YN] = kiriage(dobn, 0)
                Else
                    dobn = kirishute(dobn, 1)
                    rso![YN] = shisya(dobn, 0)
                End If
                N1W = rso![YN]
            End If
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
            dobn = kirishute(dobn, 1)
            rso![YM] = shisya(dobn, 0)
            M1W = rso![YM]
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rso![YN]
                rso![YPW] = shisya(dobn, 1)
            End If
            
            If rso![YM] = 0 Then
               MsgBox (rst![商品ｺｰﾄﾞ])
               MsgBox (rso![YMH])
               MsgBox (rSM![H方向マス目分子])
               MsgBox (rSM![H方向マス目分母])
            End If
            
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rso![YM]
            rso![YPH] = shisya(dobn, 1)
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3" Then
                 dobn1 = rso![YPW] * rso![YPW]
                 dobn2 = rso![YPH] * rso![YPH]
                 dobn3 = dobn1 + dobn2
                 dobn4 = Sqr(dobn3)
                 dobn = dobn4 / 2
                 rso![YP1] = dobn
                 P1W = rso![YP1]
            End If
            
            rso![YA1] = rso![YPW] / 2 + rSM![上枠下枠格子取付ﾋﾟｯﾁ]
            If rst![商品ﾀｲﾌﾟCD] = "S3" Or rst![商品ﾀｲﾌﾟCD] = "S4" Then
                rso![YA2] = rso![YPH] + rSM![竪枠格子取付ﾋﾟｯﾁ]
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![竪枠格子取付ﾋﾟｯﾁ]
            End If
            
            STNO = rso![注文番号]
            STEN = rso![注文枝番]
            STDT = rso![DT区分]
            STAD = rSM![格子AD値]
            
            If rst![商品ﾀｲﾌﾟCD] = "S2" Then
                rso![FS区分] = False
            End If
            
            If rst![商品ﾀｲﾌﾟCD] = "S4" Then
                STPW = rso![YPW]
                STPH = rso![YPH]
            End If
            
            rso![リベット] = get_ri(rso![YN], rso![YM], rso![商品ﾀｲﾌﾟCD]) * rst![ｾｯﾄ数2]
 
            rso.Update
                       
            Call S1_set     'クロス　飾窓
            
            End If
        
'クロス飾り窓　ＥＮＤ

'同梱・取説・ダンボールスタート
    
        
'            rst.Edit
'            rst![ﾃﾞｰﾀ作成区分] = True
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

Sub Kc_Add() '新ピッチ面格子 2010/04/30

    'かんばんデータ作成
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
    
    If DCount("[注文番号]", "j_YOTEI_A") = 0 Then
        Exit Sub
    End If
    
    Set DB = CurrentDb
    Set rst = DB.OpenRecordset("j_YOTEI_A")
    Set rsd = DB.OpenRecordset("q_SAN累積", dbOpenDynaset)
    Set rSM = DB.OpenRecordset("m_変数", dbOpenDynaset)
    Set rsn = DB.OpenRecordset("m_ﾜｰｸno")
    Set rso = DB.OpenRecordset("YOTEIﾜｰｸ")
    Set rsk = DB.OpenRecordset("m_工作図番号", dbOpenDynaset)
        
    Set DBB = CurrentDb
    Set RECS = DBB.OpenRecordset("格子部ﾜｰｸ")


    rst.MoveFirst
    Do Until rst.EOF

        sds = rst![注文番号]
        rsd.FindFirst "[注文番号]='" & sds & "'"
    
        ' NoMatch プロパティの値に基づいて戻り値を設定します。
        If rsd.NoMatch = False Then
        
            rsn.MoveFirst
            wno = rsn![採番no]
            wno = wno + 1
            rsn.Edit
            rsn![採番no] = wno
            rsn.Update
        
            rso.AddNew
            rso![注文番号] = rst![注文番号]
            rso![注文枝番] = rst![注文枝番]
            rso![COL] = rst![色]
            
            rso![採番no] = wno
            rso![DT区分] = "0"
            rso![注文区分] = rst![注文区分]
            rso![受注区分] = rst![受注区分]
            
            rso![社内管理番号1] = rst![社内管理番号1]
            kot = "0000" + rst![社内管理番号2]
            rso![社内管理番号2] = Right(kot, 4)
            
            rso![納期日] = rst![納期日]
            rso![投入予定ｾｯﾄ数] = rst![投入予定ｾｯﾄ数]
            rso![数量] = rst![数量]
           
            kigou(0) = rsd![寸法記号1]
            kigou(1) = rsd![寸法記号2]
            kigou(2) = rsd![寸法記号3]
            kigou(3) = rsd![寸法記号4]
            kigou(4) = rsd![寸法記号5]
            kigou(5) = rsd![寸法記号6]
            kigou(6) = rsd![寸法記号7]
            kigou(7) = rsd![寸法記号8]
            sunpou(0) = rsd![寸法値1]
            sunpou(1) = rsd![寸法値2]
            sunpou(2) = rsd![寸法値3]
            sunpou(3) = rsd![寸法値4]
            sunpou(4) = rsd![寸法値5]
            sunpou(5) = rsd![寸法値6]
            sunpou(6) = rsd![寸法値7]
            sunpou(7) = rsd![寸法値8]
            
            If rst![特寸ﾏｰｸ] = "1" Then
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "H" Or kigou(lct) = "H2" Then
                        rso![H寸法値] = sunpou(lct) / 10
                        'Exit Do
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
                
                lct = 0        ' 変数を初期化します。
                chk = True
                Do While lct < 8        '
                    If lct = 8 Then     ' 条件が True であれば
                        chk = False     ' フラグの値を False に設定します。
                        Exit Do         ' ループから抜けます。
                    End If
                    If kigou(lct) = "W" Or kigou(lct) = "W2" Then
                        rso![W寸法値] = sunpou(lct) / 10
                        'Exit Do
                    End If
                    lct = lct + 1       ' カウンタを増やします。
                Loop
            Else
                rso![H寸法値] = rst![H寸]
                rso![W寸法値] = rst![W寸]
            End If
            
            rso![商品ｺｰﾄﾞ] = rst![商品ｺｰﾄﾞ]
            rso![商品ﾀｲﾌﾟCD] = rst![商品ﾀｲﾌﾟCD]
            rso![変数ｸﾞﾙｰﾌﾟ] = rst![変数ｸﾞﾙｰﾌﾟ]
            rso![上枠工作図番号] = rst![上枠工作図番号]
            ksz = rst![上枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![上枠型番] = rsk![型番]
            
            rso![下枠工作図番号] = rst![下枠工作図番号]
            ksz = rst![下枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![下枠型番] = rsk![型番]
            
            rso![竪枠工作図番号] = rst![竪枠工作図番号]
            ksz = rst![竪枠工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![竪枠型番] = rsk![型番]
            
            rso![格子部工作図番号] = rst![格子部工作図番号]
            ksz = rst![格子部工作図番号]
            rsk.FindFirst "[工作図番号]='" & ksz & "'"
            rso![格子部型番] = rsk![型番]
            
            hds = rst![変数ｸﾞﾙｰﾌﾟ]
            rSM.FindFirst "[変数ｸﾞﾙｰﾌﾟ]='" & hds & "'"
            
            
            If rSM![MW分割] = True Then
                rso![YMW] = (rso![W寸法値] / 2) + rSM![MW用変数]
            Else
                rso![YMW] = rso![W寸法値] + rSM![MW用変数]
            End If
            
            rso![YMH] = rso![H寸法値] + rSM![MH用変数]
            
            rso![上下枠切断寸法] = rso![YMW] + rSM![上枠下枠寸法用]
            rso![竪枠切断寸法] = rso![YMH] + rSM![竪枠寸法用]

            '=============================格子
            Select Case rst![商品ﾀｲﾌﾟCD]
             Case "S3A"
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
             Case "S4A", "S5A"
                rso![格子竪切断寸法] = rso![YMH] + rSM![格子竪寸法用]
                rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                STYD = rso![格子横切断寸法]
                STTD = rso![格子竪切断寸法]
            End Select
            
            '---------M,N
            Select Case rst![商品ﾀｲﾌﾟCD]
             Case "S1A", "S2A"
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                rso![YN] = kiriage(dobn, 0)
                dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
                rso![YM] = kiriage(dobn, 0)
             Case "S3A" '横
                dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
                rso![YM] = shisya2(dobn, 0)
             Case "S4A", "S5A" '枡
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                rso![YN] = kirishute(dobn, 0) + 2
                dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
                rso![YM] = kirishute(dobn, 0) + 2
            End Select
            N1W = rso![YN]
            M1W = rso![YM]
            
            
            '---------PW,PH 格子ピッチ
            If rst![商品ﾀｲﾌﾟCD] <> "S3A" Then
                dobn = (rso![YMW] + rSM![W方向マス目分子]) / rso![YN]
                rso![YPW] = shisya2(dobn, 1)
            End If
            dobn = (rso![YMH] + rSM![H方向マス目分子]) / rso![YM]
            
            If rst![商品ﾀｲﾌﾟCD] = "S4A" Or rst![商品ﾀｲﾌﾟCD] = "S5A" Then
                rso![YPW] = 100
                rso![YPH] = 100
            Else
                rso![YPH] = shisya2(dobn, 1)
            End If
            
            If rst![商品ﾀｲﾌﾟCD] <> "S3A" Then
                dobn1 = rso![YPW] * rso![YPW]
                dobn2 = rso![YPH] * rso![YPH]
                dobn3 = dobn1 + dobn2
                dobn4 = Sqr(dobn3)
                dobn = dobn4 / 2
                rso![YP1] = shisya2(dobn, 1)
                P1W = rso![YP1]
            End If
            
            '---------YA1,YA2 枠AB値
            rso![YA1] = rso![YPW] / 2 + rSM![上枠下枠格子取付ﾋﾟｯﾁ]
            If rst![商品ﾀｲﾌﾟCD] = "S3A" Then
                rso![YA2] = rso![YPH] + rSM![竪枠格子取付ﾋﾟｯﾁ]
            ElseIf rst![商品ﾀｲﾌﾟCD] = "S4A" And rst![変数ｸﾞﾙｰﾌﾟ] = "239" Then
                rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 50 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![商品ﾀｲﾌﾟCD] = "S4A" And rst![変数ｸﾞﾙｰﾌﾟ] = "283" Then 'スマージュ用
                rso![YA1] = (rso![YMW] - 54 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 64 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![商品ﾀｲﾌﾟCD] = "S4A" And rst![変数ｸﾞﾙｰﾌﾟ] >= "266" Then
                rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![商品ﾀｲﾌﾟCD] = "S5A" And rst![変数ｸﾞﾙｰﾌﾟ] >= "266" Then
                rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
            ElseIf rst![商品ﾀｲﾌﾟCD] = "S4A" And rst![変数ｸﾞﾙｰﾌﾟ] >= "264" Then
                rso![YA1] = (rso![YMW] - 54 - (100 * (rso![YN] - 2))) / 2
                rso![YA2] = (rso![YMH] - 64 - (100 * (rso![YM] - 2))) / 2
            Else
                rso![YA2] = rso![YPH] / 2 + rSM![竪枠格子取付ﾋﾟｯﾁ]
            End If
            
            
            If rst![商品ﾀｲﾌﾟCD] = "S3A" Then
                rso![FS区分] = True
            End If
            
            STNO = rso![注文番号]
            STEN = rso![注文枝番]
            STDT = rso![DT区分]
            STAD = rSM![格子AD値]
            
            If rst![商品ﾀｲﾌﾟCD] = "S2A" Or rst![商品ﾀｲﾌﾟCD] = "S5A" Then
                rso![FS区分] = True
            End If
            
            If rst![商品ﾀｲﾌﾟCD] = "S4A" Or rst![商品ﾀｲﾌﾟCD] = "S5A" Then
                STPW = rso![YA1]
                STPH = rso![YA2]
            End If
 
 
            rso![リベット] = get_ri(rso![YN], rso![YM], rso![商品ﾀｲﾌﾟCD]) * rst![ｾｯﾄ数]
            
            rso.Update
           
            '格子部データの出力
            Select Case rst![商品ﾀｲﾌﾟCD]
                Case "S1A"
                    Call S1_set     'クロス
                Case "S2A"
                    Call S1_set     'クロス　飾窓
                Case "S3A"
                    Call S3_set     '横
                Case "S4A", "S5A"
                    Call S4A_set    '枡
            End Select
            
'２枚伝票スタート
            If rst![商品ﾀｲﾌﾟCD] = "S2A" Or rst![商品ﾀｲﾌﾟCD] = "S5A" Then
                rsn.MoveFirst
                wno = rsn![採番no]
                wno = wno + 1
                rsn.Edit
                rsn![採番no] = wno
                rsn.Update
        
                rso.AddNew
                rso![注文番号] = rst![注文番号]
                rso![注文枝番] = rst![注文枝番]
                rso![COL] = rst![色]
                rso![採番no] = wno
                rso![DT区分] = "0"
                rso![注文区分] = rst![注文区分]
                rso![受注区分] = rst![受注区分]
                rso![社内管理番号1] = rst![社内管理番号1]
                kot = "0000" + rst![社内管理番号2]
                rso![社内管理番号2] = Right(kot, 4)
                rso![納期日] = rst![納期日]
                rso![投入予定ｾｯﾄ数] = rst![投入予定ｾｯﾄ数]
                rso![数量] = rst![数量]
           
                kigou(0) = rsd![寸法記号1]
                kigou(1) = rsd![寸法記号2]
                kigou(2) = rsd![寸法記号3]
                kigou(3) = rsd![寸法記号4]
                kigou(4) = rsd![寸法記号5]
                kigou(5) = rsd![寸法記号6]
                kigou(6) = rsd![寸法記号7]
                kigou(7) = rsd![寸法記号8]
                sunpou(0) = rsd![寸法値1]
                sunpou(1) = rsd![寸法値2]
                sunpou(2) = rsd![寸法値3]
                sunpou(3) = rsd![寸法値4]
                sunpou(4) = rsd![寸法値5]
                sunpou(5) = rsd![寸法値6]
                sunpou(6) = rsd![寸法値7]
                sunpou(7) = rsd![寸法値8]
            
                If rst![特寸ﾏｰｸ] = "1" Then
                    lct = 0        ' 変数を初期化します。
                    chk = True
                    Do While lct < 8        '
                        If lct = 8 Then     ' 条件が True であれば
                            chk = False     ' フラグの値を False に設定します。
                            Exit Do         ' ループから抜けます。
                        End If
                        If kigou(lct) = "H" Or kigou(lct) = "H1" Then
                            rso![H寸法値] = sunpou(lct) / 10
                            'Exit Do
                        End If
                        lct = lct + 1       ' カウンタを増やします。
                    Loop
                
                    lct = 0        ' 変数を初期化します。
                    chk = True
                    Do While lct < 8        '
                        If lct = 8 Then     ' 条件が True であれば
                            chk = False     ' フラグの値を False に設定します。
                            Exit Do         ' ループから抜けます。
                        End If
                        If kigou(lct) = "W" Or kigou(lct) = "W1" Then
                            rso![W寸法値] = sunpou(lct) / 10
                            'Exit Do
                        End If
                        lct = lct + 1       ' カウンタを増やします。
                    Loop
                Else
                    rso![H寸法値] = rst![H寸2]
                    rso![W寸法値] = rst![W寸2]
                End If
            
                rso![商品ｺｰﾄﾞ] = rst![商品ｺｰﾄﾞ]
                rso![商品ﾀｲﾌﾟCD] = rst![商品ﾀｲﾌﾟCD]
                rso![変数ｸﾞﾙｰﾌﾟ] = rst![変数ｸﾞﾙｰﾌﾟ2]
                rso![上枠工作図番号] = rst![上枠工作図番号]
                ksz = rst![上枠工作図番号]
                rsk.FindFirst "[工作図番号]='" & ksz & "'"
                rso![上枠型番] = rsk![型番]
            
                rso![下枠工作図番号] = rst![下枠工作図番号]
                ksz = rst![下枠工作図番号]
                rsk.FindFirst "[工作図番号]='" & ksz & "'"
                rso![下枠型番] = rsk![型番]
            
                rso![竪枠工作図番号] = rst![竪枠工作図番号]
                ksz = rst![竪枠工作図番号]
                rsk.FindFirst "[工作図番号]='" & ksz & "'"
                rso![竪枠型番] = rsk![型番]
            
                rso![格子部工作図番号] = rst![格子部工作図番号]
                ksz = rst![格子部工作図番号]
                rsk.FindFirst "[工作図番号]='" & ksz & "'"
                rso![格子部型番] = rsk![型番]
                
                hds = rst![変数ｸﾞﾙｰﾌﾟ2]
                rSM.FindFirst "[変数ｸﾞﾙｰﾌﾟ]='" & hds & "'"
            
            
                If rSM![MW分割] = True Then
                    rso![YMW] = (rso![W寸法値] / 2) + rSM![MW用変数]
                Else
                    rso![YMW] = rso![W寸法値] + rSM![MW用変数]
                End If
                
                rso![YMH] = rso![H寸法値] + rSM![MH用変数]
                
                rso![上下枠切断寸法] = rso![YMW] + rSM![上枠下枠寸法用]
                rso![竪枠切断寸法] = rso![YMH] + rSM![竪枠寸法用]

                '=============================格子
                Select Case rst![商品ﾀｲﾌﾟCD]
                 Case "S3A"
                    rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                    STYD = rso![格子横切断寸法]
                 Case "S4A", "S5A"
                    rso![格子竪切断寸法] = rso![YMH] + rSM![格子竪寸法用]
                    rso![格子横切断寸法] = rso![YMW] + rSM![格子横寸法用]
                    STYD = rso![格子横切断寸法]
                    STTD = rso![格子竪切断寸法]
                End Select
            
                '---------M,N
                Select Case rst![商品ﾀｲﾌﾟCD]
                 Case "S1A", "S2A"
                    dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                    rso![YN] = kiriage(dobn, 0)
                    dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
                    rso![YM] = kiriage(dobn, 0)
                 Case "S3A" '横
                    dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
                    rso![YM] = shisya2(dobn, 0)
                 Case "S4A", "S5A" '枡
                    dobn = (rso![YMW] + rSM![W方向マス目分子]) / rSM![W方向マス目分母]
                    rso![YN] = kirishute(dobn, 0) + 2
                    dobn = (rso![YMH] + rSM![H方向マス目分子]) / rSM![H方向マス目分母]
                    rso![YM] = kirishute(dobn, 0) + 2
                End Select
                N1W = rso![YN]
                M1W = rso![YM]
            
            
                '---------PW,PH 格子ピッチ
                If rst![商品ﾀｲﾌﾟCD] <> "S3A" Then
                    dobn = (rso![YMW] + rSM![W方向マス目分子]) / rso![YN]
                    rso![YPW] = shisya2(dobn, 1)
                End If
                dobn = (rso![YMH] + rSM![H方向マス目分子]) / rso![YM]
                
                If rst![商品ﾀｲﾌﾟCD] = "S4A" Or rst![商品ﾀｲﾌﾟCD] = "S5A" Then
                    rso![YPW] = 100
                    rso![YPH] = 100
                Else
                    rso![YPH] = shisya2(dobn, 1)
                End If
            
                If rst![商品ﾀｲﾌﾟCD] <> "S3A" Then
                    dobn1 = rso![YPW] * rso![YPW]
                    dobn2 = rso![YPH] * rso![YPH]
                    dobn3 = dobn1 + dobn2
                    dobn4 = Sqr(dobn3)
                    dobn = dobn4 / 2
                    rso![YP1] = shisya2(dobn, 1)
                    P1W = rso![YP1]
                End If
            
                '---------YA1,YA2 枠AB値
                rso![YA1] = rso![YPW] / 2 + rSM![上枠下枠格子取付ﾋﾟｯﾁ]
                If rst![商品ﾀｲﾌﾟCD] = "S3A" Then
                    rso![YA2] = rso![YPH] + rSM![竪枠格子取付ﾋﾟｯﾁ]
                ElseIf rst![商品ﾀｲﾌﾟCD] = "S4A" And rst![変数ｸﾞﾙｰﾌﾟ] = "239" Then
                    rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 50 - (100 * (rso![YM] - 2))) / 2
                ElseIf rst![商品ﾀｲﾌﾟCD] = "S4A" And rst![変数ｸﾞﾙｰﾌﾟ] >= "266" Then
                    rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
                ElseIf rst![商品ﾀｲﾌﾟCD] = "S5A" And rst![変数ｸﾞﾙｰﾌﾟ] >= "266" Then
                    rso![YA1] = (rso![YMW] - 22 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 36 - (100 * (rso![YM] - 2))) / 2
                ElseIf rst![商品ﾀｲﾌﾟCD] = "S4A" And rst![変数ｸﾞﾙｰﾌﾟ] >= "264" Then
                    rso![YA1] = (rso![YMW] - 54 - (100 * (rso![YN] - 2))) / 2
                    rso![YA2] = (rso![YMH] - 64 - (100 * (rso![YM] - 2))) / 2
                Else
                    rso![YA2] = rso![YPH] / 2 + rSM![竪枠格子取付ﾋﾟｯﾁ]
                End If
            
            
                STNO = rso![注文番号]
                STEN = rso![注文枝番]
                STDT = rso![DT区分]
                STAD = rSM![格子AD値]
            
                If rst![商品ﾀｲﾌﾟCD] = "S2A" Or rst![商品ﾀｲﾌﾟCD] = "S5A" Then
                    rso![FS区分] = False
                End If
            
                If rst![商品ﾀｲﾌﾟCD] = "S4A" Or rst![商品ﾀｲﾌﾟCD] = "S5A" Then
                    STPW = rso![YA1]
                    STPH = rso![YA2]
                End If
 
 
                rso![リベット] = get_ri(rso![YN], rso![YM], rso![商品ﾀｲﾌﾟCD]) * rst![ｾｯﾄ数2]
                
                rso.Update
           
                '格子部データの出力
                Select Case rst![商品ﾀｲﾌﾟCD]
                    Case "S2A"
                        Call S1_set     'クロス　飾窓
                    Case "S5A"
                        Call S4A_set    '枡
                End Select
            End If
            
        
'２枚伝票　ＥＮＤ

'同梱・取説・ダンボールスタート
    
        
'            rst.Edit
'            rst![ﾃﾞｰﾀ作成区分] = True
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
            Call S1_NM3             ' M = 偶数
        Else                        ' M = 奇数
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
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = 1
    RECS![切断寸法] = STYD
    RECS![切断本数] = N1W
    RECS![寸法1] = STAD
    RECS.Update
    
End Sub
Public Sub S3_set()

    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = 1
    RECS![切断寸法] = STYD
    RECS![切断本数] = M1W - 1
    RECS![寸法1] = STAD
    RECS.Update
    
End Sub
Public Sub S4_set()

    Dim intA As Integer
    Dim M1WW As Integer
    Dim GHON As Integer
    Dim KHON As Integer
    Dim RECNO As Integer
    
    '竪格子
    RECNO = 1
    
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STTD
    RECS![寸法1] = 6.5
    RECS![格子区分] = "1"
    
    If N1W = 2 Then
        RECS![切断本数] = 1
        
        dobn = (M1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![穴数] = M1WW
        
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![寸法2] = 6.5 + STPH * lct * 2
            Case 2
                RECS![寸法3] = 6.5 + STPH * lct * 2
            Case 3
                RECS![寸法4] = 6.5 + STPH * lct * 2
            Case 4
                RECS![寸法5] = 6.5 + STPH * lct * 2
            Case 5
                RECS![寸法6] = 6.5 + STPH * lct * 2
            Case 6
                RECS![寸法7] = 6.5 + STPH * lct * 2
            Case 7
                RECS![寸法8] = 6.5 + STPH * lct * 2
            Case 8
                RECS![寸法9] = 6.5 + STPH * lct * 2
            Case 9
                RECS![寸法10] = 6.5 + STPH * lct * 2
            Case 10
                RECS![寸法11] = 6.5 + STPH * lct * 2
            Case 11
                RECS![寸法12] = 6.5 + STPH * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (M1W Mod 2) <> 0 Then
            'RECS![切断本数] = N1W      中間部修正
            RECS![切断本数] = N1W - 1
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = 6.5 + STPH * lct * 2
                Case 2
                    RECS![寸法3] = 6.5 + STPH * lct * 2
                Case 3
                    RECS![寸法4] = 6.5 + STPH * lct * 2
                Case 4
                    RECS![寸法5] = 6.5 + STPH * lct * 2
                Case 5
                    RECS![寸法6] = 6.5 + STPH * lct * 2
                Case 6
                    RECS![寸法7] = 6.5 + STPH * lct * 2
                Case 7
                    RECS![寸法8] = 6.5 + STPH * lct * 2
                Case 8
                    RECS![寸法9] = 6.5 + STPH * lct * 2
                Case 9
                    RECS![寸法10] = 6.5 + STPH * lct * 2
                Case 10
                    RECS![寸法11] = 6.5 + STPH * lct * 2
                Case 11
                    RECS![寸法12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (N1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![切断本数] = KHON
            
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = 6.5 + STPH * lct * 2
                Case 2
                    RECS![寸法3] = 6.5 + STPH * lct * 2
                Case 3
                    RECS![寸法4] = 6.5 + STPH * lct * 2
                Case 4
                    RECS![寸法5] = 6.5 + STPH * lct * 2
                Case 5
                    RECS![寸法6] = 6.5 + STPH * lct * 2
                Case 6
                    RECS![寸法7] = 6.5 + STPH * lct * 2
                Case 7
                    RECS![寸法8] = 6.5 + STPH * lct * 2
                Case 8
                    RECS![寸法9] = 6.5 + STPH * lct * 2
                Case 9
                    RECS![寸法10] = 6.5 + STPH * lct * 2
                Case 10
                    RECS![寸法11] = 6.5 + STPH * lct * 2
                Case 11
                    RECS![寸法12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = RECNO
            RECS![切断寸法] = STTD
            RECS![寸法1] = 6.5
            RECS![格子区分] = "1"
            
            RECS![切断本数] = N1W - 1 - KHON
            
            dobn = M1W / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                Case 1
                    RECS![寸法2] = 6.5 + STPH
                Case 2
                    RECS![寸法3] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 3
                    RECS![寸法4] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 4
                    RECS![寸法5] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 5
                    RECS![寸法6] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 6
                    RECS![寸法7] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 7
                    RECS![寸法8] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 8
                    RECS![寸法9] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 9
                    RECS![寸法10] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 10
                    RECS![寸法11] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                Case 11
                    RECS![寸法12] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
    
    '横格子　中間部
    RECNO = RECNO + 1
    
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STYD
    RECS![寸法1] = STAD
    RECS![格子区分] = "2"
    
    If M1W = 2 Then
        RECS![切断本数] = 1
        dobn = (N1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![穴数] = M1WW
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![寸法2] = STAD + STPW * lct * 2
            Case 2
                RECS![寸法3] = STAD + STPW * lct * 2
            Case 3
                RECS![寸法4] = STAD + STPW * lct * 2
            Case 4
                RECS![寸法5] = STAD + STPW * lct * 2
            Case 5
                RECS![寸法6] = STAD + STPW * lct * 2
            Case 6
                RECS![寸法7] = STAD + STPW * lct * 2
            Case 7
                RECS![寸法8] = STAD + STPW * lct * 2
            Case 8
                RECS![寸法9] = STAD + STPW * lct * 2
            Case 9
                RECS![寸法10] = STAD + STPW * lct * 2
            Case 10
                RECS![寸法11] = STAD + STPW * lct * 2
            Case 11
                RECS![寸法12] = STAD + STPW * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (N1W Mod 2) <> 0 Then
            RECS![切断本数] = M1W - 1
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = STAD + STPW * lct * 2
                Case 2
                    RECS![寸法3] = STAD + STPW * lct * 2
                Case 3
                    RECS![寸法4] = STAD + STPW * lct * 2
                Case 4
                    RECS![寸法5] = STAD + STPW * lct * 2
                Case 5
                    RECS![寸法6] = STAD + STPW * lct * 2
                Case 6
                    RECS![寸法7] = STAD + STPW * lct * 2
                Case 7
                    RECS![寸法8] = STAD + STPW * lct * 2
                Case 8
                    RECS![寸法9] = STAD + STPW * lct * 2
                Case 9
                    RECS![寸法10] = STAD + STPW * lct * 2
                Case 10
                    RECS![寸法11] = STAD + STPW * lct * 2
                Case 11
                    RECS![寸法12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (M1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![切断本数] = KHON
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = STAD + STPW * lct * 2
                Case 2
                    RECS![寸法3] = STAD + STPW * lct * 2
                Case 3
                    RECS![寸法4] = STAD + STPW * lct * 2
                Case 4
                    RECS![寸法5] = STAD + STPW * lct * 2
                Case 5
                    RECS![寸法6] = STAD + STPW * lct * 2
                Case 6
                    RECS![寸法7] = STAD + STPW * lct * 2
                Case 7
                    RECS![寸法8] = STAD + STPW * lct * 2
                Case 8
                    RECS![寸法9] = STAD + STPW * lct * 2
                Case 9
                    RECS![寸法10] = STAD + STPW * lct * 2
                Case 10
                    RECS![寸法11] = STAD + STPW * lct * 2
                Case 11
                    RECS![寸法12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = RECNO
            RECS![切断寸法] = STYD
            RECS![寸法1] = STAD
            RECS![格子区分] = "2"
            
            RECS![切断本数] = M1W - 1 - KHON
            
            dobn = N1W / 2
            M1WW = kirishute(dobn, 0)
            
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = STAD + STPW
                Case 2
                    RECS![寸法3] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 3
                    RECS![寸法4] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 4
                    RECS![寸法5] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 5
                    RECS![寸法6] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 6
                    RECS![寸法7] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 7
                    RECS![寸法8] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 8
                    RECS![寸法9] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 9
                    RECS![寸法10] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 10
                    RECS![寸法11] = STAD + STPW + (STPW * (lct - 1) * 2)
                Case 11
                    RECS![寸法12] = STAD + STPW + (STPW * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
         
    '横格子 上下端部
    RECNO = RECNO + 1
         
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STYD
    RECS![寸法1] = STAD
    RECS![格子区分] = "3"
    RECS![切断本数] = 2
         
    If (N1W Mod 2) = 0 Then
        M1WW = N1W
        RECS![穴数] = M1WW - 1
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![寸法2] = STAD + STPW * lct
            Case 2
                RECS![寸法3] = STAD + STPW * lct
            Case 3
                RECS![寸法4] = STAD + STPW * lct
            Case 4
                RECS![寸法5] = STAD + STPW * lct
            Case 5
                RECS![寸法6] = STAD + STPW * lct
            Case 6
                RECS![寸法7] = STAD + STPW * lct
            Case 7
                RECS![寸法8] = STAD + STPW * lct
            Case 8
                RECS![寸法9] = STAD + STPW * lct
            Case 9
                RECS![寸法10] = STAD + STPW * lct
            Case 10
                RECS![寸法11] = STAD + STPW * lct
            Case 11
                RECS![寸法12] = STAD + STPW * lct
            End Select
        Next lct
        RECS.Update
    Else
        M1WW = N1W
        RECS![穴数] = M1WW - 1
        intA = M1WW / 2
        For lct = 1 To intA
            Select Case lct
            Case 1
                If lct = intA Then
                    RECS![寸法2] = STYD / 2
                Else
                    RECS![寸法2] = STAD + STPW * lct
                End If
            Case 2
                If lct = intA Then
                    RECS![寸法3] = STYD / 2
                Else
                    RECS![寸法3] = STAD + STPW * lct
                End If
            Case 3
                If lct = intA Then
                    RECS![寸法4] = STYD / 2
                Else
                    RECS![寸法4] = STAD + STPW * lct
                End If
            Case 4
                If lct = intA Then
                    RECS![寸法5] = STYD / 2
                Else
                    RECS![寸法5] = STAD + STPW * lct
                End If
            Case 5
                If lct = intA Then
                    RECS![寸法6] = STYD / 2
                Else
                    RECS![寸法6] = STAD + STPW * lct
                End If
            Case 6
                If lct = intA Then
                    RECS![寸法7] = STYD / 2
                Else
                    RECS![寸法7] = STAD + STPW * lct
                End If
            Case 7
                If lct = intA Then
                    RECS![寸法8] = STYD / 2
                Else
                    RECS![寸法8] = STAD + STPW * lct
                End If
            Case 8
                If lct = intA Then
                    RECS![寸法9] = STYD / 2
                Else
                    RECS![寸法9] = STAD + STPW * lct
                End If
            Case 9
                If lct = intA Then
                    RECS![寸法10] = STYD / 2
                Else
                    RECS![寸法10] = STAD + STPW * lct
                End If
            Case 10
                If lct = intA Then
                    RECS![寸法11] = STYD / 2
                Else
                    RECS![寸法11] = STAD + STPW * lct
                End If
            Case 11
                If lct = intA Then
                    RECS![寸法12] = STYD / 2
                Else
                    RECS![寸法12] = STAD + STPW * lct
                End If
            End Select
        Next lct
        
        For lct = intA To M1WW - 1
            Select Case lct
            Case 1
                    RECS![寸法2] = STAD + STPW * lct
            Case 2
                    RECS![寸法3] = STAD + STPW * lct
            Case 3
                    RECS![寸法4] = STAD + STPW * lct
            Case 4
                    RECS![寸法5] = STAD + STPW * lct
            Case 5
                    RECS![寸法6] = STAD + STPW * lct
            Case 6
                    RECS![寸法7] = STAD + STPW * lct
            Case 7
                    RECS![寸法8] = STAD + STPW * lct
            Case 8
                    RECS![寸法9] = STAD + STPW * lct
            Case 9
                    RECS![寸法10] = STAD + STPW * lct
            Case 10
                    RECS![寸法11] = STAD + STPW * lct
            Case 11
                    RECS![寸法12] = STAD + STPW * lct
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
    
    '竪格子
    RECNO = 1
    
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STTD
    RECS![寸法1] = 6.5
    RECS![格子区分] = "1"
    
    If N1W = 2 Then 'N1W=YN OK!
        RECS![切断本数] = 1

        dobn = (M1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![穴数] = M1WW

        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![寸法2] = 6.5 + STPH + 100
            Case 2
                RECS![寸法3] = RECS![寸法2] + 200
            Case 3
                RECS![寸法4] = RECS![寸法3] + 200
            Case 4
                RECS![寸法5] = RECS![寸法4] + 200
            Case 5
                RECS![寸法6] = RECS![寸法5] + 200
            Case 6
                RECS![寸法7] = RECS![寸法6] + 200
            Case 7
                RECS![寸法8] = RECS![寸法7] + 200
            Case 8
                RECS![寸法9] = RECS![寸法8] + 200
            Case 9
                RECS![寸法10] = RECS![寸法9] + 200
            Case 10
                RECS![寸法11] = RECS![寸法10] + 200
            Case 11
                RECS![寸法12] = RECS![寸法11] + 200
            End Select
        Next lct
        RECS.Update
    Else
        If (M1W Mod 2) <> 0 Then '縦方向枡目数が奇数(横格子本数が偶数)なら１パターンでＯＫ N=6,n=5
            RECS![切断本数] = N1W - 1
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![寸法3] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![寸法4] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![寸法5] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![寸法6] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![寸法7] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![寸法8] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![寸法9] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![寸法10] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![寸法11] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![寸法12] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
        Else '縦方向枡目数が奇数(横格子本数が偶数)なら格子加工パターンは２パターン N=7,n=6
            dobn = (N1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![切断本数] = KHON
            
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![寸法3] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![寸法4] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![寸法5] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![寸法6] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![寸法7] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![寸法8] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![寸法9] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![寸法10] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![寸法11] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![寸法12] = 6.5 + STPH + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = RECNO
            RECS![切断寸法] = STTD
            RECS![寸法1] = 6.5
            RECS![格子区分] = "1"
            
            RECS![切断本数] = N1W - 1 - KHON
            
            dobn = M1W / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                Case 1
                    RECS![寸法2] = 6.5 + STPH
                Case 2
                    RECS![寸法3] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 3
                    RECS![寸法4] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 4
                    RECS![寸法5] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 5
                    RECS![寸法6] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 6
                    RECS![寸法7] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 7
                    RECS![寸法8] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 8
                    RECS![寸法9] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 9
                    RECS![寸法10] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 10
                    RECS![寸法11] = 6.5 + STPH + (lct - 1) * 2 * 100
                Case 11
                    RECS![寸法12] = 6.5 + STPH + (lct - 1) * 2 * 100
                End Select
            Next lct
            RECS.Update
        End If
    End If
    
    '横格子　中間部
    RECNO = RECNO + 1
    
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STYD
    RECS![寸法1] = STAD
    RECS![格子区分] = "2"
    
    If M1W = 2 Then
        RECS![切断本数] = 1
        dobn = (N1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![穴数] = M1WW
        For lct = 1 To M1WW
            Select Case lct
            Case 1
                RECS![寸法2] = STAD + STPW + 100
            Case 2
                RECS![寸法3] = STAD + STPW + (lct - 1) * 2 * 100
            Case 3
                RECS![寸法4] = STAD + STPW + (lct - 1) * 2 * 100
            Case 4
                RECS![寸法5] = STAD + STPW + (lct - 1) * 2 * 100
            Case 5
                RECS![寸法6] = STAD + STPW + (lct - 1) * 2 * 100
            Case 6
                RECS![寸法7] = STAD + STPW + (lct - 1) * 2 * 100
            Case 7
                RECS![寸法8] = STAD + STPW + (lct - 1) * 2 * 100
            Case 8
                RECS![寸法9] = STAD + STPW + (lct - 1) * 2 * 100
            Case 9
                RECS![寸法10] = STAD + STPW + (lct - 1) * 2 * 100
            Case 10
                RECS![寸法11] = STAD + STPW + (lct - 1) * 2 * 100
            Case 11
                RECS![寸法12] = STAD + STPW + (lct - 1) * 2 * 100
            End Select
        Next lct
        RECS.Update
    Else
        If (N1W Mod 2) <> 0 Then  '横方向枡目数が奇数(竪格子本数が偶数)なら格子加工パターンは１パターン M=5,m=4
            RECS![切断本数] = M1W - 1
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![寸法3] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![寸法4] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![寸法5] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![寸法6] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![寸法7] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![寸法8] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![寸法9] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![寸法10] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![寸法11] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![寸法12] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
        Else '''''縦方向枡目数が奇数(横格子本数が偶数)なら格子加工パターンは２パターン N=7,n=6
            dobn = (M1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![切断本数] = KHON
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 2
                    RECS![寸法3] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 3
                    RECS![寸法4] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 4
                    RECS![寸法5] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 5
                    RECS![寸法6] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 6
                    RECS![寸法7] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 7
                    RECS![寸法8] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 8
                    RECS![寸法9] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 9
                    RECS![寸法10] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 10
                    RECS![寸法11] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                Case 11
                    RECS![寸法12] = STAD + STPW + ((lct - 1) * 2 + 1) * 100
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = RECNO
            RECS![切断寸法] = STYD
            RECS![寸法1] = STAD
            RECS![格子区分] = "2"
            
            RECS![切断本数] = M1W - 1 - KHON
            
            dobn = N1W / 2
            M1WW = kirishute(dobn, 0)
            
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                    RECS![寸法2] = STAD + STPW
                Case 2
                    RECS![寸法3] = STAD + STPW + (lct - 1) * 2 * 100
                Case 3
                    RECS![寸法4] = STAD + STPW + (lct - 1) * 2 * 100
                Case 4
                    RECS![寸法5] = STAD + STPW + (lct - 1) * 2 * 100
                Case 5
                    RECS![寸法6] = STAD + STPW + (lct - 1) * 2 * 100
                Case 6
                    RECS![寸法7] = STAD + STPW + (lct - 1) * 2 * 100
                Case 7
                    RECS![寸法8] = STAD + STPW + (lct - 1) * 2 * 100
                Case 8
                    RECS![寸法9] = STAD + STPW + (lct - 1) * 2 * 100
                Case 9
                    RECS![寸法10] = STAD + STPW + (lct - 1) * 2 * 100
                Case 10
                    RECS![寸法11] = STAD + STPW + (lct - 1) * 2 * 100
                Case 11
                    RECS![寸法12] = STAD + STPW + (lct - 1) * 2 * 100
                End Select
            Next lct
            RECS.Update
        End If
    End If
         
    '横格子 上下端部
    RECNO = RECNO + 1
         
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STYD
    RECS![寸法1] = STAD
    RECS![格子区分] = "3"
    RECS![切断本数] = 2
         
    M1WW = N1W
    RECS![穴数] = M1WW - 1
    For lct = 1 To M1WW - 1
        Select Case lct
        Case 1
            RECS![寸法2] = STAD + STPW
        Case 2
            RECS![寸法3] = STAD + STPW + (lct - 1) * 100
        Case 3
            RECS![寸法4] = STAD + STPW + (lct - 1) * 100
        Case 4
            RECS![寸法5] = STAD + STPW + (lct - 1) * 100
        Case 5
            RECS![寸法6] = STAD + STPW + (lct - 1) * 100
        Case 6
            RECS![寸法7] = STAD + STPW + (lct - 1) * 100
        Case 7
            RECS![寸法8] = STAD + STPW + (lct - 1) * 100
        Case 8
            RECS![寸法9] = STAD + STPW + (lct - 1) * 100
        Case 9
            RECS![寸法10] = STAD + STPW + (lct - 1) * 100
        Case 10
            RECS![寸法11] = STAD + STPW + (lct - 1) * 100
        Case 11
            RECS![寸法12] = STAD + STPW + (lct - 1) * 100
        End Select
    Next lct
    RECS.Update
     
End Sub

Public Sub P5_set() '（パナ 井桁面格子）格子部ﾜｰｸテーブルに値代入

    Dim intA As Integer
    Dim M1WW As Integer
    Dim GHON As Integer
    Dim KHON As Integer
    Dim RECNO As Integer
    
    '竪格子
    RECNO = 1
    
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STTD
    RECS![寸法1] = 6.5
    RECS![格子区分] = "1"
    
    If N1W = 2 Then  'N1W ：竪格子本数 YN
        RECS![切断本数] = 1
        
        dobn = (M1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![穴数] = M1WW
        
        For lct = 1 To M1WW 'M1WW ：
            Select Case lct
            Case 1
                RECS![寸法2] = 6.5 + STPH * lct * 2 'STPH：格子間隔 YPH
            Case 2
                RECS![寸法3] = 6.5 + STPH * lct * 2
            Case 3
                RECS![寸法4] = 6.5 + STPH * lct * 2
            Case 4
                RECS![寸法5] = 6.5 + STPH * lct * 2
            Case 5
                RECS![寸法6] = 6.5 + STPH * lct * 2
            Case 6
                RECS![寸法7] = 6.5 + STPH * lct * 2
            Case 7
                RECS![寸法8] = 6.5 + STPH * lct * 2
            Case 8
                RECS![寸法9] = 6.5 + STPH * lct * 2
            Case 9
                RECS![寸法10] = 6.5 + STPH * lct * 2
            Case 10
                RECS![寸法11] = 6.5 + STPH * lct * 2
            Case 11
                RECS![寸法12] = 6.5 + STPH * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (M1W Mod 2) <> 0 Then
            'RECS![切断本数] = N1W      中間部修正
            RECS![切断本数] = N1W - 1
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![寸法2] = 6.5 + STPH * lct * 2
                    Case 2
                        RECS![寸法3] = 6.5 + STPH * lct * 2
                    Case 3
                        RECS![寸法4] = 6.5 + STPH * lct * 2
                    Case 4
                        RECS![寸法5] = 6.5 + STPH * lct * 2
                    Case 5
                        RECS![寸法6] = 6.5 + STPH * lct * 2
                    Case 6
                        RECS![寸法7] = 6.5 + STPH * lct * 2
                    Case 7
                        RECS![寸法8] = 6.5 + STPH * lct * 2
                    Case 8
                        RECS![寸法9] = 6.5 + STPH * lct * 2
                    Case 9
                        RECS![寸法10] = 6.5 + STPH * lct * 2
                    Case 10
                        RECS![寸法11] = 6.5 + STPH * lct * 2
                    Case 11
                        RECS![寸法12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (N1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![切断本数] = KHON
            
            dobn = (M1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![寸法2] = 6.5 + STPH * lct * 2
                    Case 2
                        RECS![寸法3] = 6.5 + STPH * lct * 2
                    Case 3
                        RECS![寸法4] = 6.5 + STPH * lct * 2
                    Case 4
                        RECS![寸法5] = 6.5 + STPH * lct * 2
                    Case 5
                        RECS![寸法6] = 6.5 + STPH * lct * 2
                    Case 6
                        RECS![寸法7] = 6.5 + STPH * lct * 2
                    Case 7
                        RECS![寸法8] = 6.5 + STPH * lct * 2
                    Case 8
                        RECS![寸法9] = 6.5 + STPH * lct * 2
                    Case 9
                        RECS![寸法10] = 6.5 + STPH * lct * 2
                    Case 10
                        RECS![寸法11] = 6.5 + STPH * lct * 2
                    Case 11
                        RECS![寸法12] = 6.5 + STPH * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = RECNO
            RECS![切断寸法] = STTD
            RECS![寸法1] = 6.5
            RECS![格子区分] = "1"
            
            RECS![切断本数] = N1W - 1 - KHON
            
            dobn = M1W / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![寸法2] = 6.5 + STPH
                    Case 2
                        RECS![寸法3] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 3
                        RECS![寸法4] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 4
                        RECS![寸法5] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 5
                        RECS![寸法6] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 6
                        RECS![寸法7] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 7
                        RECS![寸法8] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 8
                        RECS![寸法9] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 9
                        RECS![寸法10] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 10
                        RECS![寸法11] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                    Case 11
                        RECS![寸法12] = 6.5 + STPH + (STPH * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
    
    '横格子　中間部
    RECNO = RECNO + 1
    
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STYD
    RECS![寸法1] = STAD
    RECS![格子区分] = "2"
    
    If M1W = 2 Then
        RECS![切断本数] = 1
        dobn = (N1W - 1) / 2
        M1WW = kirishute(dobn, 0)
        RECS![穴数] = M1WW
        For lct = 1 To M1WW
            Select Case lct
                Case 1
                    RECS![寸法2] = STAD + STPW * lct * 2
                Case 2
                    RECS![寸法3] = STAD + STPW * lct * 2
                Case 3
                    RECS![寸法4] = STAD + STPW * lct * 2
                Case 4
                    RECS![寸法5] = STAD + STPW * lct * 2
                Case 5
                    RECS![寸法6] = STAD + STPW * lct * 2
                Case 6
                    RECS![寸法7] = STAD + STPW * lct * 2
                Case 7
                    RECS![寸法8] = STAD + STPW * lct * 2
                Case 8
                    RECS![寸法9] = STAD + STPW * lct * 2
                Case 9
                    RECS![寸法10] = STAD + STPW * lct * 2
                Case 10
                    RECS![寸法11] = STAD + STPW * lct * 2
                Case 11
                    RECS![寸法12] = STAD + STPW * lct * 2
            End Select
        Next lct
        RECS.Update
    Else
        If (N1W Mod 2) <> 0 Then
            RECS![切断本数] = M1W - 1
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![寸法2] = STAD + STPW * lct * 2
                    Case 2
                        RECS![寸法3] = STAD + STPW * lct * 2
                    Case 3
                        RECS![寸法4] = STAD + STPW * lct * 2
                    Case 4
                        RECS![寸法5] = STAD + STPW * lct * 2
                    Case 5
                        RECS![寸法6] = STAD + STPW * lct * 2
                    Case 6
                        RECS![寸法7] = STAD + STPW * lct * 2
                    Case 7
                        RECS![寸法8] = STAD + STPW * lct * 2
                    Case 8
                        RECS![寸法9] = STAD + STPW * lct * 2
                    Case 9
                        RECS![寸法10] = STAD + STPW * lct * 2
                    Case 10
                        RECS![寸法11] = STAD + STPW * lct * 2
                    Case 11
                        RECS![寸法12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
        Else
            dobn = (M1W - 1) / 2
            KHON = kiriage(dobn, 0)
            RECS![切断本数] = KHON
            dobn = (N1W - 1) / 2
            M1WW = kirishute(dobn, 0)
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![寸法2] = STAD + STPW * lct * 2
                    Case 2
                        RECS![寸法3] = STAD + STPW * lct * 2
                    Case 3
                        RECS![寸法4] = STAD + STPW * lct * 2
                    Case 4
                        RECS![寸法5] = STAD + STPW * lct * 2
                    Case 5
                        RECS![寸法6] = STAD + STPW * lct * 2
                    Case 6
                        RECS![寸法7] = STAD + STPW * lct * 2
                    Case 7
                        RECS![寸法8] = STAD + STPW * lct * 2
                    Case 8
                        RECS![寸法9] = STAD + STPW * lct * 2
                    Case 9
                        RECS![寸法10] = STAD + STPW * lct * 2
                    Case 10
                        RECS![寸法11] = STAD + STPW * lct * 2
                    Case 11
                        RECS![寸法12] = STAD + STPW * lct * 2
                End Select
            Next lct
            RECS.Update
                    
            RECNO = RECNO + 1
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = RECNO
            RECS![切断寸法] = STYD
            RECS![寸法1] = STAD
            RECS![格子区分] = "2"
            
            RECS![切断本数] = M1W - 1 - KHON
            
            dobn = N1W / 2
            M1WW = kirishute(dobn, 0)
            
            RECS![穴数] = M1WW
            For lct = 1 To M1WW
                Select Case lct
                    Case 1
                        RECS![寸法2] = STAD + STPW
                    Case 2
                        RECS![寸法3] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 3
                        RECS![寸法4] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 4
                        RECS![寸法5] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 5
                        RECS![寸法6] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 6
                        RECS![寸法7] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 7
                        RECS![寸法8] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 8
                        RECS![寸法9] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 9
                        RECS![寸法10] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 10
                        RECS![寸法11] = STAD + STPW + (STPW * (lct - 1) * 2)
                    Case 11
                        RECS![寸法12] = STAD + STPW + (STPW * (lct - 1) * 2)
                End Select
            Next lct
            RECS.Update
        End If
    End If
         
    '横格子 上下端部
    RECNO = RECNO + 1
         
    RECS.AddNew
    RECS![注文番号] = STNO
    RECS![採番no] = wno
    RECS![注文枝番] = STEN
    RECS![DT区分] = STDT
    RECS![格子番号] = RECNO
    RECS![切断寸法] = STYD - 19 'MW-31 STYD(=MW-12)-19
    RECS![格子区分] = "3"
    RECS![切断本数] = 2
    
    RECS![寸法1] = 0
    STAD = STPW - 5.5
    RECS![寸法2] = STAD     '格子AD値
         
    If (N1W Mod 2) = 0 Then  'N1W ：W方向ピッチ（マス目）数 YN
        M1WW = N1W
        RECS![穴数] = M1WW - 1
        For lct = 1 To M1WW - 2
            Select Case lct
                Case 1
                    RECS![寸法3] = STAD + STPW * lct
                Case 2
                    RECS![寸法4] = STAD + STPW * lct
                Case 3
                    RECS![寸法5] = STAD + STPW * lct
                Case 4
                    RECS![寸法6] = STAD + STPW * lct
                Case 5
                    RECS![寸法7] = STAD + STPW * lct
                Case 6
                    RECS![寸法8] = STAD + STPW * lct
                Case 7
                    RECS![寸法9] = STAD + STPW * lct
                Case 8
                    RECS![寸法10] = STAD + STPW * lct
                Case 9
                    RECS![寸法11] = STAD + STPW * lct
                Case 10
                    RECS![寸法12] = STAD + STPW * lct
            End Select
        Next lct
        RECS.Update
    Else
        M1WW = N1W
        RECS![穴数] = M1WW - 1
        intA = M1WW / 2
        For lct = 1 To intA
            Select Case lct
            Case 1
                If lct = intA Then
                    RECS![寸法3] = STYD / 2
                Else
                    RECS![寸法3] = STAD + STPW * lct
                End If
            Case 2
                If lct = intA Then
                    RECS![寸法4] = STYD / 2
                Else
                    RECS![寸法4] = STAD + STPW * lct
                End If
            Case 3
                If lct = intA Then
                    RECS![寸法5] = STYD / 2
                Else
                    RECS![寸法5] = STAD + STPW * lct
                End If
            Case 4
                If lct = intA Then
                    RECS![寸法6] = STYD / 2
                Else
                    RECS![寸法6] = STAD + STPW * lct
                End If
            Case 5
                If lct = intA Then
                    RECS![寸法7] = STYD / 2
                Else
                    RECS![寸法7] = STAD + STPW * lct
                End If
            Case 6
                If lct = intA Then
                    RECS![寸法8] = STYD / 2
                Else
                    RECS![寸法8] = STAD + STPW * lct
                End If
            Case 7
                If lct = intA Then
                    RECS![寸法9] = STYD / 2
                Else
                    RECS![寸法9] = STAD + STPW * lct
                End If
            Case 8
                If lct = intA Then
                    RECS![寸法10] = STYD / 2
                Else
                    RECS![寸法10] = STAD + STPW * lct
                End If
            Case 9
                If lct = intA Then
                    RECS![寸法11] = STYD / 2
                Else
                    RECS![寸法11] = STAD + STPW * lct
                End If
            Case 10
                If lct = intA Then
                    RECS![寸法12] = STYD / 2
                Else
                    RECS![寸法12] = STAD + STPW * lct
                End If
            End Select
        Next lct
        
        For lct = intA To M1WW - 2
            Select Case lct
            Case 1
                    RECS![寸法3] = STAD + STPW * lct
            Case 2
                    RECS![寸法4] = STAD + STPW * lct
            Case 3
                    RECS![寸法5] = STAD + STPW * lct
            Case 4
                    RECS![寸法6] = STAD + STPW * lct
            Case 5
                    RECS![寸法7] = STAD + STPW * lct
            Case 6
                    RECS![寸法8] = STAD + STPW * lct
            Case 7
                    RECS![寸法9] = STAD + STPW * lct
            Case 8
                    RECS![寸法10] = STAD + STPW * lct
            Case 9
                    RECS![寸法11] = STAD + STPW * lct
            Case 10
                    RECS![寸法12] = STAD + STPW * lct
            End Select
        Next lct
        RECS.Update
    End If
End Sub



Public Sub Sps_Set1()
    For lct = 1 To RECS![穴数]
        Select Case lct
            Case 1
                RECS![寸法2] = STAD + (P1W * 4) * lct
            Case 2
                RECS![寸法3] = STAD + (P1W * 4) * lct
            Case 3
                RECS![寸法4] = STAD + (P1W * 4) * lct
            Case 4
                RECS![寸法5] = STAD + (P1W * 4) * lct
            Case 5
                RECS![寸法6] = STAD + (P1W * 4) * lct
            Case 6
                RECS![寸法7] = STAD + (P1W * 4) * lct
            Case 7
                RECS![寸法8] = STAD + (P1W * 4) * lct
            Case 8
                RECS![寸法9] = STAD + (P1W * 4) * lct
            Case 9
                RECS![寸法10] = STAD + (P1W * 4) * lct
            Case 10
                RECS![寸法11] = STAD + (P1W * 4) * lct
            Case 11
                RECS![寸法12] = STAD + (P1W * 4) * lct
        End Select
    Next lct
End Sub
Public Sub Sps_Set2()
    For lct = 1 To RECS![穴数]
        Select Case lct
            Case 1
                RECS![寸法2] = STAD + (P1W * 3)
            Case 2
                RECS![寸法3] = STAD + (P1W * 4) + (P1W * 3)
            Case 3
                RECS![寸法4] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 4
                RECS![寸法5] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 5
                RECS![寸法6] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 6
                RECS![寸法7] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 7
                RECS![寸法8] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 8
                RECS![寸法9] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 9
                RECS![寸法10] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 10
                RECS![寸法11] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
            Case 11
                RECS![寸法12] = STAD + (P1W * 4) * (lct - 1) + (P1W * 3)
        End Select
    Next lct
End Sub
Public Sub Sps_Set3()
    For lct = 1 To RECS![穴数]
        Select Case lct
            Case 1
                RECS![寸法2] = STAD + (P1W * 1)
            Case 2
                RECS![寸法3] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 3
                RECS![寸法4] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 4
                RECS![寸法5] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 5
                RECS![寸法6] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 6
                RECS![寸法7] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 7
                RECS![寸法8] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 8
                RECS![寸法9] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 9
                RECS![寸法10] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 10
                RECS![寸法11] = STAD + (P1W * 4) * (lct - 1) + P1W
            Case 11
                RECS![寸法12] = STAD + (P1W * 4) * (lct - 1) + P1W
        End Select
    Next lct
End Sub
Public Sub S1_M3()

    If N1W = M1W Then
        XNO = M1W
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![注文枝番] = STEN
            RECS![採番no] = wno
            RECS![DT区分] = STDT
            RECS![格子番号] = fnx
            RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
            RECS![切断本数] = 2 * 2
            RECS![寸法1] = STAD
            RECS.Update
        Next fnx
    End If
    
    If N1W > M1W Then
        XNO = M1W + 1
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = fnx
            If fnx <= M1W Then
                RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![切断本数] = 2 * 2
            Else
                RECS![切断寸法] = STAD * 2 + P1W * (2 * M1W)
                RECS![切断本数] = (N1W - M1W) * 2
            End If
            RECS![寸法1] = STAD
            RECS.Update
        Next fnx
    End If
    
    If N1W < M1W Then
        XNO = N1W + 1
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = fnx
            If fnx <= N1W Then
                RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![切断本数] = 2 * 2
            Else
                RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                RECS![切断本数] = (M1W - N1W) * 2
            End If
            RECS![寸法1] = STAD
            RECS.Update
        Next fnx
    End If
    
End Sub
Public Sub S1_NM3()

    If N1W = M1W Then
        XNO = M1W
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = fnx
            RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
            RECS![切断本数] = 2 * 2
            If fnx >= 3 Then
                dobn = fnx / 2 - 1
                RECS![穴数] = kiriage(dobn, 0)
            End If
            RECS![寸法1] = STAD
            If RECS![穴数] >= 1 Then
                Call Sps_Set1
            End If
            RECS.Update
        Next fnx
    End If
    
    If N1W > M1W Then
        XNO = M1W + 1
        For fnx = 1 To XNO
            RECS.AddNew
            RECS![注文番号] = STNO
            RECS![採番no] = wno
            RECS![注文枝番] = STEN
            RECS![DT区分] = STDT
            RECS![格子番号] = fnx
            If fnx <= M1W Then
                RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
            Else
                RECS![切断寸法] = STAD * 2 + P1W * (2 * M1W)
            End If
            
            If fnx <= M1W Then
                RECS![切断本数] = 2 * 2
            Else
                RECS![切断本数] = (N1W - M1W) * 2
            End If
            
            If fnx >= 3 And fnx <= M1W Then
                dobn = fnx / 2 - 1
                RECS![穴数] = kiriage(dobn, 0)
            End If
            If fnx > M1W Then
                RECS![穴数] = M1W / 2 - 1
            End If
            
            RECS![寸法1] = STAD
            If RECS![穴数] >= 1 Then
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
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                If fnx <= 2 Then
                    RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                Else
                    RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                End If
                
                If fnx <= 2 Then
                    RECS![切断本数] = 2 * 2
                Else
                    RECS![切断本数] = (M1W - N1W) * 2
                End If
                
                If fnx = 3 Then
                    RECS![穴数] = 1
                End If
                RECS![寸法1] = STAD
                If RECS![穴数] >= 1 Then
                    RECS![寸法2] = STAD + (P1W * 3)
                End If
                
                RECS.Update
            Next fnx
        Else
            If ((N1W Mod 2) = 0) Or (M1W - N1W = 1) Then
                XNO = N1W + 1
                For fnx = 1 To XNO
                    RECS.AddNew
                    RECS![注文番号] = STNO
                    RECS![採番no] = wno
                    RECS![注文枝番] = STEN
                    RECS![DT区分] = STDT
                    RECS![格子番号] = fnx
                    
                    If fnx <= N1W Then
                        RECS![切断本数] = 2 * 2
                    Else
                        RECS![切断本数] = (M1W - N1W) * 2
                    End If
            
                    RECS![寸法1] = STAD
                    If fnx < 3 Then
                        RECS![穴数] = 0
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                    End If
                    If fnx >= 3 And fnx <= M1W Then
                        dobn = fnx / 2 - 1
                        RECS![穴数] = kiriage(dobn, 0)
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        If RECS![穴数] >= 1 Then
                            Call Sps_Set1
                        End If
                    End If
                    If fnx > N1W Then
                        RECS![穴数] = fnx / 2 - 1
                        RECS![穴数] = kiriage(dobn, 0)
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        If RECS![穴数] >= 1 Then
                            Call Sps_Set2
                        End If
                    End If
                    
                    RECS.Update
                Next fnx
            Else
                XNO = N1W + 2
                For fnx = 1 To XNO
                    RECS.AddNew
                    RECS![注文番号] = STNO
                    RECS![採番no] = wno
                    RECS![注文枝番] = STEN
                    RECS![DT区分] = STDT
                    RECS![格子番号] = fnx
                    
                    If fnx <= N1W Then
                        RECS![切断本数] = 2 * 2
                    End If
                    If fnx = N1W + 1 Then
                        dobn = (M1W - N1W) / 2
                        RECS![切断本数] = kiriage(dobn, 0) * 2
                    End If
                    If fnx = N1W + 2 Then
                        dobn = (M1W - N1W) / 2
                        RECS![切断本数] = kirishute(dobn, 0) * 2
                    End If
                    
                    If fnx <= N1W Then
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                    Else
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                    End If
                    
                    RECS![寸法1] = STAD
                    
                    If fnx >= 3 Then
                        dobn = fnx / 2 - 1
                        RECS![穴数] = kiriage(dobn, 0)
                    End If
                    
                    If RECS![穴数] >= 1 Then
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
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![切断本数] = 2 * 2
                If fnx >= 4 Then
                    RECS![穴数] = 1
                End If
                RECS![寸法1] = STAD
                If RECS![穴数] >= 1 Then
                    RECS![寸法2] = STAD + (P1W * 5)
                End If
                RECS.Update
            Next fnx
        Case 7
            XNO = M1W
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![切断本数] = 2 * 2
                If fnx >= 3 And fnx <= 5 Then
                    RECS![穴数] = 1
                End If
                If fnx >= 6 Then
                    RECS![穴数] = 2
                End If
                RECS![寸法1] = STAD
                If RECS![穴数] >= 1 Then
                    Select Case RECS![穴数]
                        Case 1
                            RECS![寸法2] = STAD + (P1W * 4)
                        Case 2
                            RECS![寸法2] = STAD + (P1W * 4)
                            RECS![寸法3] = STAD + (P1W * 4) + (P1W * 6)
                    End Select
                End If
                RECS.Update
            Next fnx
        Case 9
            XNO = M1W
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![切断本数] = 2 * 2
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 3 To 5
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8, 9
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 11
            XNO = M1W
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                RECS![切断本数] = 2 * 2
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 3, 4
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                   Case 8, 9
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                   Case 10, 11
                        RECS![穴数] = 4
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![寸法5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
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
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 4, 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 5)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * M1W)
                        RECS![切断本数] = (N1W - M1W) * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 7
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 6)
                    Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * M1W)
                        RECS![切断本数] = (N1W - M1W) * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
        Case 9
            XNO = 10
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8, 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                    Case 10
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * M1W)
                        RECS![切断本数] = (N1W - M1W) * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx

        Case 11
            XNO = 12
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8, 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 10, 11
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 4
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![寸法5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                   Case 12
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * M1W)
                        RECS![切断本数] = (N1W - M1W) * 2
                        RECS![穴数] = 4
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![寸法5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
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
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 2)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 4
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 5
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 5)
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
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
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 4 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 5
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + P1W
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3, 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                   Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + (P1W * 7)     '04/02/01
                End Select
                
                RECS.Update
            Next fnx
        Case 5
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
        Case 6
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 6)
                    Case 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 6)
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
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 4 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 2)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + P1W
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3, 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 5)
                    Case 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
               End Select
                
                RECS.Update
            Next fnx
        Case 5
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 5)
                    Case 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 6
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 5)
                   Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 5)
                        RECS![寸法4] = STAD + P1W + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 7
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![注文枝番] = STEN
                RECS![採番no] = wno
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 5)
                        RECS![寸法4] = STAD + (P1W * 3) + (P1W * 5) + (P1W * 5)
                End Select
                
                RECS.Update
            Next fnx
        Case 8
            XNO = 9
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 6, 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                    Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 5)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 5) + (P1W * 5)
                    Case 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 5)
                        RECS![寸法4] = STAD + (P1W * 3) + (P1W * 5) + (P1W * 5)
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
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 8 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                End Select
                
                RECS.Update
            Next fnx
            
        Case 3
            XNO = 6
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 4 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 4)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + P1W
                End Select
                
                RECS.Update
            Next fnx
        Case 4
            XNO = 7
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3, 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 4 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 4)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 3)
                   Case 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 5
            XNO = 8
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 4)
                    Case 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 4)
                    Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
        Case 6
            XNO = 9
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5, 6
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 4)
                    Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 4)
                        RECS![寸法4] = STAD + P1W + (P1W * 4) + (P1W * 6)
                    Case 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 6)
                End Select
                
                RECS.Update
            Next fnx
            
        Case 7
            XNO = 9
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                    Case 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 4)
                        RECS![寸法4] = STAD + P1W + (P1W * 4) + (P1W * 6)
                End Select
           
                RECS.Update
            Next fnx
        Case 8
            XNO = 10
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                    Case 10
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 4
                        RECS![寸法2] = STAD + P1W
                        RECS![寸法3] = STAD + P1W + (P1W * 4)
                        RECS![寸法4] = STAD + P1W + (P1W * 4) + (P1W * 6)
                        RECS![寸法5] = STAD + P1W + (P1W * 4) + (P1W * 6) + (P1W * 4)
               End Select
                
                RECS.Update
            Next fnx
            
        Case 9
            XNO = 10
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8 To 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 10
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 4
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                        RECS![寸法5] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                End Select
           
                RECS.Update
            Next fnx
        Case 10
            XNO = 11
            For fnx = 1 To XNO
                RECS.AddNew
                RECS![注文番号] = STNO
                RECS![採番no] = wno
                RECS![注文枝番] = STEN
                RECS![DT区分] = STDT
                RECS![格子番号] = fnx
                
                RECS![寸法1] = STAD
                Select Case fnx
                    Case 1 To 2
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                    Case 3 To 4
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 1
                        RECS![寸法2] = STAD + (P1W * 4)
                    Case 5 To 7
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 2
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                    Case 8, 9
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 3
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                    Case 10
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * fnx - 1)
                        RECS![切断本数] = 2 * 2
                        RECS![穴数] = 4
                        RECS![寸法2] = STAD + (P1W * 4)
                        RECS![寸法3] = STAD + (P1W * 4) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6)
                        RECS![寸法5] = STAD + (P1W * 4) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                    Case 11
                        RECS![切断寸法] = STAD * 2 + P1W * (2 * N1W)
                        RECS![切断本数] = 1 * 2
                        RECS![穴数] = 4
                        RECS![寸法2] = STAD + (P1W * 3)
                        RECS![寸法3] = STAD + (P1W * 3) + (P1W * 4)
                        RECS![寸法4] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6)
                        RECS![寸法5] = STAD + (P1W * 3) + (P1W * 4) + (P1W * 6) + (P1W * 4)
                End Select
                
                RECS.Update
            Next fnx
            
    End Select
End Sub


Function get_sho(Sho As String, hen) '受注履歴用　（2015/7/16
    Select Case Sho
     Case "XX"
      get_sho = "Aイナバ"
     Case "P1", "P1A"
      get_sho = "Aクロス"
     Case "P2"
      get_sho = "Aサン"
     Case "P3", "P3A"
      get_sho = "Aクリア"
     Case "P4", "P4A"
      get_sho = "Aクリア"
     Case "P5"
      get_sho = "A井桁"
     Case "HS3"
      get_sho = "A支柱"
     Case "HT4"
        If hen = "PD11" Then
          get_sho = "AHIT支柱"
        ElseIf hen = "PD21" Then
          get_sho = "CHIT支柱"
        End If
     Case "SM1"
      get_sho = "AS型"

     Case "HS1"
      get_sho = "CHSク"
     Case "HS2"
      get_sho = "CHS竪"
     Case "HS4"
      get_sho = "CHWP"
     Case "HS5"
      get_sho = "CHW柱"
     Case "HS6"
      get_sho = "CHWレ"
     Case "HS7"
      get_sho = "CHS横"
     Case "HT1"
      get_sho = "CHITク"
     Case "HT2"
      get_sho = "CHIT竪"
     Case "HT3"
      get_sho = "CHIT横"
     Case "KK1"
      get_sho = "C腰壁ク"
     Case "KK2"
      get_sho = "C腰壁竪"
     Case "KK3"
      get_sho = "C腰壁横"
     Case "KK4"
      get_sho = "A腰壁支"
     Case "NS1"
      get_sho = "C2辺ク"
     Case "NS2"
      get_sho = "C2辺支"
     Case "FT1"
      get_sho = "BEX定"

     Case Else
        If Sho Like "S*" Then
            get_sho = "A受注"
        ElseIf Sho Like "M*" Then
            get_sho = "Cメイク"
        ElseIf Sho Like "G*" Then
            get_sho = "C柱"
        End If
    End Select

End Function

Function get_sho2(Sho As String, hen As Variant, 注文区分 As String)   '受注状況改造用　（2018/03/02)
    If 注文区分 = "A" Then
        Select Case Sho
         Case "XX"
          get_sho2 = "Aイナバ"
         Case "P1", "P1A"
          get_sho2 = "Aクロス"
         Case "P2"
          get_sho2 = "Aサン"
         Case "P3", "P3A"
          get_sho2 = "Aクリア"
         Case "P4", "P4A"
          get_sho2 = "Aクリア"
         Case "P5"
          get_sho2 = "A井桁"
         Case "HS3"
          get_sho2 = "A支柱"
         Case "HT4"
            If hen = "PD11" Then
              get_sho2 = "AHIT支柱"
            ElseIf hen = "PD21" Then
              get_sho2 = "CHIT支柱"
            End If
         Case "SM1"
          get_sho2 = "AS型"
    
         Case "HS1"
          get_sho2 = "CHSク"
         Case "HS2"
          get_sho2 = "CHS竪"
         Case "HS4"
          get_sho2 = "CHWP"
         Case "HS5"
          get_sho2 = "CHW柱"
         Case "HS6"
          get_sho2 = "CHWレ"
         Case "HS7"
          get_sho2 = "CHS横"
         Case "HT1"
          get_sho2 = "CHITク"
         Case "HT2"
          get_sho2 = "CHIT竪"
         Case "HT3"
          get_sho2 = "CHIT横"
         Case "KK1"
          get_sho2 = "C腰壁ク"
         Case "KK2"
          get_sho2 = "C腰壁竪"
         Case "KK3"
          get_sho2 = "C腰壁横"
         Case "KK4"
          get_sho2 = "A腰壁支"
         Case "NS1"
          get_sho2 = "C2辺ク"
         Case "NS2"
          get_sho2 = "C2辺支"
         Case "FT1"
          get_sho2 = "BEX補"
    
         Case Else
            If Sho Like "S*" Then
                get_sho2 = "A受注"
            ElseIf Sho Like "M*" Then
                get_sho2 = "Cメイク"
            ElseIf Sho Like "G*" Then
                get_sho2 = "C柱"
            End If
        End Select
   Else
        Select Case Sho
         Case "XX"
          get_sho2 = "A規格"
         Case "P1", "P1A"
          get_sho2 = "A規格"
         Case "P2"
          get_sho2 = "A規格"
         Case "P3", "P3A"
          get_sho2 = "A規格"
         Case "P4", "P4A"
          get_sho2 = "A規格"
         Case "P5"
          get_sho2 = "A規格"
         Case "HS3"
          get_sho2 = "A規格"
         Case "HT4"
            If hen = "PD11" Then
              get_sho2 = "A規格"
            ElseIf hen = "PD21" Then
              get_sho2 = "C規格"
            End If
         Case "SM1"
          get_sho2 = "A規格"
    
         Case "HS1"
          get_sho2 = "C規格"
         Case "HS2"
          get_sho2 = "C規格"
         Case "HS4"
          get_sho2 = "C規格"
         Case "HS5"
          get_sho2 = "C規格"
         Case "HS6"
          get_sho2 = "C規格"
         Case "HS7"
          get_sho2 = "C規格"
         Case "HT1"
          get_sho2 = "C規格"
         Case "HT2"
          get_sho2 = "C規格"
         Case "HT3"
          get_sho2 = "C規格"
         Case "KK1"
          get_sho2 = "C規格"
         Case "KK2"
          get_sho2 = "C規格"
         Case "KK3"
          get_sho2 = "C規格"
         Case "KK4"
          get_sho2 = "A規格"
         Case "NS1"
          get_sho2 = "C規格"
         Case "NS2"
          get_sho2 = "C規格"
         Case "FT1"
          get_sho2 = "BEX定"
    
         Case Else
            If Sho Like "S*" Then
                get_sho2 = "A規格"
            ElseIf Sho Like "M*" Then
                get_sho2 = "C規格"
            ElseIf Sho Like "G*" Then
                get_sho2 = "C規格"
            End If
        End Select
    End If
End Function

