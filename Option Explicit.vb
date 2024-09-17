Option Explicit
    Public gameflg As Integer
    Public changeflg As Integer
    Public Steam As String
    Public Spos As String
    Public Sname As String
    Public SrArray() As Integer
    Public Shit As Integer
    Public Serror As Integer
    Public Slob As Integer
    Public avg As String
    Public hr As String
    Public rbi As String
    Public obp As String
    Public anda As String
    Public ops As String
    Public pitch As Integer
    Public inpitch As Integer
    Public k As Integer
    Public hian As Integer
    Public si As Integer
    Public pitchinning As String
    Public Stoday1 As String
    Public Stoday2 As String
    Public Stoday3 As String
    Public Stoday4 As String
    Public Stoday5 As String
    Public Stoday6 As String
    Public Kteam As String
    Public Kpos As String
    Public Kname As String
    Public KrArray() As Integer
    Public Khit As Integer
    Public Kerror As Integer
    Public Klob As Integer
    Public Ktoday1 As String
    Public Ktoday2 As String
    Public Ktoday3 As String
    Public Ktoday4 As String
    Public Ktoday5 As String
    Public Ktoday6 As String
    Public Anamep As String
    Public Aname1 As String
    Public Aname2 As String
    Public Aname3 As String
    Public Anamel As String
    Public Anamer As String
    Public i As Integer
    Public Maxrow As Integer
    Public ball As Integer
    Public strike As Integer
    Public out As Integer
    Public nowinning As Integer
    Public nowbatter As Integer
    Public nowattack As Integer
    Public strcnt As Integer
    Public fontname As String
    Public jogen As Integer
    Public seban As String
    Public seban_central As String
    Public seban_pacific As String
    Public namae As String
    Public r As Integer
    Public souruisi1 As String
    Public souruisi2 As String
    Public souruisi3 As String
    Public seikan1 As String
    Public seikan2 As String
    Public seikan3 As String
    Public seikan4 As String
    Public kiroku As String
    Public eumu As Integer
    Public touruisi As String
    Public run As Integer
    Public intcnt As Integer
    Public seikasan As String
    Public daten As Integer
    Public proama As Integer
    Public l As Integer
    Public a As Integer
    Public b As Integer
    Public sw As Integer
    Public runner1 As Integer
    Public runner2 As Integer
    Public runner3 As Integer
    Public result As Integer
    Public pos As String
    Public returnflg As Integer
    '標準
    Public basic_pos_Interior As Interior
    Public basic_pos_Font As Long
    Public basic_pos_Borders As Borders
    Public basic_pos_Color1 As Long
    Public basic_pos_Color2 As Long
    Public basic_name_Interior As Interior
    Public basic_name_Font As Long
    Public basic_name_Borders As Borders
    Public basic_name_Color1 As Long
    Public basic_name_Color2 As Long
    '打者
    Public batter_pos_Interior As Interior
    Public batter_pos_Font As Long
    Public batter_pos_Borders As Borders
    Public batter_pos_Color1 As Long
    Public batter_pos_Color2 As Long
    Public batter_name_Interior As Interior
    Public batter_name_Font As Long
    Public batter_name_Borders As Borders
    Public batter_name_Color1 As Long
    Public batter_name_Color2 As Long
    '走者
    Public runner_pos_Interior As Interior
    Public runner_pos_Font As Long
    Public runner_pos_Borders As Borders
    Public runner_pos_Color1 As Long
    Public runner_pos_Color2 As Long
    Public runner_name_Interior As Interior
    Public runner_name_Font As Long
    Public runner_name_Borders As Borders
    Public runner_name_Color1 As Long
    Public runner_name_Color2 As Long
    '次回先頭
    Public next_pos_Interior As Interior
    Public next_pos_Font As Long
    Public next_pos_Borders As Borders
    Public next_pos_Color1 As Long
    Public next_pos_Color2 As Long
    Public next_name_Interior As Interior
    Public next_name_Font As Long
    Public next_name_Borders As Borders
    Public next_name_Color1 As Long
    Public next_name_Color2 As Long
    'スコア
    Public score_inning_Interior As Interior
    Public score_inning_Font As Long
    Public score_inning_Borders As Borders
    Public score_inning_Color1 As Long
    Public score_inning_Color2 As Long
    Public score_entity_Interior As Interior
    Public score_entity_Font As Long
    Public score_entity_Borders As Borders
    Public score_entity_Color1 As Long
    Public score_entity_Color2 As Long
    '成績
    Public grade_item_Interior As Interior
    Public grade_item_Font As Long
    Public grade_item_Borders As Borders
    Public grade_item_Color1 As Long
    Public grade_item_Color2 As Long
    Public grade_entity_Interior As Interior
    Public grade_entity_Font As Long
    Public grade_entity_Borders As Borders
    Public grade_entity_Color1 As Long
    Public grade_entity_Color2 As Long
    'チーム名
    Public team_Interior As Interior
    Public team_Font As Long
    Public team_Borders As Borders
    Public team_Color1 As Long
    Public team_Color2 As Long
    '審判
    Public umpire_Interior As Interior
    Public umpire_Font As Long
    Public umpire_Borders As Borders
    Public umpire_Color1 As Long
    Public umpire_Color2 As Long
    '背景色
    Public back_Interior As Long
    '打席結果(安打)
    Public result_hit_Interior As Interior
    Public result_hit_Font As Long
    Public result_hit_Borders As Borders
    Public result_hit_Color1 As Long
    Public result_hit_Color2 As Long
    '打席結果(凡打)
    Public result_out_Interior As Interior
    Public result_out_Font As Long
    Public result_out_Borders As Borders
    Public result_out_Color1 As Long
    Public result_out_Color2 As Long
    '打席結果(その他)
    Public result_other_Interior As Interior
    Public result_other_Font As Long
    Public result_other_Borders As Borders
    Public result_other_Color1 As Long
    Public result_other_Color2 As Long
    '交代(消滅)
    Public dis_pos_Interior As Interior
    Public dis_pos_Font As Long
    Public dis_pos_Borders As Borders
    Public dis_pos_Color1 As Long
    Public dis_pos_Color2 As Long
    Public dis_name_Interior As Interior
    Public dis_name_Font As Long
    Public dis_name_Borders As Borders
    Public dis_name_Color1 As Long
    Public dis_name_Color2 As Long
    '交代(点灯)
    Public app_pos_Interior As Interior
    Public app_pos_Font As Long
    Public app_pos_Borders As Borders
    Public app_pos_Color1 As Long
    Public app_pos_Color2 As Long
    Public app_name_Interior As Interior
    Public app_name_Font As Long
    Public app_name_Borders As Borders
    Public app_name_Color1 As Long
    Public app_name_Color2 As Long
Sub Start()
    
    If gameflg <> 1 Then
        proama = 0
        returnflg = 0
        Steam = InputBox("先攻チーム")
        Kteam = InputBox("後攻チーム")
        
        jogen = Val(InputBox("イニング上限"))
        seikasan = InputBox("成績加算")
        
        Call 設定取得
        Range("CW10:HY10,EL113:FG127").Font.Color = back_Interior
        Range("B2:KW138").Interior.Color = back_Interior
        Range("M23:X122,IO23:IZ122").Font.Color = back_Interior
        Call returnrange("Steam1", team_Interior, team_Font, team_Borders, team_Color1, team_Color2)
        Call returnrange("Kteam1", team_Interior, team_Font, team_Borders, team_Color1, team_Color2)
        
        If Dir(ThisWorkbook.Path & "\スコアボード用ロゴ\" & Steam & ".png") <> "" Then
            ActiveSheet.Pictures.Insert( _
            ThisWorkbook.Path & "\スコアボード用ロゴ\" & Steam & ".png").Select
        Else
            If Len(Steam) = LenB(StrConv(Steam, vbFromUnicode)) Then
                Range("Steam1").HorizontalAlignment = xlCenter
                strcnt = Len(Steam)
                Call Fonta
            Else
                Range("Steam1").HorizontalAlignment = xlDistributed
                strcnt = Len(Steam)
                Call Font
            End If
            Range("Steam1").Font.Name = fontname
            Range("Steam1").Value = Steam
        End If
        
        If Steam = "広島" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(218, 0, 0)
                Range("Steam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "巨人" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(255, 128, 0)
                Range("Steam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "ヤクルト" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(28, 47, 142)
                Range("Steam1").Font.Color = RGB(0, 255, 0)
            End If
        ElseIf Steam = "DeNA" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(0, 73, 143)
                Range("Steam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "阪神" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(247, 206, 0)
                Range("Steam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "中日" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(4, 49, 148)
                Range("Steam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "ソフトバンク" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(243, 192, 0)
                Range("Steam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "西武" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(0, 33, 75)
                Range("Steam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "日本ハム" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(255, 255, 255)
                Range("Steam1").Font.Color = RGB(1, 96, 153)
            End If
        ElseIf Steam = "ロッテ" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(255, 255, 255)
                Range("Steam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "オリックス" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(0, 3, 31)
                Range("Steam1").Font.Color = RGB(182, 165, 75)
            End If
        ElseIf Steam = "楽天" Then
            If Range("teamname").Value = "" Then
                Range("Steam1").Interior.Color = RGB(125, 5, 28)
                Range("Steam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "セ" Then
            Range("Steam1").Interior.Color = RGB(15, 143, 46)
                Range("Steam1").Font.Color = RGB(255, 255, 255)
        ElseIf Steam = "パ" Then
            Range("Steam1").Interior.Color = RGB(63, 177, 229)
                Range("Steam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "ドジャース" Then
    If Range("teamname").Value = "" Then
        Range("Steam2").Interior.Color = RGB(0, 90, 156)
        Range("Steam2").Font.Color = RGB(255, 255, 255)
    End If
End If

        
        If Range("Steam1").Value = "" And proama <> 1 Then
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.ShapeRange.Height = Range("M8:M20").Height
            '調整
            If Selection.ShapeRange.Width > Range("M8:BJ8").Width Then
                '幅
                Selection.ShapeRange.Width = Range("M8:BJ8").Width
                '左位置
                Selection.ShapeRange.Left = Range("Steam1").Left
                '縦位置
                Selection.ShapeRange.Top = Range("M8:M20").Top + _
                (Range("M8:M20").Height - Selection.ShapeRange.Height) / 2
            Else
                Selection.ShapeRange.Top = Range("Steam1").Top
                Selection.ShapeRange.Left = Range("M8:BJ8").Left + _
                (Range("M8:BJ8").Width - Selection.ShapeRange.Width) / 2
            End If
        End If
        
        If Dir(ThisWorkbook.Path & "\スコアボード用ロゴ\" & Kteam & ".png") <> "" Then
            ActiveSheet.Pictures.Insert( _
            ThisWorkbook.Path & "\スコアボード用ロゴ\" & Kteam & ".png").Select
        Else
            If Len(Kteam) = LenB(StrConv(Kteam, vbFromUnicode)) Then
                Range("Kteam1").HorizontalAlignment = xlCenter
                strcnt = Len(Kteam)
                Call Fonta
            Else
                Range("Kteam1").HorizontalAlignment = xlDistributed
                strcnt = Len(Kteam)
                Call Font
            End If
            Range("Kteam1").Font.Name = fontname
            Range("Kteam1").Value = Kteam
        End If
        
        proama = 0
        If Kteam = "広島" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(218, 0, 0)
                Range("Kteam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "巨人" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(255, 128, 0)
                Range("Kteam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "ヤクルト" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(28, 47, 142)
                Range("Kteam1").Font.Color = RGB(0, 255, 0)
            End If
        ElseIf Kteam = "DeNA" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(0, 73, 143)
                Range("Kteam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "阪神" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(247, 206, 0)
                Range("Kteam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "中日" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(4, 49, 148)
                Range("Kteam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "ソフトバンク" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(243, 192, 0)
                Range("Kteam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "西武" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(0, 33, 75)
                Range("Kteam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "日本ハム" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(255, 255, 255)
                Range("Kteam1").Font.Color = RGB(1, 96, 153)
            End If
        ElseIf Kteam = "ロッテ" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(255, 255, 255)
                Range("Kteam1").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "オリックス" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(0, 3, 31)
                Range("Kteam1").Font.Color = RGB(182, 165, 75)
            End If
        ElseIf Kteam = "楽天" Then
            If Range("teamname").Value = "" Then
                Range("Kteam1").Interior.Color = RGB(125, 5, 28)
                Range("Kteam1").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "セ" Then
            Range("Kteam1").Interior.Color = RGB(15, 143, 46)
                Range("Kteam1").Font.Color = RGB(255, 255, 255)
        ElseIf Kteam = "パ" Then
            Range("Kteam1").Interior.Color = RGB(63, 177, 229)
                Range("Kteam1").Font.Color = RGB(255, 255, 255)
        End If
ElseIf Steam = "ドジャース" Then
    If Range("teamname").Value = "" Then
        Range("Steam2").Interior.Color = RGB(0, 90, 156)
        Range("Steam2").Font.Color = RGB(255, 255, 255)
    End If
End If

        
        If Range("Kteam1").Value = "" And proama <> 1 Then
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.ShapeRange.Height = Range("IO8:KL20").Height
            '調整
            If Selection.ShapeRange.Width > Range("IO8:KL8").Width Then
                '幅
                Selection.ShapeRange.Width = Range("IO8:KL8").Width
                '左位置
                Selection.ShapeRange.Left = Range("Kteam1").Left
                '縦位置
                Selection.ShapeRange.Top = Range("IO8:IP20").Top + _
                (Range("IO8:IP20").Height - Selection.ShapeRange.Height) / 2
            Else
            Selection.ShapeRange.Top = Range("Kteam1").Top
            Selection.ShapeRange.Left = Range("IO8:KL8").Left + _
            (Range("IO8:KL8").Width - Selection.ShapeRange.Width) / 2
            End If
        End If
        Range("LD153").Select
    End If
End Sub
Sub Font()
    If strcnt <= 4 Then
        fontname = "IPAexゴシック"
    ElseIf strcnt <= 6 Then
        fontname = "TTEdit2/3ゴシック"
    Else
        fontname = "TTEdit半角ゴシック"
    End If
End Sub
Sub Fonta()
    If strcnt <= 8 Then
        fontname = "IPAexゴシック"
    ElseIf strcnt <= 12 Then
        fontname = "TTEdit2/3ゴシック"
    Else
        fontname = "TTEdit半角ゴシック"
    End If
End Sub
Sub Font_sei()
    If strcnt <= 11 Then
        fontname = "IPAexゴシック"
    ElseIf strcnt <= 17 Then
        fontname = "TTEdit2/3ゴシック"
    Else
        fontname = "TTEdit半角ゴシック"
    End If
End Sub
Sub Fonta_sei()
    If strcnt <= 10 Then
        fontname = "IPAexゴシック"
    ElseIf strcnt <= 16 Then
        fontname = "TTEdit2/3ゴシック"
    Else
        fontname = "TTEdit半角ゴシック"
    End If
End Sub
Sub IntFont()
    
    If intcnt > 1 Then
        fontname = "Arial Narrow"
    Else
        fontname = "Arial"
    End If
    
End Sub
Sub Umpire()
    Anamep = InputBox("球審")
    Aname1 = InputBox("一塁塁審")
    Aname2 = InputBox("二塁塁審")
    Aname3 = InputBox("三塁塁審")
    Anamel = InputBox("左翼線審")
    Anamer = InputBox("右翼線審")

    If Anamel = "" Then
        If IsNull(Range("umpire").MergeArea.Borders.LineStyle) Or Range("umpire").MergeArea.Borders.LineStyle <> -4142 Then
            Call returnrange("Anamep", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Aname1", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Aname2", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Aname3", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
        Else
            returnflg = 1
            Call returnrange("Anamep,Aname1,Aname2,Aname3", umpire_Interior, umpire_Font, _
                             umpire_Borders, umpire_Color1, umpire_Color2)
        End If
        Range("Anamep").Value = Anamep
        Range("Aname1").Value = Aname1
        Range("Aname2").Value = Aname2
        Range("Aname3").Value = Aname3
    Else
        If IsNull(Range("umpire").MergeArea.Borders.LineStyle) Or Range("umpire").MergeArea.Borders.LineStyle <> -4142 Then
            Call returnrange("Anamep", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Aname1", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Aname2", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Aname3", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Anamel", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
            Call returnrange("Anamer", umpire_Interior, umpire_Font, umpire_Borders, _
                             umpire_Color1, umpire_Color2)
        Else
            returnflg = 1
            Call returnrange("Anamep,Aname1,Aname2,Aname3,Anamel,Anamer", umpire_Interior, _
                             umpire_Font, umpire_Borders, umpire_Color1, umpire_Color2)
        End If
        Range("Anamep").Value = Anamep
        Range("Aname1").Value = Aname1
        Range("Aname2").Value = Aname2
        Range("Aname3").Value = Aname3
        Range("Anamel").Value = Anamel
        Range("Anamer").Value = Anamer
    End If
End Sub
Sub gamestart()
    
    
    If gameflg <> 1 Then
        gameflg = 1
        nowinning = 1
        nowattack = 1
        nowbatter = 1
        
        Range("inin1").Value = "1"
        Range("inin2").Value = "2"
        Range("inin3").Value = "3"
        Range("inin4").Value = "4"
        Range("inin5").Value = "5"
        Range("inin6").Value = "6"
        Range("inin7").Value = "7"
        Range("inin8").Value = "8"
        Range("inin9").Value = "9"
        Range("ininr").Value = "R"
        Range("ininh").Value = "H"
        Range("inine").Value = "E"
        Range("ininlob").Value = "LOB"
        
        Range("balln").Value = "B"
        Range("striken").Value = "S"
        Range("outn").Value = "O"
        Range("ball1,ball2,ball3,strike1,strike2,auto1,auto2").Value = "●"
        If back_Interior = -4142 Then
            Range("EL113:FG127").Interior.Color = RGB(0, 0, 0)
        End If
        
        Call returnrange("Spos1", batter_pos_Interior, batter_pos_Font, batter_pos_Borders, _
                          batter_pos_Color1, batter_pos_Color2)
        Call returnrange("Sname1", batter_name_Interior, batter_name_Font, batter_name_Borders, _
                          batter_name_Color1, batter_name_Color2)
        Range("Sline").Interior.Color = RGB(255, 0, 0)
        Call returnrange("Kpos1", next_pos_Interior, next_pos_Font, next_pos_Borders, _
                          next_pos_Color1, next_pos_Color2)
        Call returnrange("Kname1", next_name_Interior, next_name_Font, next_name_Borders, _
                          next_name_Color1, next_name_Color2)
        
        Call returnrange("Steam2", score_entity_Interior, score_entity_Font, score_entity_Borders, score_entity_Color1, score_entity_Color2)
        Call returnrange("Kteam2", score_entity_Interior, score_entity_Font, score_entity_Borders, score_entity_Color1, score_entity_Color2)
        Call ScoreCellReset
        If IsNull(Range("scoreinin").MergeArea.Borders.LineStyle) Or Range("scoreinin").MergeArea.Borders.LineStyle <> -4142 Then
            Call returnrange("BP10", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin1", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin2", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin3", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin4", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin5", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin6", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin7", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin8", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inin9", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("ininr", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("ininh", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("inine", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
            Call returnrange("ininlob", score_inning_Interior, score_inning_Font, score_inning_Borders, _
                             score_inning_Color1, score_inning_Color2)
        Else
            returnflg = 1
            Call returnrange("BP10,inin1,inin2,inin3,inin4,inin5,inin6,inin7,inin8,inin9,ininr,ininh,inine,ininlob", _
                             score_inning_Interior, score_inning_Font, score_inning_Borders, score_inning_Color1, score_inning_Color2)
        End If
        Range("Srun1").Interior.Color = RGB(255, 0, 0)
        Range("Srun1").Font.Color = RGB(255, 255, 255)
        Range("balln,striken,outn").Font.Color = RGB(255, 255, 255)
        Range("ball1,ball2,ball3,strike1,strike2,auto1,auto2").Font.Color = RGB(38, 38, 38)
        ReDim SrArray(2)
        Range("Srunr").Value = CStr(SrArray(0))
        Range("Shit").Value = CStr(Shit)
        Range("Serror").Value = CStr(Serror)
        Range("Slob").Value = CStr(Slob)
        ReDim KrArray(2)
        Range("Krunr").Value = CStr(KrArray(0))
        Range("Khit").Value = CStr(Khit)
        Range("Kerror").Value = CStr(Kerror)
        Range("Klob").Value = CStr(Klob)
        Call 成績表示
        Call 投手成績表示
        ActiveSheet.Shapes("一塁").Visible = True
        ActiveSheet.Shapes("二塁").Visible = True
        ActiveSheet.Shapes("三塁").Visible = True

        If Dir(ThisWorkbook.Path & "\スコアボード用ロゴ\" & Steam & ".png") <> "" Then
            ActiveSheet.Pictures.Insert( _
            ThisWorkbook.Path & "\スコアボード用ロゴ\" & Steam & ".png").Select
        Else
            If Len(Steam) = LenB(StrConv(Steam, vbFromUnicode)) Then
                Range("Steam2").HorizontalAlignment = xlCenter
                strcnt = Len(Steam)
                Call Fonta
            Else
                Range("Steam2").HorizontalAlignment = xlDistributed
                strcnt = Len(Steam)
                Call Font
            End If
            Range("Steam2").Font.Name = fontname
            Range("Steam2").Value = Steam
        End If

        If Steam = "広島" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(218, 0, 0)
                Range("Steam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "巨人" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(255, 128, 0)
                Range("Steam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "ヤクルト" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(28, 47, 142)
                Range("Steam2").Font.Color = RGB(0, 255, 0)
            End If
        ElseIf Steam = "DeNA" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(0, 73, 143)
                Range("Steam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "阪神" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(247, 206, 0)
                Range("Steam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "中日" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(4, 49, 148)
                Range("Steam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "ソフトバンク" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(243, 192, 0)
                Range("Steam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "西武" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(0, 33, 75)
                Range("Steam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "日本ハム" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(255, 255, 255)
                Range("Steam2").Font.Color = RGB(1, 96, 153)
            End If
        ElseIf Steam = "ロッテ" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(255, 255, 255)
                Range("Steam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Steam = "オリックス" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(0, 3, 31)
                Range("Steam2").Font.Color = RGB(182, 165, 75)
            End If
        ElseIf Steam = "楽天" Then
            If Range("teamname").Value = "" Then
                Range("Steam2").Interior.Color = RGB(125, 5, 28)
                Range("Steam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "セ" Then
            Range("Steam2").Interior.Color = RGB(15, 143, 46)
                Range("Steam2").Font.Color = RGB(255, 255, 255)
        ElseIf Steam = "パ" Then
            Range("Steam2").Interior.Color = RGB(63, 177, 229)
                Range("Steam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Steam = "ドジャース" Then
    If Range("teamname").Value = "" Then
        Range("Steam2").Interior.Color = RGB(0, 90, 156)
        Range("Steam2").Font.Color = RGB(255, 255, 255)
    End If
End If

        
        If Range("Steam2").Value = "" Then
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.ShapeRange.Height = Range("BP17:CV27").Height
            '調整
            If Selection.ShapeRange.Width > Range("BP17:CV17").Width Then
                '幅
                Selection.ShapeRange.Width = Range("BP17:CV17").Width
                '左位置
                Selection.ShapeRange.Left = Range("Steam2").Left
                '縦位置
                Selection.ShapeRange.Top = Range("BP17:CV27").Top + _
                (Range("BP17:CV27").Height - Selection.ShapeRange.Height) / 2
            Else
                Selection.ShapeRange.Top = Range("Steam2").Top
                Selection.ShapeRange.Left = Range("BP17:CV27").Left + _
                (Range("BP17:CV27").Width - Selection.ShapeRange.Width) / 2
            End If
        End If
        
        If Dir(ThisWorkbook.Path & "\スコアボード用ロゴ\" & Kteam & ".png") <> "" Then
            ActiveSheet.Pictures.Insert( _
            ThisWorkbook.Path & "\スコアボード用ロゴ\" & Kteam & ".png").Select
        Else
            If Len(Kteam) = LenB(StrConv(Kteam, vbFromUnicode)) Then
                Range("Kteam2").HorizontalAlignment = xlCenter
                strcnt = Len(Kteam)
                Call Fonta
            Else
                Range("Kteam2").HorizontalAlignment = xlDistributed
                strcnt = Len(Kteam)
                Call Font
            End If
            Range("Kteam2").Font.Name = fontname
            Range("Kteam2").Value = Kteam
        End If
        
        If Kteam = "広島" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(218, 0, 0)
                Range("Kteam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "巨人" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(255, 128, 0)
                Range("Kteam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "ヤクルト" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(28, 47, 142)
                Range("Kteam2").Font.Color = RGB(0, 255, 0)
            End If
        ElseIf Kteam = "DeNA" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(0, 73, 143)
                Range("Kteam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "阪神" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(247, 206, 0)
                Range("Kteam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "中日" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(4, 49, 148)
                Range("Kteam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "ソフトバンク" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(243, 192, 0)
                Range("Kteam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "西武" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(0, 33, 75)
                Range("Kteam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "日本ハム" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(255, 255, 255)
                Range("Kteam2").Font.Color = RGB(1, 96, 153)
            End If
        ElseIf Kteam = "ロッテ" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(255, 255, 255)
                Range("Kteam2").Font.Color = RGB(0, 0, 0)
            End If
        ElseIf Kteam = "オリックス" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(0, 3, 31)
                Range("Kteam2").Font.Color = RGB(182, 165, 75)
            End If
        ElseIf Kteam = "楽天" Then
            If Range("teamname").Value = "" Then
                Range("Kteam2").Interior.Color = RGB(125, 5, 28)
                Range("Kteam2").Font.Color = RGB(255, 255, 255)
            End If
        ElseIf Kteam = "セ" Then
            Range("Kteam2").Interior.Color = RGB(15, 143, 46)
                Range("Kteam2").Font.Color = RGB(255, 255, 255)
        ElseIf Kteam = "パ" Then
            Range("Kteam2").Interior.Color = RGB(63, 177, 229)
                Range("Kteam2").Font.Color = RGB(255, 255, 255)
        End If
                ElseIf Steam = "ドジャース" Then
    If Range("teamname").Value = "" Then
        Range("Steam2").Interior.Color = RGB(0, 90, 156)
        Range("Steam2").Font.Color = RGB(255, 255, 255)
    End If
End If
        
        If Range("Kteam2").Value = "" Then
            Selection.ShapeRange.LockAspectRatio = msoTrue
            Selection.ShapeRange.Height = Range("BP28:CV38").Height
            '調整
            If Selection.ShapeRange.Width > Range("BP28:CV28").Width Then
                '幅
                Selection.ShapeRange.Width = Range("BP28:CV28").Width
                '左位置
                Selection.ShapeRange.Left = Range("Kteam2").Left
                '縦位置
                Selection.ShapeRange.Top = Range("BP28:CV38").Top + _
                (Range("BP28:CV38").Height - Selection.ShapeRange.Height) / 2
            Else
                Selection.ShapeRange.Top = Range("Kteam2").Top
                Selection.ShapeRange.Left = Range("BP28:CV38").Left + _
                (Range("BP28:CV38").Width - Selection.ShapeRange.Width) / 2
            End If
        End If
        
        Range("LD153").Select
        
    End If
End Sub

Sub Yuちゃん勝負眼表示()
    If nowattack = 1 Then
        ActiveSheet.Shapes("Yuちゃん").Top = Range("FR110").Top
        ActiveSheet.Shapes("Yuちゃん").Left = Range("FR110").Left
    Else
        ActiveSheet.Shapes("Yuちゃん").Top = Range("BP110").Top
        ActiveSheet.Shapes("Yuちゃん").Left = Range("BP110").Left
    End If
End Sub

Sub Yuちゃん勝負眼消去()
    ActiveSheet.Shapes("Yuちゃん").Top = Range("CZ192").Top
    ActiveSheet.Shapes("Yuちゃん").Left = Range("CZ192").Left
End Sub

Sub なかちゃん勝負眼表示()
    If nowattack = 1 Then
        ActiveSheet.Shapes("なかちゃん").Top = Range("GO110").Top
        ActiveSheet.Shapes("なかちゃん").Left = Range("GO110").Left
    Else
        ActiveSheet.Shapes("なかちゃん").Top = Range("CM110").Top
        ActiveSheet.Shapes("なかちゃん").Left = Range("CM110").Left
    End If
End Sub

Sub なかちゃん勝負眼消去()
    ActiveSheet.Shapes("なかちゃん").Top = Range("DY192").Top
    ActiveSheet.Shapes("なかちゃん").Left = Range("DY192").Left
End Sub

Sub シャーリーン勝負眼表示()
    If nowattack = 1 Then
        ActiveSheet.Shapes("シャーリーン").Top = Range("HM110").Top
        ActiveSheet.Shapes("シャーリーン").Left = Range("HM110").Left
    Else
        ActiveSheet.Shapes("シャーリーン").Top = Range("DK110").Top
        ActiveSheet.Shapes("シャーリーン").Left = Range("DK110").Left
    End If
End Sub

Sub シャーリーン勝負眼消去()
    ActiveSheet.Shapes("シャーリーン").Top = Range("EX192").Top
    ActiveSheet.Shapes("シャーリーン").Left = Range("EX192").Left
End Sub

Sub さとみちゃん勝負眼表示()
    ActiveSheet.Shapes("さとみちゃん").Top = Range("DQ46").Top
    ActiveSheet.Shapes("さとみちゃん").Left = Range("DQ46").Left
End Sub

Sub さとみちゃん勝負眼消去()
    ActiveSheet.Shapes("さとみちゃん").Top = Range("FW192").Top
    ActiveSheet.Shapes("さとみちゃん").Left = Range("FW192").Left
End Sub

Sub まさしさん勝負眼表示()
    ActiveSheet.Shapes("まさしさん").Top = Range("FK46").Top
    ActiveSheet.Shapes("まさしさん").Left = Range("FK46").Left
End Sub

Sub まさしさん勝負眼消去()
    ActiveSheet.Shapes("まさしさん").Top = Range("GW192").Top
    ActiveSheet.Shapes("まさしさん").Left = Range("GW192").Left
End Sub

Sub リプレイ検証表示()
    ActiveSheet.Shapes("リプレイ").Top = Range("DR46").Top
    ActiveSheet.Shapes("リプレイ").Left = Range("DR46").Left
End Sub

Sub リプレイ検証消去()
    ActiveSheet.Shapes("リプレイ").Top = Range("FW192").Top
    ActiveSheet.Shapes("リプレイ").Left = Range("FW192").Left
End Sub

Sub 次打者()
        
    With Worksheets("試合情報")
        .Range("H6:H8").Value = ""
        .Range("F3").Value = nowinning
        .Range("F4").Value = nowattack
        .Range("F5").Value = nowbatter
        .Range("F6").Value = ball
        .Range("F7").Value = strike
        .Range("F8").Value = out
        .Range("H3").Value = kiroku
        .Range("H4").Value = daten
        .Range("H5").Value = hian
        .Range("J3").Value = pitch
        .Range("J4").Value = k
        .Range("J5").Value = si
        .Range("J6").Value = pitchinning
        .Range("J7").Value = inpitch
    End With
    
    a = 23
    b = 1
    
    runner1 = 0
    runner2 = 0
    runner3 = 0
    
    If nowattack = 1 Then
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern = 1 Then
                If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 13).Interior.Pattern <> 1 And Cells(a, 24).Interior.Pattern = 1 Then
                If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern <> 1 Then
                If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            Else
                If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            End If
            b = b + 1
        Next a
    Else
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 249).Interior.Pattern <> 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern <> 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            Else
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            End If
            b = b + 1
        Next a
    End If
    
    With Worksheets("試合情報")
        If runner1 <> 0 Then
            .Range("H6").Value = runner1
            If runner2 <> 0 Then
                .Range("H7").Value = runner2
                If runner3 <> 0 Then
                    .Range("H8").Value = runner3
                End If
            End If
        End If
    End With
    
    If nowattack = 1 Then
        If result = 0 Then
            Call returncells(23 + ((nowbatter - 1) * 11), 13, basic_pos_Interior, basic_pos_Font, _
                             basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 24, basic_name_Interior, basic_name_Font, _
                             basic_name_Borders, basic_name_Color1, basic_name_Color2)
        End If
        If nowbatter = 9 Then
            nowbatter = 1
        Else
            nowbatter = nowbatter + 1
        End If
        Call returncells(23 + ((nowbatter - 1) * 11), 13, batter_pos_Interior, batter_pos_Font, _
                         batter_pos_Borders, batter_pos_Color1, batter_pos_Color2)
        Call returncells(23 + ((nowbatter - 1) * 11), 24, batter_name_Interior, batter_name_Font, _
                         batter_name_Borders, batter_name_Color1, batter_name_Color2)
    Else
        If result = 0 Then
            Call returncells(23 + ((nowbatter - 1) * 11), 249, basic_pos_Interior, basic_pos_Font, _
                             basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 260, basic_name_Interior, basic_name_Font, _
                             basic_name_Borders, basic_name_Color1, basic_name_Color2)
        End If
        If nowbatter = 9 Then
            nowbatter = 1
        Else
            nowbatter = nowbatter + 1
        End If
        Call returncells(23 + ((nowbatter - 1) * 11), 249, batter_pos_Interior, batter_pos_Font, _
                         batter_pos_Borders, batter_pos_Color1, batter_pos_Color2)
        Call returncells(23 + ((nowbatter - 1) * 11), 260, batter_name_Interior, batter_name_Font, _
                         batter_name_Borders, batter_name_Color1, batter_name_Color2)
    End If
    inpitch = 0
End Sub
Sub 成績表示()
    
    Dim arr()
    If nowattack = 1 Then
        seban = CStr(Cells(143 + ((nowbatter - 1) * 4), 10).Value)
        
        arr() = getinfo(Steam, seban)
        
        If nowinning = 1 And nowattack = 1 And nowbatter = 1 Then
            If IsNull(Range("gradeitem").MergeArea.Borders.LineStyle) Or Range("gradeitem").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Ssein1", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein2", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein3", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein4", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein5", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein6", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
            Else
                returnflg = 1
                Call returnrange("Ssein1,Ssein2,Ssein3,Ssein4,Ssein5,Ssein6", _
                                 grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
            End If
            If IsNull(Range("gradeent").MergeArea.Borders.LineStyle) Or Range("gradeent").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Snum", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sname", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie1", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie2", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie3", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie4", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie5", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie6", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
            Else
                returnflg = 1
                Call returnrange("Snum,Sname,Sseie1,Sseie2,Sseie3,Sseie4,Sseie5,Sseie6", _
                                 grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
            End If
        End If
        If Steam = "セ" Or Steam = "パ" Then
            seban = arr(0)
            avg = arr(2)
            anda = arr(4)
            hr = arr(7)
            rbi = arr(8)
            obp = arr(9)
            ops = arr(10)
            namae = arr(14)
        Else
            avg = arr(1)
            anda = arr(3)
            hr = arr(6)
            rbi = arr(7)
            obp = arr(8)
            ops = arr(9)
            namae = arr(13)
        End If
            
        If Len(namae) = LenB(StrConv(namae, vbFromUnicode)) Then
            strcnt = Len(namae)
            Call Fonta_sei
        Else
            strcnt = LenB(StrConv(namae, vbFromUnicode))
            Call Font_sei
        End If
        Range("Sname").Font.Name = fontname
        Range("Ssein4").Font.Name = "IPAexゴシック"
        Range("Ssein1").Value = "打　率"
        Range("Ssein2").Value = "本塁打"
        Range("Ssein3").Value = "打　点"
        Range("Ssein4").Value = "安　打"
        Range("Ssein5").Value = "出塁率"
        Range("Ssein6").Value = "OPS"
        If avg <> "1.00" Then
            Range("Sseie1").NumberFormatLocal = ".000"
        Else
            Range("Sseie1").NumberFormatLocal = "@"
        End If
        If obp <> "1.00" Then
            Range("Sseie5").NumberFormatLocal = ".000"
        Else
            Range("Sseie5").NumberFormatLocal = "@"
        End If
        Range("Sseie6").NumberFormat = ".000"
        Range("Snum").Value = seban
        Range("Sname").Value = namae
        Range("Sseie1").Value = CStr(avg)
        Range("Sseie2").Value = CStr(hr)
        Range("Sseie3").Value = CStr(rbi)
        Range("Sseie4").Value = CStr(anda)
        Range("Sseie5").Value = CStr(obp)
        Range("Sseie6").Value = CStr(ops)
    Else
        seban = CStr(Cells(143 + ((nowbatter - 1) * 4), 281).Value)
        
        arr() = getinfo(Kteam, seban)
        
        If Kteam = "セ" Or Kteam = "パ" Then
            seban = arr(0)
            avg = arr(2)
            anda = arr(4)
            hr = arr(7)
            rbi = arr(8)
            obp = arr(9)
            ops = arr(10)
            namae = arr(14)
        Else
            avg = arr(1)
            anda = arr(3)
            hr = arr(6)
            rbi = arr(7)
            obp = arr(8)
            ops = arr(9)
            namae = arr(13)
        End If
            
        If Len(namae) = LenB(StrConv(namae, vbFromUnicode)) Then
            strcnt = Len(namae)
            Call Fonta_sei
        Else
            strcnt = LenB(StrConv(namae, vbFromUnicode))
            Call Font_sei
        End If
        Range("Kname").Font.Name = fontname
        Range("Ksein4").Font.Name = "IPAexゴシック"
        Range("Ksein1").Value = "打　率"
        Range("Ksein2").Value = "本塁打"
        Range("Ksein3").Value = "打　点"
        Range("Ksein4").Value = "安　打"
        Range("Ksein5").Value = "出塁率"
        Range("Ksein6").Value = "OPS"
        If avg <> "1.00" Then
            Range("Kseie1").NumberFormatLocal = ".000"
        Else
            Range("Kseie1").NumberFormatLocal = "@"
        End If
        If obp <> "1.00" Then
            Range("Kseie5").NumberFormatLocal = ".000"
        Else
            Range("Kseie5").NumberFormatLocal = "@"
        End If
        Range("Kseie6").NumberFormat = ".000"
        Range("Knum").Value = seban
        Range("Kname").Value = namae
        Range("Kseie1").Value = CStr(avg)
        Range("Kseie2").Value = CStr(hr)
        Range("Kseie3").Value = CStr(rbi)
        Range("Kseie4").Value = CStr(anda)
        Range("Kseie5").Value = CStr(obp)
        Range("Kseie6").Value = CStr(ops)
    End If
    
End Sub
Sub 投手成績表示()
    
    Dim arr()
    
    With Worksheets("試合情報")
    If nowattack = 1 Then
        If Range("text21").Value = "1" Or Range("text21").Value = "P" Then
            seban = CStr(Range("text31").Value)
        ElseIf Range("text22").Value = "1" Or Range("text22").Value = "P" Then
            seban = CStr(Range("text32").Value)
        ElseIf Range("text23").Value = "1" Or Range("text23").Value = "P" Then
            seban = CStr(Range("text33").Value)
        ElseIf Range("text24").Value = "1" Or Range("text24").Value = "P" Then
            seban = CStr(Range("text34").Value)
        ElseIf Range("text25").Value = "1" Or Range("text25").Value = "P" Then
            seban = CStr(Range("text35").Value)
        ElseIf Range("text26").Value = "1" Or Range("text26").Value = "P" Then
            seban = CStr(Range("text36").Value)
        ElseIf Range("text27").Value = "1" Or Range("text27").Value = "P" Then
            seban = CStr(Range("text37").Value)
        ElseIf Range("text28").Value = "1" Or Range("text28").Value = "P" Then
            seban = CStr(Range("text38").Value)
        ElseIf Range("text29").Value = "1" Or Range("text29").Value = "P" Then
            seban = CStr(Range("text39").Value)
        ElseIf Range("text30").Value = "1" Or Range("text30").Value = "P" Then
            seban = CStr(Range("text40").Value)
        Else
            seban = ""
        End If
        
        
        If seban <> "" Then
            arr() = getinfo(Kteam, seban)
            
            If nowinning = 1 And nowattack = 1 And nowbatter = 1 Then
                If IsNull(Range("gradeitem").MergeArea.Borders.LineStyle) Or Range("gradeitem").MergeArea.Borders.LineStyle <> -4142 Then
                    Call returnrange("Ksein1", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                     grade_item_Color1, grade_item_Color2)
                    Call returnrange("Ksein2", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                     grade_item_Color1, grade_item_Color2)
                    Call returnrange("Ksein3", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                     grade_item_Color1, grade_item_Color2)
                    Call returnrange("Ksein4", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                     grade_item_Color1, grade_item_Color2)
                    Call returnrange("Ksein5", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                     grade_item_Color1, grade_item_Color2)
                    Call returnrange("Ksein6", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                     grade_item_Color1, grade_item_Color2)
                Else
                    returnflg = 1
                    Call returnrange("Ksein1,Ksein2,Ksein3,Ksein4,Ksein5,Ksein6", _
                                     grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                     grade_item_Color1, grade_item_Color2)
                End If
                If IsNull(Range("gradeent").MergeArea.Borders.LineStyle) Or Range("gradeent").MergeArea.Borders.LineStyle <> -4142 Then
                    Call returnrange("Knum", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                    Call returnrange("Kname", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                    Call returnrange("Kseie1", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                    Call returnrange("Kseie2", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                    Call returnrange("Kseie3", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                    Call returnrange("Kseie4", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                    Call returnrange("Kseie5", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                    Call returnrange("Kseie6", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                Else
                    returnflg = 1
                    Call returnrange("Knum,Kname,Kseie1,Kseie2,Kseie3,Kseie4,Kseie5,Kseie6", _
                                     grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                     grade_entity_Color1, grade_entity_Color2)
                End If
            End If
            
            If Kteam = "セ" Or Kteam = "パ" Then
                seban = arr(0)
                namae = arr(14)
            Else
                namae = arr(13)
            End If
            
            pitchinning = .Range("C3").Value
            si = .Range("C4").Value
            pitch = .Range("C5").Value
            hian = .Range("C6").Value
            k = .Range("C7").Value
            
            strcnt = Len(namae)
            If Len(namae) = LenB(StrConv(namae, vbFromUnicode)) Then
                strcnt = Len(namae)
                Call Fonta_sei
            Else
                strcnt = LenB(StrConv(namae, vbFromUnicode))
                Call Font_sei
            End If
            Range("Kname").Font.Name = fontname
            If Kteam = "セ" Or Kteam = "パ" Then
                If Kteam = "セ" Then
                    Range("Knum").Value = seban_central
                Else
                    Range("Knum").Value = seban_pacific
                End If
            Else
                Range("Knum").Value = seban
            End If
            Range("Ksein4").Font.Name = "TTEdit2/3ゴシック"
            Range("Kname").Font.Name = fontname
            Range("Knum").Value = seban
            Range("Kname").Value = namae
            Range("Ksein1").Value = "投球数"
            Range("Ksein2").Value = "奪三振"
            Range("Ksein3").Value = "被安打"
            Range("Ksein4").Value = "打席内投球数"
            Range("Ksein5").Value = "与四球"
            Range("Ksein6").Value = "投球回"
            Range("Kseie1,Kseie5,Kseie6").NumberFormatLocal = "G/標準"
            Range("Kseie1").Value = pitch
            Range("Kseie2").Value = k
            Range("Kseie3").Value = hian
            Range("Kseie4").Value = 0
            Range("Kseie5").Value = si
            If pitchinning Mod 3 = 0 Then
                Range("Kseie6").Value = pitchinning / 3
            Else
                Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
        Else
            Range("Ksein4").Font.Name = "TTEdit2/3ゴシック"
            Range("Kname").Font.Name = fontname
            Range("Knum").Value = ""
            Range("Kname").Value = ""
            Range("Ksein1").Value = "投球数"
            Range("Ksein2").Value = "奪三振"
            Range("Ksein3").Value = "被安打"
            Range("Ksein4").Value = "打席内投球数"
            Range("Ksein5").Value = "与四球"
            Range("Ksein6").Value = "投球回"
            Range("Kseie1,Kseie5,Kseie6").NumberFormatLocal = "G/標準"
            Range("Kseie1").Value = 0
            Range("Kseie2").Value = 0
            Range("Kseie3").Value = 0
            Range("Kseie4").Value = 0
            Range("Kseie5").Value = 0
            Range("Kseie6").Value = 0
        End If
    Else
        If Range("text1").Value = "1" Or Range("text1").Value = "P" Then
            seban = CStr(Range("text11").Value)
        ElseIf Range("text2").Value = "1" Or Range("text2").Value = "P" Then
            seban = CStr(Range("text12").Value)
        ElseIf Range("text3").Value = "1" Or Range("text3").Value = "P" Then
            seban = CStr(Range("text13").Value)
        ElseIf Range("text4").Value = "1" Or Range("text4").Value = "P" Then
            seban = CStr(Range("text14").Value)
        ElseIf Range("text5").Value = "1" Or Range("text5").Value = "P" Then
            seban = CStr(Range("text15").Value)
        ElseIf Range("text6").Value = "1" Or Range("text6").Value = "P" Then
            seban = CStr(Range("text16").Value)
        ElseIf Range("text7").Value = "1" Or Range("text7").Value = "P" Then
            seban = CStr(Range("text17").Value)
        ElseIf Range("text8").Value = "1" Or Range("text8").Value = "P" Then
            seban = CStr(Range("text18").Value)
        ElseIf Range("text9").Value = "1" Or Range("text9").Value = "P" Then
            seban = CStr(Range("text19").Value)
        ElseIf Range("text10").Value = "1" Or Range("text10").Value = "P" Then
            seban = CStr(Range("text20").Value)
        Else
            seban = ""
        End If
         
        If seban <> "" Then
            arr() = getinfo(Steam, seban)
            
            If Steam = "セ" Or Steam = "パ" Then
                seban = arr(0)
                namae = arr(14)
            Else
                namae = arr(13)
            End If
            
            pitchinning = .Range("B3").Value
            si = .Range("B4").Value
            pitch = .Range("B5").Value
            hian = .Range("B6").Value
            k = .Range("B7").Value
            
            strcnt = Len(namae)
            If Len(namae) = LenB(StrConv(namae, vbFromUnicode)) Then
                strcnt = Len(namae)
                Call Fonta_sei
            Else
                strcnt = LenB(StrConv(namae, vbFromUnicode))
                Call Font_sei
            End If
            Range("Sname").Font.Name = fontname
            If Steam = "セ" Or Steam = "パ" Then
                If Steam = "セ" Then
                    Range("Snum").Value = seban_central
                Else
                    Range("Snum").Value = seban_pacific
                End If
            Else
                Range("Snum").Value = seban
            End If
            Range("Ssein4").Font.Name = "TTEdit2/3ゴシック"
            Range("Sname").Font.Name = fontname
            Range("Snum").Value = seban
            Range("Sname").Value = namae
            Range("Ssein1").Value = "投球数"
            Range("Ssein2").Value = "奪三振"
            Range("Ssein3").Value = "被安打"
            Range("Ssein4").Value = "打席内投球数"
            Range("Ssein5").Value = "与四球"
            Range("Ssein6").Value = "投球回"
            Range("Sseie1,Sseie5,Sseie6").NumberFormatLocal = "G/標準"
            Range("Sseie1").Value = pitch
            Range("Sseie2").Value = k
            Range("Sseie3").Value = hian
            Range("Sseie4").Value = 0
            Range("Sseie5").Value = si
            If pitchinning Mod 3 = 0 Then
                Range("Sseie6").Value = pitchinning / 3
            Else
                Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
        Else
            Range("Ssein4").Font.Name = "TTEdit2/3ゴシック"
            Range("Sname").Font.Name = fontname
            Range("Snum").Value = ""
            Range("Sname").Value = ""
            Range("Ssein1").Value = "投球数"
            Range("Ssein2").Value = "奪三振"
            Range("Ssein3").Value = "被安打"
            Range("Ssein4").Value = "打席内投球数"
            Range("Ssein5").Value = "与四球"
            Range("Ssein6").Value = "投球回"
            Range("Sseie1,Sseie5,Sseie6").NumberFormatLocal = "G/標準"
            Range("Sseie1").Value = 0
            Range("Sseie2").Value = 0
            Range("Sseie3").Value = 0
            Range("Sseie4").Value = 0
            Range("Sseie5").Value = 0
            Range("Sseie6").Value = 0
        End If
    End If
    End With
End Sub
Sub 投手成績リセット()
    
    Dim arr()
    If nowattack = 1 Then
        If Range("text21").Value = "1" Or Range("text21").Value = "P" Then
            seban = CStr(Range("text31").Value)
        ElseIf Range("text22").Value = "1" Or Range("text22").Value = "P" Then
            seban = CStr(Range("text32").Value)
        ElseIf Range("text23").Value = "1" Or Range("text23").Value = "P" Then
            seban = CStr(Range("text33").Value)
        ElseIf Range("text24").Value = "1" Or Range("text24").Value = "P" Then
            seban = CStr(Range("text34").Value)
        ElseIf Range("text25").Value = "1" Or Range("text25").Value = "P" Then
            seban = CStr(Range("text35").Value)
        ElseIf Range("text26").Value = "1" Or Range("text26").Value = "P" Then
            seban = CStr(Range("text36").Value)
        ElseIf Range("text27").Value = "1" Or Range("text27").Value = "P" Then
            seban = CStr(Range("text37").Value)
        ElseIf Range("text28").Value = "1" Or Range("text28").Value = "P" Then
            seban = CStr(Range("text38").Value)
        ElseIf Range("text29").Value = "1" Or Range("text29").Value = "P" Then
            seban = CStr(Range("text39").Value)
        ElseIf Range("text30").Value = "1" Or Range("text30").Value = "P" Then
            seban = CStr(Range("text40").Value)
        End If
        
        arr() = getinfo(Kteam, seban)
        
            If IsNull(Range("gradeitem").MergeArea.Borders.LineStyle) Or Range("gradeitem").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Ssein1", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein2", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein3", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein4", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein5", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ssein6", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
            Else
                returnflg = 1
                Call returnrange("Ssein1,Ssein2,Ssein3,Ssein4,Ssein5,Ssein6", _
                                 grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
            End If
            If IsNull(Range("gradeent").MergeArea.Borders.LineStyle) Or Range("gradeent").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Snum", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sname", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie1", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie2", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie3", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie4", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie5", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Sseie6", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
            Else
                returnflg = 1
                Call returnrange("Snum,Sname,Sseie1,Sseie2,Sseie3,Sseie4,Sseie5,Sseie6", _
                                 grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
            End If
        If Kteam = "セ" Or Kteam = "パ" Then
            seban = arr(0)
            namae = arr(14)
        Else
            namae = arr(13)
        End If
            
        If Len(namae) = LenB(StrConv(namae, vbFromUnicode)) Then
            strcnt = Len(namae)
            Call Fonta_sei
        Else
            strcnt = LenB(StrConv(namae, vbFromUnicode))
            Call Font_sei
        End If
        
        Range("Ksein4").Font.Name = "TTEdit2/3ゴシック"
        Range("Kname").Font.Name = fontname
        Range("Knum").Value = seban
        Range("Kname").Value = namae
        Range("Ksein1").Value = "投球数"
        Range("Ksein2").Value = "奪三振"
        Range("Ksein3").Value = "被安打"
        Range("Ksein4").Value = "打席内投球数"
        Range("Ksein5").Value = "与四球"
        Range("Ksein6").Value = "投球回"
        Range("Kseie1,Kseie5,Kseie6").NumberFormatLocal = "G/標準"
        Range("Kseie1").Value = 0
        Range("Kseie2").Value = 0
        Range("Kseie3").Value = 0
        Range("Kseie4").Value = 0
        Range("Kseie5").Value = 0
        Range("Kseie6").Value = 0
        pitch = 0
        k = 0
        hian = 0
        si = 0
        pitchinning = 0
    Else
        If Range("text1").Value = "1" Or Range("text1").Value = "P" Then
            seban = CStr(Range("text11").Value)
        ElseIf Range("text2").Value = "1" Or Range("text2").Value = "P" Then
            seban = CStr(Range("text12").Value)
        ElseIf Range("text3").Value = "1" Or Range("text3").Value = "P" Then
            seban = CStr(Range("text13").Value)
        ElseIf Range("text4").Value = "1" Or Range("text4").Value = "P" Then
            seban = CStr(Range("text14").Value)
        ElseIf Range("text5").Value = "1" Or Range("text5").Value = "P" Then
            seban = CStr(Range("text15").Value)
        ElseIf Range("text6").Value = "1" Or Range("text6").Value = "P" Then
            seban = CStr(Range("text16").Value)
        ElseIf Range("text7").Value = "1" Or Range("text7").Value = "P" Then
            seban = CStr(Range("text17").Value)
        ElseIf Range("text8").Value = "1" Or Range("text8").Value = "P" Then
            seban = CStr(Range("text18").Value)
        ElseIf Range("text9").Value = "1" Or Range("text9").Value = "P" Then
            seban = CStr(Range("text19").Value)
        ElseIf Range("text10").Value = "1" Or Range("text10").Value = "P" Then
            seban = CStr(Range("text20").Value)
        End If
        
        arr() = getinfo(Steam, seban)
        
            If IsNull(Range("gradeitem").MergeArea.Borders.LineStyle) Or Range("gradeitem").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Ksein1", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ksein2", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ksein3", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ksein4", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ksein5", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
                Call returnrange("Ksein6", grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
            Else
                returnflg = 1
                Call returnrange("Ksein1,Ksein2,Ksein3,Ksein4,Ksein5,Ksein6", _
                                 grade_item_Interior, grade_item_Font, grade_item_Borders, _
                                 grade_item_Color1, grade_item_Color2)
            End If
            If IsNull(Range("gradeent").MergeArea.Borders.LineStyle) Or Range("gradeent").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Knum", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Kname", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Kseie1", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Kseie2", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Kseie3", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Kseie4", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Kseie5", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
                Call returnrange("Kseie6", grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
            Else
                returnflg = 1
                Call returnrange("Knum,Kname,Kseie1,Kseie2,Kseie3,Kseie4,Kseie5,Kseie6", _
                                 grade_entity_Interior, grade_entity_Font, grade_entity_Borders, _
                                 grade_entity_Color1, grade_entity_Color2)
            End If
        If Steam = "セ" Or Steam = "パ" Then
            seban = arr(0)
            namae = arr(14)
        Else
            namae = arr(13)
        End If
            
        If Len(namae) = LenB(StrConv(namae, vbFromUnicode)) Then
            strcnt = Len(namae)
            Call Fonta_sei
        Else
            strcnt = LenB(StrConv(namae, vbFromUnicode))
            Call Font_sei
        End If
        
        Range("Ssein4").Font.Name = "TTEdit2/3ゴシック"
        Range("Sname").Font.Name = fontname
        Range("Snum").Value = seban
        Range("Sname").Value = namae
        Range("Ssein1").Value = "投球数"
        Range("Ssein2").Value = "奪三振"
        Range("Ssein3").Value = "被安打"
        Range("Ssein4").Value = "打席内投球数"
        Range("Ssein5").Value = "与四球"
        Range("Ssein6").Value = "投球回"
        Range("Sseie1,Sseie5,Sseie6").NumberFormatLocal = "G/標準"
        Range("Sseie1").Value = 0
        Range("Sseie2").Value = 0
        Range("Sseie3").Value = 0
        Range("Sseie4").Value = 0
        Range("Sseie5").Value = 0
        Range("Sseie6").Value = 0
        pitch = 0
        k = 0
        hian = 0
        si = 0
        pitchinning = 0
    End If
End Sub
Sub 走塁死()
    souruisi1 = Range("souruisi1").Value
    souruisi2 = Range("souruisi2").Value
    souruisi3 = Range("souruisi3").Value
    If nowattack = 1 Then
        If souruisi1 <> "" Then
            Call returncells(23 + CInt(((souruisi1 - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                             basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
            Call returncells(23 + CInt(((souruisi1 - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                             basic_name_Borders, basic_name_Color1, basic_name_Color2)
            out = out + 1
            pitchinning = pitchinning + 1
            If souruisi2 <> "" Then
                Call returncells(23 + CInt(((souruisi2 - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                                 basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                Call returncells(23 + CInt(((souruisi2 - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                                 basic_name_Borders, basic_name_Color1, basic_name_Color2)
                out = out + 1
                pitchinning = pitchinning + 1
                If souruisi3 <> "" Then
                    Call returncells(23 + CInt(((souruisi3 - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                                     basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                    Call returncells(23 + CInt(((souruisi3 - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                                     basic_name_Borders, basic_name_Color1, basic_name_Color2)
                    out = out + 1
                    pitchinning = pitchinning + 1
                End If
            End If
        End If
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
        If pitchinning Mod 3 = 0 Then
            Range("Kseie6").Value = pitchinning / 3
        Else
            Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    Else
        If souruisi1 <> "" Then
            Call returncells(23 + CInt(((souruisi1 - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                             basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
            Call returncells(23 + CInt(((souruisi1 - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                             basic_name_Borders, basic_name_Color1, basic_name_Color2)
            out = out + 1
            pitchinning = pitchinning + 1
            If souruisi2 <> "" Then
                Call returncells(23 + CInt(((souruisi2 - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                                 basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                Call returncells(23 + CInt(((souruisi2 - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                                 basic_name_Borders, basic_name_Color1, basic_name_Color2)
                out = out + 1
                pitchinning = pitchinning + 1
                If souruisi3 <> "" Then
                    Call returncells(23 + CInt(((souruisi3 - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                                     basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                    Call returncells(23 + CInt(((souruisi3 - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                                     basic_name_Borders, basic_name_Color1, basic_name_Color2)
                    out = out + 1
                    pitchinning = pitchinning + 1
                End If
            End If
        End If
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
        If pitchinning Mod 3 = 0 Then
            Range("Sseie6").Value = pitchinning / 3
        Else
            Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    End If
    Range("souruisi1,souruisi2,souruisi3").Value = ""
End Sub
Sub 得点()
    run = 0
    seikan1 = Range("seikan1").Value
    seikan2 = Range("seikan2").Value
    seikan3 = Range("seikan3").Value
    seikan4 = Range("seikan4").Value
    If nowattack = 1 Then
        If seikan1 <> "" Then
            Call returncells(23 + CInt(((seikan1 - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                             basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
            Call returncells(23 + CInt(((seikan1 - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                             basic_name_Borders, basic_name_Color1, basic_name_Color2)
            run = run + 1
            If seikan2 <> "" Then
                Call returncells(23 + CInt(((seikan2 - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                                 basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                Call returncells(23 + CInt(((seikan2 - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                                 basic_name_Borders, basic_name_Color1, basic_name_Color2)
                run = run + 1
                If seikan3 <> "" Then
                    Call returncells(23 + CInt(((seikan3 - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                                     basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                    Call returncells(23 + CInt(((seikan3 - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                                     basic_name_Borders, basic_name_Color1, basic_name_Color2)
                    run = run + 1
                    If seikan4 <> "" Then
                        Call returncells(23 + CInt(((seikan4 - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                                         basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                        Call returncells(23 + CInt(((seikan4 - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                                         basic_name_Borders, basic_name_Color1, basic_name_Color2)
                        run = run + 1
                    End If
                End If
            End If
            Call 得点代入
            Call 得点表示
        End If
    Else
        If seikan1 <> "" Then
            Call returncells(23 + CInt(((seikan1 - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                             basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
            Call returncells(23 + CInt(((seikan1 - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                             basic_name_Borders, basic_name_Color1, basic_name_Color2)
            run = run + 1
            If seikan2 <> "" Then
                Call returncells(23 + CInt(((seikan2 - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                                 basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                Call returncells(23 + CInt(((seikan2 - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                                 basic_name_Borders, basic_name_Color1, basic_name_Color2)
                run = run + 1
                If seikan3 <> "" Then
                    Call returncells(23 + CInt(((seikan3 - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                                     basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                    Call returncells(23 + CInt(((seikan3 - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                                     basic_name_Borders, basic_name_Color1, basic_name_Color2)
                    run = run + 1
                    If seikan4 <> "" Then
                        Call returncells(23 + CInt(((seikan4 - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                                         basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
                        Call returncells(23 + CInt(((seikan4 - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                                         basic_name_Borders, basic_name_Color1, basic_name_Color2)
                        run = run + 1
                    End If
                End If
            End If
            Call 得点代入
            Call 得点表示
        End If
    End If
    If seikan1 = CStr(nowbatter - 1) Or seikan2 = CStr(nowbatter - 1) Or seikan3 = CStr(nowbatter - 1) Or seikan4 = CStr(nowbatter - 1) Then
        If nowattack = 1 Then
            Call returncells(23 + ((nowbatter - 1) * 11), 13, batter_pos_Interior, batter_pos_Font, _
                     batter_pos_Borders, batter_pos_Color1, batter_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 24, batter_name_Interior, batter_name_Font, _
                             batter_name_Borders, batter_name_Color1, batter_name_Color2)
        Else
            Call returncells(23 + ((nowbatter - 1) * 11), 249, batter_pos_Interior, batter_pos_Font, _
                     batter_pos_Borders, batter_pos_Color1, batter_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 260, batter_name_Interior, batter_name_Font, _
                             batter_name_Borders, batter_name_Color1, batter_name_Color2)
        End If
    End If
    Range("seikan1,seikan2,seikan3,seikan4").Value = ""
End Sub
Sub 得点代入()
    If nowattack = 1 Then
        SrArray(nowinning) = SrArray(nowinning) + run
        intcnt = Len(CStr(SrArray(nowinning)))
        Call IntFont
    Else
        KrArray(nowinning) = KrArray(nowinning) + run
        intcnt = Len(CStr(KrArray(nowinning)))
        Call IntFont
    End If
End Sub
Sub 得点表示()
    Dim a As Integer
    Dim inning As Integer
    inning = nowinning Mod 9
    If nowattack = 1 Then
        If inning = 1 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun1").Font.Name = fontname
            Range("Srun1").Value = CStr(SrArray(nowinning))
        ElseIf inning = 2 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun2").Font.Name = fontname
            Range("Srun2").Value = CStr(SrArray(nowinning))
        ElseIf inning = 3 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun3").Font.Name = fontname
            Range("Srun3").Value = CStr(SrArray(nowinning))
        ElseIf inning = 4 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun4").Font.Name = fontname
            Range("Srun4").Value = CStr(SrArray(nowinning))
        ElseIf inning = 5 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun5").Font.Name = fontname
            Range("Srun5").Value = CStr(SrArray(nowinning))
        ElseIf inning = 6 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun6").Font.Name = fontname
            Range("Srun6").Value = CStr(SrArray(nowinning))
        ElseIf inning = 7 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun7").Font.Name = fontname
            Range("Srun7").Value = CStr(SrArray(nowinning))
        ElseIf inning = 8 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun8").Font.Name = fontname
            Range("Srun8").Value = CStr(SrArray(nowinning))
        ElseIf inning = 0 Then
            intcnt = Len(CStr(SrArray(nowinning)))
            Call IntFont
            Range("Srun9").Font.Name = fontname
            Range("Srun9").Value = CStr(SrArray(nowinning))
        End If
        
        SrArray(0) = 0
        a = 1
        Do While a < UBound(SrArray)
            SrArray(0) = SrArray(0) + SrArray(a)
            a = a + 1
        Loop
        intcnt = Len(CStr(SrArray(0)))
        Call IntFont
        Range("Srunr").Font.Name = fontname
        Range("Srunr").Value = CStr(SrArray(0))
        
    Else
        If nowinning = 1 Or Right(nowinning, 1) = 0 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun1").Font.Name = fontname
            Range("Krun1").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 2 Or Right(nowinning, 1) = 1 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun2").Font.Name = fontname
            Range("Krun2").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 3 Or Right(nowinning, 1) = 2 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun3").Font.Name = fontname
            Range("Krun3").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 4 Or Right(nowinning, 1) = 3 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun4").Font.Name = fontname
            Range("Krun4").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 5 Or Right(nowinning, 1) = 4 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun5").Font.Name = fontname
            Range("Krun5").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 6 Or Right(nowinning, 1) = 5 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun6").Font.Name = fontname
            Range("Krun6").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 7 Or Right(nowinning, 1) = 6 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun7").Font.Name = fontname
            Range("Krun7").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 8 Or Right(nowinning, 1) = 7 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun8").Font.Name = fontname
            Range("Krun8").Value = CStr(KrArray(nowinning))
        ElseIf nowinning = 9 Or Right(nowinning, 1) = 8 Then
            intcnt = Len(CStr(KrArray(nowinning)))
            Call IntFont
            Range("Krun9").Font.Name = fontname
            Range("Krun9").Value = CStr(KrArray(nowinning))
        End If
        
        KrArray(0) = 0
        a = 1
        Do While a < UBound(KrArray)
            KrArray(0) = KrArray(0) + KrArray(a)
            a = a + 1
        Loop
        intcnt = Len(CStr(KrArray(0)))
        Call IntFont
        Range("Krunr").Font.Name = fontname
        Range("Krunr").Value = CStr(KrArray(0))
    End If
End Sub
Sub 出塁()
    
    result = 0
    eumu = Range("eumu").Value
    If nowattack = 1 Then
        seban = CStr(Range("Snum").Value)
    Else
        seban = CStr(Range("Knum").Value)
    End If
    r = nowbatter
    Call 記録代入
    If seikasan = 1 Then
        If kiroku <> "" And InStr(kiroku, "球") < 1 And InStr(kiroku, "遠") < 1 And _
           InStr(kiroku, "妨") < 1 And InStr(kiroku, "犠") < 1 Then
            If nowattack = 1 Then
                Call updateinfo(Steam, seban, 5, 1)
            Else
                Call updateinfo(Kteam, seban, 5, 1)
            End If
        End If
        If InStr(kiroku, "妨") >= 1 Then
            If nowattack = 1 Then
                Kerror = Kerror + 1
                intcnt = Len(CStr(Kerror))
                Call IntFont
                Range("Kerror").Font.Name = fontname
                Range("Kerror").Value = CStr(Kerror)
            Else
                Serror = Serror + 1
                intcnt = Len(CStr(Serror))
                Call IntFont
                Range("Serror").Font.Name = fontname
                Range("Serror").Value = CStr(Serror)
            End If
        End If
        If InStr(kiroku, "安") >= 1 Then
            If nowattack = 1 Then
                Call updateinfo(Steam, seban, 6, 1)
            Else
                Call updateinfo(Kteam, seban, 6, 1)
            End If
        End If
        If InStr(kiroku, "２") >= 1 Then
            If nowattack = 1 Then
                Call updateinfo(Steam, seban, 6, 1)
                Call updateinfo(Steam, seban, 7, 1)
            Else
                Call updateinfo(Kteam, seban, 6, 1)
                Call updateinfo(Kteam, seban, 7, 1)
            End If
        End If
        If InStr(kiroku, "３") >= 1 Then
            If nowattack = 1 Then
                Call updateinfo(Steam, seban, 6, 1)
                Call updateinfo(Steam, seban, 8, 1)
            Else
                Call updateinfo(Kteam, seban, 6, 1)
                Call updateinfo(Kteam, seban, 8, 1)
            End If
        End If
        If InStr(kiroku, "四") >= 1 Then
            If nowattack = 1 Then
                Call updateinfo(Steam, seban, 13, 1)
            Else
                Call updateinfo(Kteam, seban, 13, 1)
            End If
        End If
        If InStr(kiroku, "死") >= 1 Then
            If nowattack = 1 Then
                Call updateinfo(Steam, seban, 14, 1)
            Else
                Call updateinfo(Kteam, seban, 14, 1)
            End If
        End If
        If InStr(kiroku, "犠") >= 1 And InStr(kiroku, "失") < 1 Then
            If nowattack = 1 Then
                Call updateinfo(Steam, seban, 15, 1)
            Else
                Call updateinfo(Kteam, seban, 15, 1)
            End If
        End If
        daten = Range("daten").Value
        If nowattack = 1 Then
            Call updateinfo(Steam, seban, 10, daten)
        Else
            Call updateinfo(Kteam, seban, 10, daten)
        End If
        Range("daten").Value = ""
    End If
    If nowattack = 1 Then
        Call returncells(23 + ((nowbatter - 1) * 11), 13, runner_pos_Interior, runner_pos_Font, _
                         runner_pos_Borders, runner_pos_Color1, runner_pos_Color2)
        Call returncells(23 + ((nowbatter - 1) * 11), 24, runner_name_Interior, runner_name_Font, _
                         runner_name_Borders, runner_name_Color1, runner_name_Color2)
    Else
        Call returncells(23 + ((nowbatter - 1) * 11), 249, runner_pos_Interior, runner_pos_Font, _
                         runner_pos_Borders, runner_pos_Color1, runner_pos_Color2)
        Call returncells(23 + ((nowbatter - 1) * 11), 260, runner_name_Interior, runner_name_Font, _
                         runner_name_Borders, runner_name_Color1, runner_name_Color2)
    End If
    
    If Range("eumu").Value <> "" Then
        If Range("eumu").Value = 1 Then
            Range("kekkae").Value = "E1"
        ElseIf Range("eumu").Value = 2 Then
            Range("kekkae").Value = "E2"
        ElseIf Range("eumu").Value = 3 Then
            Range("kekkae").Value = "E3"
        ElseIf Range("eumu").Value = 4 Then
            Range("kekkae").Value = "E4"
        ElseIf Range("eumu").Value = 5 Then
            Range("kekkae").Value = "E5"
        ElseIf Range("eumu").Value = 6 Then
            Range("kekkae").Value = "E6"
        ElseIf Range("eumu").Value = 7 Then
            Range("kekkae").Value = "E7"
        ElseIf Range("eumu").Value = 8 Then
            Range("kekkae").Value = "E8"
        ElseIf Range("eumu").Value = 9 Then
            Range("kekkae").Value = "E9"
        Else
            Range("kekkae").Value = "E"
        End If
        If nowattack = 1 Then
            Kerror = Kerror + 1
            intcnt = Len(CStr(Kerror))
            Call IntFont
            Range("Kerror").Font.Name = fontname
            Range("Kerror").Value = CStr(Kerror)
        Else
            Serror = Serror + 1
            intcnt = Len(CStr(Serror))
            Call IntFont
            Range("Serror").Font.Name = fontname
            Range("Serror").Value = CStr(Serror)
        End If
    End If

    result = 1
    Call 次打者
    result = 0
    Call 走塁死
    ball = 0
    strike = 0
    If InStr(kiroku, "遠") < 1 Then
        pitch = pitch + 1
    End If
    inpitch = 0
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
    End If
    If InStr(kiroku, "犠") >= 1 Then
        Worksheets("試合情報").Range("A1").Value = 1
    End If
    Call ボール点灯
    Call ストライク点灯
    Call アウト点灯
    Call 得点
    Call 成績表示
    Call 本日結果表示
    Call gameset
    If InStr(kiroku, "安") < 1 And InStr(kiroku, "２") < 1 And InStr(kiroku, "３") < 1 Then
        Call 攻守交代
    End If
    Range("daten").Value = ""
End Sub
Sub ボール点灯()
    If ball = 1 Then
        Range("ball1").Font.Color = RGB(0, 155, 255)
        Range("ball2,ball3").Font.Color = RGB(38, 38, 38)
    ElseIf ball = 2 Then
        Range("ball1,ball2").Font.Color = RGB(0, 155, 255)
        Range("ball3").Font.Color = RGB(38, 38, 38)
    ElseIf ball = 3 Then
        Range("ball1,ball2,ball3").Font.Color = RGB(0, 155, 255)
    Else
        Range("ball1,ball2,ball3").Font.Color = RGB(38, 38, 38)
    End If
End Sub
Sub ボール()
    ball = ball + 1
    If ball = 4 Then
        ball = 3
        Call 出塁
        ball = 0
        strike = 0
        si = si + 1
        If nowattack = 1 Then
            Range("Kseie5").Value = CStr(si)
        Else
            Range("Sseie5").Value = CStr(si)
        End If
        Call 走塁死
    Else
        pitch = pitch + 1
        inpitch = inpitch + 1
    End If
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
    End If
    Call 得点
    Call アウト点灯
    Call ストライク点灯
    Call ボール点灯
    Call gameset
End Sub
Sub ストライク点灯()
    If strike = 1 Then
        Range("strike1").Font.Color = RGB(255, 200, 0)
        Range("strike2").Font.Color = RGB(38, 38, 38)
    ElseIf strike = 2 Then
        Range("strike1,strike2").Font.Color = RGB(255, 200, 0)
    Else
        Range("strike1,strike2").Font.Color = RGB(38, 38, 38)
    End If
End Sub
Sub ストライク()
    strike = strike + 1
    If strike = 3 Then
        If seikasan = 1 Then
            If nowattack = 1 Then
                seban = Range("Snum").Value
                Call updateinfo(Steam, seban, 5, 1)
            Else
                seban = Range("Knum").Value
                Call updateinfo(Kteam, seban, 5, 1)
            End If
        End If
        Call 記録代入
        Call 次打者
        Worksheets("試合情報").Range("F7").Value = 2
        ball = 0
        strike = 0
        k = k + 1
        out = out + 1
        pitchinning = pitchinning + 1
        If nowattack = 1 Then
            Range("Kseie2").Value = CStr(k)
            Range("Kseie4").Value = CStr(inpitch)
            If pitchinning Mod 3 = 0 Then
                Range("Kseie6").Value = pitchinning / 3
            Else
                Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
        Else
            Range("Sseie2").Value = CStr(k)
            Range("Sseie4").Value = CStr(inpitch)
            If pitchinning Mod 3 = 0 Then
                Range("Sseie6").Value = pitchinning / 3
            Else
                Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
        End If
        Call 走塁死
        Call 成績表示
        Call 本日結果表示
    Else
        inpitch = inpitch + 1
    End If
    pitch = pitch + 1
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
    End If
    Call 得点
    Call アウト点灯
    Call ボール点灯
    Call ストライク点灯
    Call gameset
    Call 攻守交代
End Sub
Sub アウト点灯()
    If out = 1 Then
        Range("auto1").Font.Color = RGB(255, 0, 0)
        Range("auto2").Font.Color = RGB(38, 38, 38)
    ElseIf out = 2 Then
        Range("auto1,auto2").Font.Color = RGB(255, 0, 0)
    Else
        Range("auto1,auto2").Font.Color = RGB(38, 38, 38)
    End If
End Sub
Sub 凡退()
    Call 記録代入
    daten = Range("daten").Value
    If seikasan = 1 Then
        If nowattack = 1 Then
            seban = Range("Snum").Value
            Call updateinfo(Steam, seban, 5, 1)
            Call updateinfo(Steam, seban, 10, daten)
        Else
            seban = Range("Knum").Value
            Call updateinfo(Kteam, seban, 5, 1)
            Call updateinfo(Kteam, seban, 10, daten)
        End If
    End If
    Call 次打者
    Call 走塁死
    Call 成績表示
    Call 本日結果表示
    Range("daten").Value = ""
    ball = 0
    strike = 0
    out = out + 1
    pitch = pitch + 1
    pitchinning = pitchinning + 1
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
        If pitchinning Mod 3 = 0 Then
            Range("Kseie6").Value = pitchinning / 3
        Else
            Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
        If pitchinning Mod 3 = 0 Then
            Range("Sseie6").Value = pitchinning / 3
        Else
            Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    End If
    Call ボール点灯
    Call ストライク点灯
    Call アウト点灯
    Call 得点
    Call gameset
    Call 攻守交代
End Sub
Sub 攻守交代()
    If out >= 3 Then
    With Worksheets("試合情報")
        If nowattack = 1 Then
            Call 得点表示
            a = 23
            Range("Sline").Interior.Color = back_Interior
            For a = 23 To 111 Step 11
                sw = 0
                If Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern = 1 Then
                    If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                    And Cells(a, 13).Font.Color = runner_pos_Font _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                    And Cells(a, 24).Font.Color = runner_name_Font _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Slob = Slob + 1
                    End If
                ElseIf Cells(a, 13).Interior.Pattern <> 1 And Cells(a, 24).Interior.Pattern = 1 Then
                    If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                    And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                    And Cells(a, 13).Font.Color = runner_pos_Font _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                    And Cells(a, 24).Font.Color = runner_name_Font _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Slob = Slob + 1
                    End If
                ElseIf Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern <> 1 Then
                    If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                    And Cells(a, 13).Font.Color = runner_pos_Font _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                    And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                    And Cells(a, 24).Font.Color = runner_name_Font _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Slob = Slob + 1
                    End If
                Else
                    If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                    And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                    And Cells(a, 13).Font.Color = runner_pos_Font _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                    And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                    And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                    And Cells(a, 24).Font.Color = runner_name_Font _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Slob = Slob + 1
                    End If
                End If
            Next a
            intcnt = Len(CStr(Slob))
            Call IntFont
            Range("Slob").Font.Name = fontname
            Range("Slob").Value = CStr(Slob)
            If IsNull(Range("basicpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("batterpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("runnerpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("nextpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("dispos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("apppos").MergeArea.Borders.LineStyle) _
            Or Range("basicpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("batterpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("runnerpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("nextpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("dispos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("apppos").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Spos1", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos2", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos3", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos4", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos5", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos6", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos7", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos8", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Spos9", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
            Else
                returnflg = 1
                Call returnrange("Spos1,Spos2,Spos3,Spos4,Spos5,Spos6,Spos7,Spos8,Spos9", _
                                 basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
            End If
            If IsNull(Range("basicname").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("battername").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("runnername").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("nextname").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("disname").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("appname").MergeArea.Borders.LineStyle) _
            Or Range("basicname").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("battername").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("runnername").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("nextname").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("disname").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("appname").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Sname1", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname2", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname3", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname4", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname5", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname6", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname7", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname8", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Sname9", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
            Else
                returnflg = 1
                Call returnrange("Sname1,Sname2,Sname3,Sname4,Sname5,Sname6,Sname7,Sname8,Sname9", _
                                 basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
            End If
            Range("BP117:DO131").Borders.LineStyle = -4142
            Range("Stoday1,Stoday2,Stoday3,Stoday4,Stoday5,Stoday6").Interior.Color = back_Interior
            Range("Stoday1,Stoday2,Stoday3,Stoday4,Stoday5,Stoday6").Value = ""
            .Range("C3").Value = pitchinning
            .Range("C4").Value = si
            .Range("C5").Value = pitch
            .Range("C6").Value = hian
            .Range("C7").Value = k
            .Range("B8").Value = nowbatter
            inpitch = 0
            pitchinning = .Range("B3").Value
            si = .Range("B4").Value
            pitch = .Range("B5").Value
            hian = .Range("B6").Value
            k = .Range("B7").Value
            nowattack = 2
            ReDim Preserve KrArray(nowinning + 1)
            Call RunScore
            Call returncells(23 + ((nowbatter - 1) * 11), 13, next_pos_Interior, next_pos_Font, _
                             next_pos_Borders, next_pos_Color1, next_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 24, next_name_Interior, next_name_Font, _
                             next_name_Borders, next_name_Color1, next_name_Color2)
            nowbatter = .Range("C8")
            Range("Kline").Interior.Color = RGB(255, 0, 0)
            Call returncells(23 + ((nowbatter - 1) * 11), 249, batter_pos_Interior, batter_pos_Font, _
                             batter_pos_Borders, batter_pos_Color1, batter_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 260, batter_name_Interior, batter_name_Font, _
                             batter_name_Borders, batter_name_Color1, batter_name_Color2)
            Call 成績表示
            Call 投手成績表示
            Call 本日結果表示
            ball = 0
            strike = 0
            out = 0
            Call ボール点灯
            Call ストライク点灯
            Call アウト点灯
            ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
            ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
            ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        Else
            Call 得点表示
            a = 23
            For a = 23 To 111 Step 11
                sw = 0
                If Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern = 1 Then
                    If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                    And Cells(a, 249).Font.Color = runner_pos_Font _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                    And Cells(a, 260).Font.Color = runner_name_Font _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Klob = Klob + 1
                    End If
                ElseIf Cells(a, 249).Interior.Pattern <> 1 And Cells(a, 260).Interior.Pattern = 1 Then
                    If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                    And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                    And Cells(a, 249).Font.Color = runner_pos_Font _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                    And Cells(a, 260).Font.Color = runner_name_Font _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Klob = Klob + 1
                    End If
                ElseIf Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern <> 1 Then
                    If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                    And Cells(a, 249).Font.Color = runner_pos_Font _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                    And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                    And Cells(a, 260).Font.Color = runner_name_Font _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Klob = Klob + 1
                    End If
                Else
                    If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                    And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                    And Cells(a, 249).Font.Color = runner_pos_Font _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                    And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                    And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                    And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                    And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                    And Cells(a, 260).Font.Color = runner_name_Font _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                    And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                        Klob = Klob + 1
                    End If
                End If
            Next a
            intcnt = Len(CStr(Klob))
            Call IntFont
            Range("Klob").Font.Name = fontname
            Range("Klob").Value = CStr(Klob)
            Call 延長突入
            nowinning = nowinning + 1
            Range("Kline").Interior.Color = back_Interior
            If IsNull(Range("basicpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("batterpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("runnerpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("nextpos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("dispos").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("apppos").MergeArea.Borders.LineStyle) _
            Or Range("basicpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("batterpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("runnerpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("nextpos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("dispos").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("apppos").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Kpos1", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos2", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos3", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos4", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos5", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos6", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos7", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos8", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
                Call returnrange("Kpos9", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
            Else
                returnflg = 1
                Call returnrange("Kpos1,Kpos2,Kpos3,Kpos4,Kpos5,Kpos6,Kpos7,Kpos8,Kpos9", _
                                 basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                                 basic_pos_Color1, basic_pos_Color2)
            End If
            If IsNull(Range("basicname").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("battername").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("runnername").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("nextname").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("disname").MergeArea.Borders.LineStyle) _
            Or IsNull(Range("appname").MergeArea.Borders.LineStyle) _
            Or Range("basicname").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("battername").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("runnername").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("nextname").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("disname").MergeArea.Borders.LineStyle <> -4142 _
            Or Range("appname").MergeArea.Borders.LineStyle <> -4142 Then
                Call returnrange("Kname1", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname2", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname3", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname4", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname5", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname6", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname7", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname8", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
                Call returnrange("Kname9", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
            Else
                returnflg = 1
                Call returnrange("Kname1,Kname2,Kname3,Kname4,Kname5,Kname6,Kname7,Kname8,Kname9", _
                                 basic_name_Interior, basic_name_Font, basic_name_Borders, _
                                 basic_name_Color1, basic_name_Color2)
            End If
            Range("GJ117:II131").Borders.LineStyle = -4142
            Range("Ktoday1,Ktoday2,Ktoday3,Ktoday4,Ktoday5,Ktoday6").Interior.Color = back_Interior
            Range("Ktoday1,Ktoday2,Ktoday3,Ktoday4,Ktoday5,Ktoday6").Value = ""
            .Range("B3").Value = pitchinning
            .Range("B4").Value = si
            .Range("B5").Value = pitch
            .Range("B6").Value = hian
            .Range("B7").Value = k
            .Range("C8").Value = nowbatter
            inpitch = 0
            pitchinning = .Range("C3").Value
            si = .Range("C4").Value
            pitch = .Range("C5").Value
            hian = .Range("C6").Value
            k = .Range("C7").Value
            nowattack = 1
            ReDim Preserve SrArray(nowinning + 1)
            Call RunScore
            Call returncells(23 + ((nowbatter - 1) * 11), 249, next_pos_Interior, next_pos_Font, _
                             next_pos_Borders, next_pos_Color1, next_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 260, next_name_Interior, next_name_Font, _
                             next_name_Borders, next_name_Color1, next_name_Color2)
            nowbatter = .Range("B8")
            Call returncells(23 + ((nowbatter - 1) * 11), 13, batter_pos_Interior, batter_pos_Font, _
                             batter_pos_Borders, batter_pos_Color1, batter_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 24, batter_name_Interior, batter_name_Font, _
                             batter_name_Borders, batter_name_Color1, batter_name_Color2)
            Range("Sline").Interior.Color = RGB(255, 0, 0)
            Call 成績表示
            Call 投手成績表示
            Call 本日結果表示
            ball = 0
            strike = 0
            out = 0
            Call ボール点灯
            Call ストライク点灯
            Call アウト点灯
            ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
            ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
            ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        End If
    End With
    End If
End Sub
Sub 延長突入()
    If jogen >= 10 And (nowinning Mod 9) = 0 And nowattack = 2 And out = 3 And SrArray(0) = KrArray(0) Then
        Range("Srun1,Srun2,Srun3,Srun4,Srun5,Srun6,Srun7,Srun8,Srun9").Value = ""
        Range("Krun1,Krun2,Krun3,Krun4,Krun5,Krun6,Krun7,Krun8,Krun9").Value = ""
        Range("inin1,inin2,inin3,inin4,inin5,inin6,inin7,inin8,inin9").Value = ""
        Call returnrange("Krun9", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        If nowinning + 1 <= jogen Then
            Range("inin1").Value = nowinning + 1
        End If
        If nowinning + 2 <= jogen Then
            Range("inin2").Value = nowinning + 2
        End If
        If nowinning + 3 <= jogen Then
            Range("inin3").Value = nowinning + 3
        End If
        If nowinning + 4 <= jogen Then
            Range("inin4").Value = nowinning + 4
        End If
        If nowinning + 5 <= jogen Then
            Range("inin5").Value = nowinning + 5
        End If
        If nowinning + 6 <= jogen Then
            Range("inin6").Value = nowinning + 6
        End If
        If nowinning + 7 <= jogen Then
            Range("inin7").Value = nowinning + 7
        End If
        If nowinning + 8 <= jogen Then
            Range("inin8").Value = nowinning + 8
        End If
        If nowinning + 9 <= jogen Then
            Range("inin9").Value = nowinning + 9
        End If
        Range("Srun1").Interior.Color = RGB(255, 0, 0)
    End If
End Sub
Sub gameset()
    If nowinning = 9 And nowattack = 1 And SrArray(0) < KrArray(0) And out = 3 Then
        Call 得点表示
        Range("Krun9") = "X"
        Range("Sline").Interior.Color = back_Interior
        a = 23
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern = 1 Then
                If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Slob = Slob + 1
                End If
            ElseIf Cells(a, 13).Interior.Pattern <> 1 And Cells(a, 24).Interior.Pattern = 1 Then
                If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Slob = Slob + 1
                End If
            ElseIf Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern <> 1 Then
                If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Slob = Slob + 1
                End If
            Else
                If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 13).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 24).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Slob = Slob + 1
                End If
            End If
        Next a
        intcnt = Len(CStr(Slob))
        Call IntFont
        Range("Slob").Font.Name = fontname
        Range("Slob").Value = CStr(Slob)
        returnflg = 1
        Call MemberCellReset
        ball = 0
        strike = 0
        out = 0
        Call ボール点灯
        Call ストライク点灯
        Call アウト点灯
        Call returnrange("Srun9", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Range("BP46:DO108,GJ46:II108,BP117:DO131,GJ117:II131").Borders.LineStyle = -4142
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Interior.Color = back_Interior
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Value = ""
        Range("CW28:HY28").Font.Color = RGB(255, 255, 0)
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        gameflg = 0
    ElseIf nowinning >= 9 And nowattack = 2 And SrArray(0) > KrArray(0) And out = 3 Then
        Call 得点表示
        a = 23
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            ElseIf Cells(a, 249).Interior.Pattern <> 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            ElseIf Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern <> 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            Else
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            End If
        Next a
        intcnt = Len(CStr(Klob))
        Call IntFont
        Range("Klob").Font.Name = fontname
        Range("Klob").Value = CStr(Klob)
        returnflg = 1
        Call MemberCellReset
        ball = 0
        strike = 0
        out = 0
        Call ボール点灯
        Call ストライク点灯
        Call アウト点灯
        Call ScoreCellReset
        Range("Kline").Interior.Color = back_Interior
        Range("BP46:DO108,GJ46:II108,BP117:DO131,GJ117:II131").Borders.LineStyle = -4142
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Interior.Color = back_Interior
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Value = ""
        Range("CW17:HY17").Font.Color = RGB(255, 255, 0)
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        gameflg = 0
    ElseIf nowinning >= 9 And nowattack = 2 And SrArray(0) < KrArray(0) Then
        Dim inning As Integer
        inning = nowinning Mod 9
        Select Case inning
        Case 0
            Range("Krun9").Font.Name = "Arial Narrow"
            Range("Krun9").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun9", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 1
            Range("Krun1").Font.Name = "Arial Narrow"
            Range("Krun1").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun1", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 2
            Range("Krun2").Font.Name = "Arial Narrow"
            Range("Krun2").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun2", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 3
            Range("Krun3").Font.Name = "Arial Narrow"
            Range("Krun3").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun3", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 4
            Range("Krun4").Font.Name = "Arial Narrow"
            Range("Krun4").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun4", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 5
            Range("Krun5").Font.Name = "Arial Narrow"
            Range("Krun5").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun5", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 6
            Range("Krun6").Font.Name = "Arial Narrow"
            Range("Krun6").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun6", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 7
            Range("Krun7").Font.Name = "Arial Narrow"
            Range("Krun7").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun7", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        Case 8
            Range("Krun8").Font.Name = "Arial Narrow"
            Range("Krun8").Value = KrArray(nowinning) & "x"
            Call returnrange("Krun8", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
        End Select
        Range("Kline").Interior.Color = back_Interior
        a = 23
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            ElseIf Cells(a, 249).Interior.Pattern <> 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            ElseIf Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern <> 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            Else
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            End If
        Next a
        intcnt = Len(CStr(Klob))
        Call IntFont
        Range("Klob").Font.Name = fontname
        Range("Klob").Value = CStr(Klob)
        returnflg = 1
        Call MemberCellReset
        ball = 0
        strike = 0
        out = 0
        Call ボール点灯
        Call ストライク点灯
        Call アウト点灯
        Range("BP46:DO108,GJ46:II108,BP117:DO131,GJ117:II131").Borders.LineStyle = -4142
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Interior.Color = back_Interior
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Value = ""
        Range("CW28:HY28").Font.Color = RGB(255, 255, 0)
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        gameflg = 0
    End If
    If jogen = nowinning And nowattack = 2 And SrArray(0) = KrArray(0) And out = 3 Then
        Call 得点表示
        Range("Kline").Interior.Color = back_Interior
        a = 23
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            ElseIf Cells(a, 249).Interior.Pattern <> 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            ElseIf Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern <> 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            Else
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).MergeArea.Borders(xlEdgeTop).Color = runner_pos_Borders(xlEdgeTop).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeBottom).Color = runner_pos_Borders(xlEdgeBottom).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeLeft).Color = runner_pos_Borders(xlEdgeLeft).Color _
                And Cells(a, 249).MergeArea.Borders(xlEdgeRight).Color = runner_pos_Borders(xlEdgeRight).Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).MergeArea.Borders(xlEdgeTop).Color = runner_name_Borders(xlEdgeTop).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeBottom).Color = runner_name_Borders(xlEdgeBottom).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeLeft).Color = runner_name_Borders(xlEdgeLeft).Color _
                And Cells(a, 260).MergeArea.Borders(xlEdgeRight).Color = runner_name_Borders(xlEdgeRight).Color Then
                    Klob = Klob + 1
                End If
            End If
        Next a
        intcnt = Len(CStr(Klob))
        Call IntFont
        Range("Klob").Font.Name = fontname
        Range("Klob").Value = CStr(Klob)
        returnflg = 1
        Call MemberCellReset
        Call ScoreCellReset
        ball = 0
        strike = 0
        out = 0
        Call ボール点灯
        Call ストライク点灯
        Call アウト点灯
        Range("BP46:DO108,GJ46:II108,BP117:DO131,GJ117:II131").Borders.LineStyle = -4142
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Interior.Color = back_Interior
        Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Value = ""
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        gameflg = 0
    End If
End Sub
Sub 記録代入()
    kiroku = Range("kiroku").Value
    If Range("daten").Value <> "" Then
        Select Case Range("daten").Value
            Case 1
                kiroku = kiroku + "①"
            Case 2
                kiroku = kiroku + "②"
            Case 3
                kiroku = kiroku + "③"
            Case 4
                kiroku = kiroku + "④"
        End Select
    End If
    With Worksheets("試合情報")
        If nowattack = 1 Then
            Maxrow = .Cells(Rows.count, nowbatter).End(xlUp).Row
            .Cells((Maxrow + 1), nowbatter).Value = kiroku
        Else
            Maxrow = .Cells(Rows.count, (nowbatter + 9)).End(xlUp).Row
            .Cells((Maxrow + 1), (nowbatter + 9)).Value = kiroku
        End If
    End With
    Range("kiroku") = ""
End Sub
Sub 本日結果表示()
    With Worksheets("試合情報")
        If nowattack = 1 Then
            Maxrow = .Cells(Rows.count, nowbatter).End(xlUp).Row
            Range("BP117:CZ125").Interior.Color = back_Interior
            Range("BP117:DO131").Borders.LineStyle = -4142
            Range("Stoday1,Stoday2,Stoday3,Stoday4,Stoday5,Stoday6").Value = ""
            Stoday1 = ""
            Stoday2 = ""
            Stoday3 = ""
            Stoday4 = ""
            Stoday5 = ""
            Stoday6 = ""
            If Maxrow >= 17 Then
                Stoday6 = .Cells(Maxrow, nowbatter).Value
                Stoday5 = .Cells(Maxrow - 1, nowbatter).Value
                Stoday4 = .Cells(Maxrow - 2, nowbatter).Value
                Stoday3 = .Cells(Maxrow - 3, nowbatter).Value
                Stoday2 = .Cells(Maxrow - 4, nowbatter).Value
                Stoday1 = .Cells(Maxrow - 5, nowbatter).Value
            Else
                Stoday6 = .Cells(16, nowbatter).Value
                Stoday5 = .Cells(15, nowbatter).Value
                Stoday4 = .Cells(14, nowbatter).Value
                Stoday3 = .Cells(13, nowbatter).Value
                Stoday2 = .Cells(12, nowbatter).Value
                Stoday1 = .Cells(11, nowbatter).Value
            End If
            If Stoday1 <> "" Then
                If InStr(Stoday1, "安") >= 1 Or InStr(Stoday1, "２") >= 1 Or InStr(Stoday1, "３") >= 1 _
                Or InStr(Stoday1, "本") >= 1 Then
                    Call returnrange("Stoday1", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                     result_hit_Color1, result_hit_Color2)
                ElseIf InStr(Stoday1, "犠") >= 1 Or InStr(Stoday1, "球") >= 1 Or InStr(Stoday1, "ス") >= 1 Then
                    Call returnrange("Stoday1", result_other_Interior, result_other_Font, result_other_Borders, _
                                     result_other_Color1, result_other_Color2)
                Else
                    Call returnrange("Stoday1", result_out_Interior, result_out_Font, result_out_Borders, _
                                     result_out_Color1, result_out_Color2)
                End If
                If Len(Stoday1) = 2 Then
                    Range("Stoday1").Font.Name = "IPAexゴシック"
                Else
                    Range("Stoday1").Font.Name = "TTEdit2/3ゴシック"
                End If
                Range("Stoday1").Value = Stoday1
                If Stoday2 <> "" Then
                    If InStr(Stoday2, "安") >= 1 Or InStr(Stoday2, "２") >= 1 Or InStr(Stoday2, "３") >= 1 _
                    Or InStr(Stoday2, "本") >= 1 Then
                        Call returnrange("Stoday2", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                         result_hit_Color1, result_hit_Color2)
                    ElseIf InStr(Stoday2, "犠") >= 1 Or InStr(Stoday2, "球") >= 1 Or InStr(Stoday2, "ス") >= 1 Then
                        Call returnrange("Stoday2", result_other_Interior, result_other_Font, result_other_Borders, _
                                         result_other_Color1, result_other_Color2)
                    Else
                        Call returnrange("Stoday2", result_out_Interior, result_out_Font, result_out_Borders, _
                                         result_out_Color1, result_out_Color2)
                    End If
                    If Len(Stoday2) = 2 Then
                        Range("Stoday2").Font.Name = "IPAexゴシック"
                    Else
                        Range("Stoday2").Font.Name = "TTEdit2/3ゴシック"
                    End If
                    Range("Stoday2").Value = Stoday2
                    If Stoday3 <> "" Then
                        If InStr(Stoday3, "安") >= 1 Or InStr(Stoday3, "２") >= 1 Or InStr(Stoday3, "３") >= 1 _
                        Or InStr(Stoday3, "本") >= 1 Then
                            Call returnrange("Stoday3", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                             result_hit_Color1, result_hit_Color2)
                        ElseIf InStr(Stoday3, "犠") >= 1 Or InStr(Stoday3, "球") >= 1 Or InStr(Stoday3, "ス") >= 1 Then
                            Call returnrange("Stoday3", result_other_Interior, result_other_Font, result_other_Borders, _
                                             result_other_Color1, result_other_Color2)
                        Else
                            Call returnrange("Stoday3", result_out_Interior, result_out_Font, result_out_Borders, _
                                             result_out_Color1, result_out_Color2)
                        End If
                        If Len(Stoday3) = 2 Then
                            Range("Stoday3").Font.Name = "IPAexゴシック"
                        Else
                            Range("Stoday3").Font.Name = "TTEdit2/3ゴシック"
                        End If
                        Range("Stoday3").Value = Stoday3
                        If Stoday4 <> "" Then
                            If InStr(Stoday4, "安") >= 1 Or InStr(Stoday4, "２") >= 1 Or InStr(Stoday4, "３") >= 1 _
                            Or InStr(Stoday4, "本") >= 1 Then
                                Call returnrange("Stoday4", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                                 result_hit_Color1, result_hit_Color2)
                            ElseIf InStr(Stoday4, "犠") >= 1 Or InStr(Stoday4, "球") >= 1 Or InStr(Stoday4, "ス") >= 1 Then
                                Call returnrange("Stoday4", result_other_Interior, result_other_Font, result_other_Borders, _
                                                 result_other_Color1, result_other_Color2)
                            Else
                                Call returnrange("Stoday4", result_out_Interior, result_out_Font, result_out_Borders, _
                                                 result_out_Color1, result_out_Color2)
                            End If
                            If Len(Stoday4) = 2 Then
                                Range("Stoday4").Font.Name = "IPAexゴシック"
                            Else
                                Range("Stoday4").Font.Name = "TTEdit2/3ゴシック"
                            End If
                            Range("Stoday4").Value = Stoday4
                            If Stoday5 <> "" Then
                                If InStr(Stoday5, "安") >= 1 Or InStr(Stoday5, "２") >= 1 Or InStr(Stoday5, "３") >= 1 _
                                Or InStr(Stoday5, "本") >= 1 Then
                                    Call returnrange("Stoday5", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                                     result_hit_Color1, result_hit_Color2)
                                ElseIf InStr(Stoday5, "犠") >= 1 Or InStr(Stoday5, "球") >= 1 Or InStr(Stoday5, "ス") >= 1 Then
                                    Call returnrange("Stoday5", result_other_Interior, result_other_Font, result_other_Borders, _
                                                     result_other_Color1, result_other_Color2)
                                Else
                                    Call returnrange("Stoday5", result_out_Interior, result_out_Font, result_out_Borders, _
                                                     result_out_Color1, result_out_Color2)
                                End If
                                If Len(Stoday5) = 2 Then
                                    Range("Stoday5").Font.Name = "IPAexゴシック"
                                Else
                                    Range("Stoday5").Font.Name = "TTEdit2/3ゴシック"
                                End If
                                Range("Stoday5").Value = Stoday5
                                If Stoday6 <> "" Then
                                    If InStr(Stoday6, "安") >= 1 Or InStr(Stoday6, "２") >= 1 Or InStr(Stoday6, "３") >= 1 _
                                    Or InStr(Stoday6, "本") >= 1 Then
                                        Call returnrange("Stoday6", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                                         result_hit_Color1, result_hit_Color2)
                                    ElseIf InStr(Stoday6, "犠") >= 1 Or InStr(Stoday6, "球") >= 1 Or InStr(Stoday6, "ス") >= 1 Then
                                        Call returnrange("Stoday6", result_other_Interior, result_other_Font, result_other_Borders, _
                                                         result_other_Color1, result_other_Color2)
                                    Else
                                        Call returnrange("Stoday6", result_out_Interior, result_out_Font, result_out_Borders, _
                                                         result_out_Color1, result_out_Color2)
                                    End If
                                    If Len(Stoday6) = 2 Then
                                        Range("Stoday6").Font.Name = "IPAexゴシック"
                                    Else
                                        Range("Stoday6").Font.Name = "TTEdit2/3ゴシック"
                                    End If
                                    Range("Stoday6").Value = Stoday6
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            Maxrow = .Cells(Rows.count, (nowbatter + 9)).End(xlUp).Row
            Range("GJ117:HT125").Interior.Color = back_Interior
            Range("GJ117:II131").Borders.LineStyle = -4142
            Range("Ktoday1,Ktoday2,Ktoday3,Ktoday4,Ktoday5,Ktoday6").Value = ""
            Ktoday1 = ""
            Ktoday2 = ""
            Ktoday3 = ""
            Ktoday4 = ""
            Ktoday5 = ""
            Ktoday6 = ""
            If Maxrow >= 17 Then
                Ktoday6 = .Cells(Maxrow, nowbatter + 9).Value
                Ktoday5 = .Cells(Maxrow - 1, nowbatter + 9).Value
                Ktoday4 = .Cells(Maxrow - 2, nowbatter + 9).Value
                Ktoday3 = .Cells(Maxrow - 3, nowbatter + 9).Value
                Ktoday2 = .Cells(Maxrow - 4, nowbatter + 9).Value
                Ktoday1 = .Cells(Maxrow - 5, nowbatter + 9).Value
            Else
                Ktoday6 = .Cells(16, nowbatter + 9).Value
                Ktoday5 = .Cells(15, nowbatter + 9).Value
                Ktoday4 = .Cells(14, nowbatter + 9).Value
                Ktoday3 = .Cells(13, nowbatter + 9).Value
                Ktoday2 = .Cells(12, nowbatter + 9).Value
                Ktoday1 = .Cells(11, nowbatter + 9).Value
            End If
            If Ktoday1 <> "" Then
                If InStr(Ktoday1, "安") >= 1 Or InStr(Ktoday1, "２") >= 1 Or InStr(Ktoday1, "３") >= 1 _
                Or InStr(Ktoday1, "本") >= 1 Then
                    Call returnrange("Ktoday1", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                     result_hit_Color1, result_hit_Color2)
                ElseIf InStr(Ktoday1, "犠") >= 1 Or InStr(Ktoday1, "球") >= 1 Or InStr(Ktoday1, "ス") >= 1 Then
                    Call returnrange("Ktoday1", result_other_Interior, result_other_Font, result_other_Borders, _
                                     result_other_Color1, result_other_Color2)
                Else
                    Call returnrange("Ktoday1", result_out_Interior, result_out_Font, result_out_Borders, _
                                     result_out_Color1, result_out_Color2)
                End If
                If Len(Ktoday1) = 2 Then
                    Range("Ktoday1").Font.Name = "IPAexゴシック"
                Else
                    Range("Ktoday1").Font.Name = "TTEdit2/3ゴシック"
                End If
                Range("Ktoday1").Value = Ktoday1
                If Ktoday2 <> "" Then
                    If InStr(Ktoday2, "安") >= 1 Or InStr(Ktoday2, "２") >= 1 Or InStr(Ktoday2, "３") >= 1 _
                    Or InStr(Ktoday2, "本") >= 1 Then
                        Call returnrange("Ktoday2", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                         result_hit_Color1, result_hit_Color2)
                    ElseIf InStr(Ktoday2, "犠") >= 1 Or InStr(Ktoday2, "球") >= 1 Or InStr(Ktoday2, "ス") >= 1 Then
                        Call returnrange("Ktoday2", result_other_Interior, result_other_Font, result_other_Borders, _
                                         result_other_Color1, result_other_Color2)
                    Else
                        Call returnrange("Ktoday2", result_out_Interior, result_out_Font, result_out_Borders, _
                                         result_out_Color1, result_out_Color2)
                    End If
                    If Len(Ktoday2) = 2 Then
                        Range("Ktoday2").Font.Name = "IPAexゴシック"
                    Else
                        Range("Ktoday2").Font.Name = "TTEdit2/3ゴシック"
                    End If
                    Range("Ktoday2").Value = Ktoday2
                    If Ktoday3 <> "" Then
                        If InStr(Ktoday3, "安") >= 1 Or InStr(Ktoday3, "２") >= 1 Or InStr(Ktoday3, "３") >= 1 _
                        Or InStr(Ktoday3, "本") >= 1 Then
                            Call returnrange("Ktoday3", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                             result_hit_Color1, result_hit_Color2)
                        ElseIf InStr(Ktoday3, "犠") >= 1 Or InStr(Ktoday3, "球") >= 1 Or InStr(Ktoday3, "ス") >= 1 Then
                            Call returnrange("Ktoday3", result_other_Interior, result_other_Font, result_other_Borders, _
                                             result_other_Color1, result_other_Color2)
                        Else
                            Call returnrange("Ktoday3", result_out_Interior, result_out_Font, result_out_Borders, _
                                             result_out_Color1, result_out_Color2)
                        End If
                        If Len(Ktoday3) = 2 Then
                            Range("Ktoday3").Font.Name = "IPAexゴシック"
                        Else
                            Range("Ktoday3").Font.Name = "TTEdit2/3ゴシック"
                        End If
                        Range("Ktoday3").Value = Ktoday3
                        If Ktoday4 <> "" Then
                            If InStr(Ktoday4, "安") >= 1 Or InStr(Ktoday4, "２") >= 1 Or InStr(Ktoday4, "３") >= 1 _
                            Or InStr(Ktoday4, "本") >= 1 Then
                                Call returnrange("Ktoday4", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                                 result_hit_Color1, result_hit_Color2)
                            ElseIf InStr(Ktoday4, "犠") >= 1 Or InStr(Ktoday4, "球") >= 1 Or InStr(Ktoday4, "ス") >= 1 Then
                                Call returnrange("Ktoday4", result_other_Interior, result_other_Font, result_other_Borders, _
                                                 result_other_Color1, result_other_Color2)
                            Else
                                Call returnrange("Ktoday4", result_out_Interior, result_out_Font, result_out_Borders, _
                                                 result_out_Color1, result_out_Color2)
                            End If
                            If Len(Ktoday4) = 2 Then
                                Range("Ktoday4").Font.Name = "IPAexゴシック"
                            Else
                                Range("Ktoday4").Font.Name = "TTEdit2/3ゴシック"
                            End If
                            Range("Ktoday4").Value = Ktoday4
                            If Ktoday5 <> "" Then
                                If InStr(Ktoday5, "安") >= 1 Or InStr(Ktoday5, "２") >= 1 Or InStr(Ktoday5, "３") >= 1 _
                                Or InStr(Ktoday5, "本") >= 1 Then
                                    Call returnrange("Ktoday5", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                                     result_hit_Color1, result_hit_Color2)
                                ElseIf InStr(Ktoday5, "犠") >= 1 Or InStr(Ktoday5, "球") >= 1 Or InStr(Ktoday5, "ス") >= 1 Then
                                    Call returnrange("Ktoday5", result_other_Interior, result_other_Font, result_other_Borders, _
                                                     result_other_Color1, result_other_Color2)
                                Else
                                    Call returnrange("Ktoday5", result_out_Interior, result_out_Font, result_out_Borders, _
                                                     result_out_Color1, result_out_Color2)
                                End If
                                If Len(Ktoday5) = 2 Then
                                    Range("Ktoday5").Font.Name = "IPAexゴシック"
                                Else
                                    Range("Ktoday5").Font.Name = "TTEdit2/3ゴシック"
                                End If
                                Range("Ktoday5").Value = Ktoday5
                                If Ktoday6 <> "" Then
                                    If InStr(Ktoday6, "安") >= 1 Or InStr(Ktoday6, "２") >= 1 Or InStr(Ktoday6, "３") >= 1 _
                                    Or InStr(Ktoday6, "本") >= 1 Then
                                        Call returnrange("Ktoday6", result_hit_Interior, result_hit_Font, result_hit_Borders, _
                                                         result_hit_Color1, result_hit_Color2)
                                    ElseIf InStr(Ktoday6, "犠") >= 1 Or InStr(Ktoday6, "球") >= 1 Or InStr(Ktoday6, "ス") >= 1 Then
                                        Call returnrange("Ktoday6", result_other_Interior, result_other_Font, result_other_Borders, _
                                                         result_other_Color1, result_other_Color2)
                                    Else
                                        Call returnrange("Ktoday6", result_out_Interior, result_out_Font, result_out_Borders, _
                                                         result_out_Color1, result_out_Color2)
                                    End If
                                    If Len(Ktoday6) = 2 Then
                                        Range("Ktoday6").Font.Name = "IPAexゴシック"
                                    Else
                                        Range("Ktoday6").Font.Name = "TTEdit2/3ゴシック"
                                    End If
                                    Range("Ktoday6").Value = Ktoday6
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Sub 単打()
    Range("kekkah").Value = "H"
    eumu = Range("eumu").Value
    If nowattack = 1 Then
        Shit = Shit + 1
        intcnt = Len(CStr(Shit))
        Call IntFont
        Range("Shit").Font.Name = fontname
        Range("Shit").Value = CStr(Shit)
    Else
        Khit = Khit + 1
        intcnt = Len(CStr(Khit))
        Call IntFont
        Range("Khit").Font.Name = fontname
        Range("Khit").Value = CStr(Khit)
    End If
    If Range("eumu").Value <> "" Then
        If Range("eumu").Value = 1 Then
            Range("kekkae").Value = "E1"
        ElseIf Range("eumu").Value = 2 Then
            Range("kekkae").Value = "E2"
        ElseIf Range("eumu").Value = 3 Then
            Range("kekkae").Value = "E3"
        ElseIf Range("eumu").Value = 4 Then
            Range("kekkae").Value = "E4"
        ElseIf Range("eumu").Value = 5 Then
            Range("kekkae").Value = "E5"
        ElseIf Range("eumu").Value = 6 Then
            Range("kekkae").Value = "E6"
        ElseIf Range("eumu").Value = 7 Then
            Range("kekkae").Value = "E7"
        ElseIf Range("eumu").Value = 8 Then
            Range("kekkae").Value = "E8"
        ElseIf Range("eumu").Value = 9 Then
            Range("kekkae").Value = "E9"
        Else
            Range("kekkae").Value = "E"
        End If
        If nowattack = 1 Then
            Kerror = Kerror + 1
            intcnt = Len(CStr(Kerror))
            Call IntFont
            Range("Kerror").Font.Name = fontname
            Range("Kerror").Value = CStr(Kerror)
        Else
            Serror = Serror + 1
            intcnt = Len(CStr(Serror))
            Call IntFont
            Range("Serror").Font.Name = fontname
            Range("Serror").Value = CStr(Serror)
        End If
    End If
    Application.Wait [now()+"0:00:05"]
    Call 出塁
    hian = hian + 1
    If gameflg = 1 Then
        If nowattack = 1 Then
            Range("Kseie3").Value = CStr(hian)
        Else
            Range("Sseie3").Value = CStr(hian)
        End If
    End If
    Range("kekkah").Value = ""
    Range("kekkae").Value = ""
    Range("eumu").Value = ""
    Call 攻守交代
End Sub
Sub 二塁打()
    Range("kekkah").Value = "2H"
    eumu = Range("eumu").Value
    If nowattack = 1 Then
        Shit = Shit + 1
        intcnt = Len(CStr(Shit))
        Call IntFont
        Range("Shit").Font.Name = fontname
        Range("Shit").Value = CStr(Shit)
    Else
        Khit = Khit + 1
        intcnt = Len(CStr(Khit))
        Call IntFont
        Range("Khit").Font.Name = fontname
        Range("Khit").Value = CStr(Khit)
    End If
    If Range("eumu").Value <> "" Then
        If Range("eumu").Value = 1 Then
            Range("kekkae").Value = "E1"
        ElseIf Range("eumu").Value = 2 Then
            Range("kekkae").Value = "E2"
        ElseIf Range("eumu").Value = 3 Then
            Range("kekkae").Value = "E3"
        ElseIf Range("eumu").Value = 4 Then
            Range("kekkae").Value = "E4"
        ElseIf Range("eumu").Value = 5 Then
            Range("kekkae").Value = "E5"
        ElseIf Range("eumu").Value = 6 Then
            Range("kekkae").Value = "E6"
        ElseIf Range("eumu").Value = 7 Then
            Range("kekkae").Value = "E7"
        ElseIf Range("eumu").Value = 8 Then
            Range("kekkae").Value = "E8"
        ElseIf Range("eumu").Value = 9 Then
            Range("kekkae").Value = "E9"
        Else
            Range("kekkae").Value = "E"
        End If
        If nowattack = 1 Then
            Kerror = Kerror + 1
            intcnt = Len(CStr(Kerror))
            Call IntFont
            Range("Kerror").Font.Name = fontname
            Range("Kerror").Value = CStr(Kerror)
        Else
            Serror = Serror + 1
            intcnt = Len(CStr(Serror))
            Call IntFont
            Range("Serror").Font.Name = fontname
            Range("Serror").Value = CStr(Serror)
        End If
    End If
    Application.Wait [now()+"0:00:05"]
    Call 出塁
    hian = hian + 1
    If gameflg = 1 Then
        If nowattack = 1 Then
            Range("Kseie3").Value = CStr(hian)
        Else
            Range("Sseie3").Value = CStr(hian)
        End If
    End If
    Range("kekkah").Value = ""
    Range("kekkae").Value = ""
    Range("eumu").Value = ""
    Call 攻守交代
End Sub
Sub 三塁打()
    Range("kekkah").Value = "3H"
    eumu = Range("eumu").Value
    If nowattack = 1 Then
        Shit = Shit + 1
        intcnt = Len(CStr(Shit))
        Call IntFont
        Range("Shit").Font.Name = fontname
        Range("Shit").Value = CStr(Shit)
    Else
        Khit = Khit + 1
        intcnt = Len(CStr(Khit))
        Call IntFont
        Range("Khit").Font.Name = fontname
        Range("Khit").Value = CStr(Khit)
    End If
    If Range("eumu").Value <> "" Then
        If Range("eumu").Value = 1 Then
            Range("kekkae").Value = "E1"
        ElseIf Range("eumu").Value = 2 Then
            Range("kekkae").Value = "E2"
        ElseIf Range("eumu").Value = 3 Then
            Range("kekkae").Value = "E3"
        ElseIf Range("eumu").Value = 4 Then
            Range("kekkae").Value = "E4"
        ElseIf Range("eumu").Value = 5 Then
            Range("kekkae").Value = "E5"
        ElseIf Range("eumu").Value = 6 Then
            Range("kekkae").Value = "E6"
        ElseIf Range("eumu").Value = 7 Then
            Range("kekkae").Value = "E7"
        ElseIf Range("eumu").Value = 8 Then
            Range("kekkae").Value = "E8"
        ElseIf Range("eumu").Value = 9 Then
            Range("kekkae").Value = "E9"
        Else
            Range("kekkae").Value = "E"
        End If
        If nowattack = 1 Then
            Kerror = Kerror + 1
            intcnt = Len(CStr(Kerror))
            Call IntFont
            Range("Kerror").Font.Name = fontname
            Range("Kerror").Value = CStr(Kerror)
        Else
            Serror = Serror + 1
            intcnt = Len(CStr(Serror))
            Call IntFont
            Range("Serror").Font.Name = fontname
            Range("Serror").Value = CStr(Serror)
        End If
    End If
    Application.Wait [now()+"0:00:05"]
    Call 出塁
    hian = hian + 1
    If gameflg = 1 Then
        If nowattack = 1 Then
            Range("Kseie3").Value = CStr(hian)
        Else
            Range("Sseie3").Value = CStr(hian)
        End If
    End If
    Range("kekkah").Value = ""
    Range("kekkae").Value = ""
    Range("eumu").Value = ""
    Call 攻守交代
End Sub
Sub 本塁打()
    Range("kekkah").Value = "HR"
    Call 記録代入
    If nowattack = 1 Then
        Shit = Shit + 1
        intcnt = Len(CStr(Shit))
        Call IntFont
        Range("Shit").Font.Name = fontname
        Range("Shit").Value = CStr(Shit)
    Else
        Khit = Khit + 1
        intcnt = Len(CStr(Khit))
        Call IntFont
        Range("Khit").Font.Name = fontname
        Range("Khit").Value = CStr(Khit)
    End If
    Application.Wait [now()+"0:00:05"]
    daten = Range("daten").Value
    If seikasan = 1 Then
        If nowattack = 1 Then
            seban = Range("Snum").Value
            Call updateinfo(Steam, seban, 5, 1)
            Call updateinfo(Steam, seban, 6, 1)
            Call updateinfo(Steam, seban, 9, 1)
            Call updateinfo(Steam, seban, 10, daten)
        Else
            seban = Range("Knum").Value
            Call updateinfo(Kteam, seban, 5, 1)
            Call updateinfo(Kteam, seban, 6, 1)
            Call updateinfo(Kteam, seban, 9, 1)
            Call updateinfo(Kteam, seban, 10, daten)
        End If
    End If
    Call 次打者
    Range("daten").Value = ""
    hian = hian + 1
    If gameflg = 1 Then
        If nowattack = 1 Then
            Range("Kseie3").Value = CStr(hian)
        Else
            Range("Sseie3").Value = CStr(hian)
        End If
    End If
    Call 得点
    pitch = pitch + 1
    If gameflg = 1 Then
        If nowattack = 1 Then
            Range("Kseie1").Value = CStr(pitch)
        Else
            Range("Sseie1").Value = CStr(pitch)
        End If
    End If
    ball = 0
    strike = 0
    Call ボール点灯
    Call ストライク点灯
    Call 成績表示
    Call 本日結果表示
    Range("kekkah").Value = ""
    ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    Call gameset
End Sub
Sub エラー()
    kiroku = Range("kiroku").Value
    If Range("eumu").Value <> "" Then
        Range("kekkae").Font.Name = "Arial Narrow"
    End If
    If InStr(kiroku, "投") >= 1 Then
        Range("kekkae").Value = "E1"
    ElseIf InStr(kiroku, "捕") >= 1 Then
        Range("kekkae").Value = "E2"
    ElseIf InStr(kiroku, "一") >= 1 Then
        Range("kekkae").Value = "E3"
    ElseIf InStr(kiroku, "二") >= 1 Then
        Range("kekkae").Value = "E4"
    ElseIf InStr(kiroku, "三") >= 1 Then
        Range("kekkae").Value = "E5"
    ElseIf InStr(kiroku, "遊") >= 1 Then
        Range("kekkae").Value = "E6"
    ElseIf InStr(kiroku, "左") >= 1 Then
        Range("kekkae").Value = "E7"
    ElseIf InStr(kiroku, "中") >= 1 Then
        Range("kekkae").Value = "E8"
    ElseIf InStr(kiroku, "右") >= 1 Then
        Range("kekkae").Value = "E9"
    Else
        Range("kekkae").Value = "E"
    End If
    If Range("eumu").Value <> "" Then
        Range("kekkae").Value = Range("kekkae").Value & Range("eumu").Value
        If nowattack = 1 Then
            Kerror = Kerror + 1
        Else
            Serror = Serror + 1
        End If
    End If
    Application.Wait [now()+"0:00:05"]
    If nowattack = 1 Then
        Kerror = Kerror + 1
        intcnt = Len(CStr(Kerror))
        Call IntFont
        Range("Kerror").Font.Name = fontname
        Range("Kerror").Value = CStr(Kerror)
    Else
        Serror = Serror + 1
        intcnt = Len(CStr(Serror))
        Call IntFont
        Range("Serror").Font.Name = fontname
        Range("Serror").Value = CStr(Serror)
    End If
    Call 出塁
    Range("kekkae").Value = ""
    Range("eumu").Value = ""
    Range("kekkae").Font.Name = "Arial"
    Call gameset
End Sub
Sub 野選()
    Range("kekkafc").Value = "Fc"
    Application.Wait [now()+"0:00:05"]
    Call 出塁
    Range("kekkafc").Value = ""
    Call gameset
End Sub
Sub 盗塁死()
    touruisi = Range("touruisi").Value
    With Worksheets("試合情報")
        .Range("H6:H8").Value = ""
        .Range("F3").Value = nowinning
        .Range("F4").Value = nowattack
        .Range("F5").Value = nowbatter
        .Range("F6").Value = ball
        .Range("F7").Value = strike
        .Range("F8").Value = out
        .Range("H3").Value = kiroku
        .Range("H4").Value = daten
        .Range("H5").Value = hian
        .Range("J3").Value = pitch
        .Range("J4").Value = k
        .Range("J5").Value = si
        .Range("J6").Value = pitchinning
        .Range("J7").Value = inpitch
    End With
        
    a = 15
    b = 1
    
    runner1 = 0
    runner2 = 0
    runner3 = 0
    
    If nowattack = 1 Then
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern = 1 Then
                If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 13).Interior.Pattern <> 1 And Cells(a, 24).Interior.Pattern = 1 Then
                If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 13).Interior.Pattern = 1 And Cells(a, 24).Interior.Pattern <> 1 Then
                If Cells(a, 13).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            Else
                If Cells(a, 13).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 13).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 13).Font.Color = runner_pos_Font _
                And Cells(a, 13).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 24).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 24).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 24).Font.Color = runner_name_Font _
                And Cells(a, 24).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            End If
            b = b + 1
        Next a
    Else
        For a = 23 To 111 Step 11
            sw = 0
            If Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 249).Interior.Pattern <> 1 And Cells(a, 260).Interior.Pattern = 1 Then
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            ElseIf Cells(a, 249).Interior.Pattern = 1 And Cells(a, 260).Interior.Pattern <> 1 Then
                If Cells(a, 249).Interior.Color = runner_pos_Interior.Color _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            Else
                If Cells(a, 249).Interior.Gradient.ColorStops.Item(1).Color = runner_pos_Color1 _
                And Cells(a, 249).Interior.Gradient.ColorStops.Item(2).Color = runner_pos_Color2 _
                And Cells(a, 249).Font.Color = runner_pos_Font _
                And Cells(a, 249).Borders.Color = runner_pos_Borders.Color _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(1).Color = runner_name_Color1 _
                And Cells(a, 260).Interior.Gradient.ColorStops.Item(2).Color = runner_name_Color2 _
                And Cells(a, 260).Interior.Color = runner_name_Interior.Color _
                And Cells(a, 260).Font.Color = runner_name_Font _
                And Cells(a, 260).Borders.Color = runner_name_Borders.Color Then
                    If nowbatter <> b Then
                        If runner1 = 0 And sw = 0 Then
                            runner1 = b
                            sw = 1
                        End If
                        If runner2 = 0 And sw = 0 Then
                            runner2 = b
                            sw = 1
                        End If
                        If runner3 = 0 And sw = 0 Then
                            runner3 = b
                        End If
                    End If
                End If
            End If
            b = b + 1
        Next a
    End If
    
    With Worksheets("試合情報")
        If runner1 <> 0 Then
            .Range("H6").Value = runner1
            If runner2 <> 0 Then
                .Range("H7").Value = runner2
                If runner3 <> 0 Then
                    .Range("H8").Value = runner3
                End If
            End If
        End If
    End With
    If nowattack = 1 Then
        Call returncells(23 + CInt(((touruisi - 1) * 11)), 13, basic_pos_Interior, basic_pos_Font, _
                         basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
        Call returncells(23 + CInt(((touruisi - 1) * 11)), 24, basic_name_Interior, basic_name_Font, _
                         basic_name_Borders, basic_name_Color1, basic_name_Color2)
    Else
        Call returncells(23 + CInt(((touruisi - 1) * 11)), 249, basic_pos_Interior, basic_pos_Font, _
                         basic_pos_Borders, basic_pos_Color1, basic_pos_Color2)
        Call returncells(23 + CInt(((touruisi - 1) * 11)), 260, basic_name_Interior, basic_name_Font, _
                         basic_name_Borders, basic_name_Color1, basic_name_Color2)
    End If
    Range("touruisi") = ""
    out = out + 1
    pitchinning = pitchinning + 1
    If nowattack = 1 Then
        If pitchinning Mod 3 = 0 Then
            Range("Kseie6").Value = pitchinning / 3
        Else
            Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    Else
        If pitchinning Mod 3 = 0 Then
            Range("Sseie6").Value = pitchinning / 3
        Else
            Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    End If
    Call アウト点灯
    Call gameset
    Call 攻守交代
End Sub
Sub 申告敬遠()
    ball = 0
    strike = 0
    Call 出塁
    si = si + 1
    If nowattack = 1 Then
        Range("Kseie5").Value = CStr(si)
    Else
        Range("Sseie5").Value = CStr(si)
    End If
    Call 走塁死
    Call ボール点灯
    Call ストライク点灯
    Call アウト点灯
    Call 得点
    Call 成績表示
    Call 本日結果表示

End Sub
Sub 犠打()
    Call 記録代入
    daten = Range("daten").Value
    If seikasan = 1 Then
        If nowattack = 1 Then
            seban = Range("Snum").Value
            Call updateinfo(Steam, seban, 10, daten)
        Else
            seban = Range("Knum").Value
            Call updateinfo(Kteam, seban, 10, daten)
        End If
    End If
    If InStr(kiroku, "失") >= 1 Then
        Call エラー
    Else
        Call 次打者
        pitch = pitch + 1
        pitchinning = pitchinning + 1
        If nowattack = 1 Then
            Range("Kseie1").Value = CStr(pitch)
            Range("Kseie4").Value = CStr(inpitch)
            If pitchinning Mod 3 = 0 Then
                Range("Kseie6").Value = pitchinning / 3
            Else
                Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
        Else
            Range("Sseie1").Value = CStr(pitch)
            Range("Sseie4").Value = CStr(inpitch)
            If pitchinning Mod 3 = 0 Then
                Range("Sseie6").Value = pitchinning / 3
            Else
                Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
        End If
        out = out + 1
    End If
    Worksheets("試合情報").Range("A1").Value = ""
    Call 得点
    Call 成績表示
    Call 本日結果表示
    ball = 0
    strike = 0
    Call ボール点灯
    Call ストライク点灯
    Call アウト点灯
    Call gameset
    Range("daten").Value = ""
End Sub
Sub 犠飛()
    Call 記録代入
    daten = Range("daten").Value
    If seikasan = 1 Then
        If nowattack = 1 Then
            seban = Range("Snum").Value
            Call updateinfo(Steam, seban, 15, 1)
            Call updateinfo(Steam, seban, 10, daten)
        Else
            seban = Range("Knum").Value
            Call updateinfo(Kteam, seban, 15, 1)
            Call updateinfo(Kteam, seban, 10, daten)
        End If
    End If
    Call 次打者
    pitch = pitch + 1
    pitchinning = pitchinning + 1
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
        If pitchinning Mod 3 = 0 Then
            Range("Kseie6").Value = pitchinning / 3
        Else
            Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
        If pitchinning Mod 3 = 0 Then
            Range("Sseie6").Value = pitchinning / 3
        Else
            Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
        End If
    End If
    Worksheets("試合情報").Range("A1").Value = 1
    Call 得点
    Call 成績表示
    Call 本日結果表示
    ball = 0
    strike = 0
    out = out + 1
    Call ボール点灯
    Call ストライク点灯
    Call アウト点灯
    Call gameset
    Range("daten").Value = ""
End Sub

Sub ファール()
    pitch = pitch + 1
    inpitch = inpitch + 1
    If strike < 2 Then
        strike = strike + 1
    End If
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
    End If
    Call ストライク点灯
End Sub
Sub エラー2()
    Range("kekkae").Value = "E"
    eumu = Range("eumu")
    If nowattack = 1 Then
        Kerror = Kerror + 1
        intcnt = Len(CStr(Kerror))
        Call IntFont
        Range("Kerror").Font.Name = fontname
        Range("Kerror").Value = CStr(Kerror)
    Else
        Serror = Serror + 1
        intcnt = Len(CStr(Serror))
        Call IntFont
        Range("Serror").Font.Name = fontname
        Range("Serror").Value = CStr(Serror)
    End If
    Call 得点
    Application.Wait [now()+"0:00:05"]
    Range("kekkae").Value = ""
End Sub
Sub 振り逃げ()
    Call 出塁
    k = k + 1
    If nowattack = 1 Then
        Range("Kseie2").Value = CStr(k)
    Else
        Range("Sseie2").Value = CStr(k)
    End If
End Sub
Sub alldelete()
    gameflg = 0
    Dim myRng As Range
    Dim sp As Variant
    Set myRng = Range("A1:MA43")
    For Each sp In ActiveSheet.Shapes
    If Not Intersect(Range(sp.TopLeftCell, sp.BottomRightCell), myRng) Is Nothing Then
    sp.Delete
    End If
    Next
    Set myRng = Nothing
    Range("B2:KW138").Interior.Color = RGB(0, 0, 0)
    Range("B2:KW138").Borders.LineStyle = -4142
    Range("Steam1,Steam2,Kteam1,Kteam2,M23:X122,BP17:XY28,IO23:IZ122,BP46:CP97,GJ46:HJ97,DT75:GE108,BP117:CZ125,GJ117:HT125").Value = ""
    Range("inin1").Value = "1"
    Range("inin2").Value = "2"
    Range("inin3").Value = "3"
    Range("inin4").Value = "4"
    Range("inin5").Value = "5"
    Range("inin6").Value = "6"
    Range("inin7").Value = "7"
    Range("inin8").Value = "8"
    Range("inin9").Value = "9"
    With Worksheets("試合情報")
        .Range("B3:C7").Value = 0
        .Range("A11:R17,F3:F8,H3:H8,J3:J7").Value = ""
        .Range("B8:C8").Value = 1
    End With
    ball = 0
    strike = 0
    out = 0
    Erase SrArray()
    Erase KrArray()
    Shit = 0
    Serror = 0
    Slob = 0
    Khit = 0
    Kerror = 0
    Klob = 0
    inpitch = 0
    ActiveSheet.Shapes("一塁").Visible = False
    ActiveSheet.Shapes("二塁").Visible = False
    ActiveSheet.Shapes("三塁").Visible = False
    Range("CW10:HY10,EL113:FG127").Font.Color = RGB(0, 0, 0)
    Range("CW10:HY10,EL113:FG127").Value = ""
    Range("CW17:HY28").Font.Color = RGB(255, 255, 255)
    Range("CW17:HY28").Font.Name = "Arial"
End Sub

Sub gameset_2()
    Dim inning As Integer
    inning = nowinning Mod 9
    If nowattack = 1 Then
        Select Case inning
        Case 0
            Range("Srun9").Font.Name = "Arial Narrow"
            Range("Srun9").Value = SrArray(nowinning) & "x"
        Case 1
            Range("Srun1").Font.Name = "Arial Narrow"
            Range("Srun1").Value = SrArray(nowinning) & "x"
        Case 2
            Range("Srun2").Font.Name = "Arial Narrow"
            Range("Srun2").Value = SrArray(nowinning) & "x"
        Case 3
            Range("Srun3").Font.Name = "Arial Narrow"
            Range("Srun3").Value = SrArray(nowinning) & "x"
        Case 4
            Range("Srun4").Font.Name = "Arial Narrow"
            Range("Srun4").Value = SrArray(nowinning) & "x"
        Case 5
            Range("Srun5").Font.Name = "Arial Narrow"
            Range("Srun5").Value = SrArray(nowinning) & "x"
        Case 6
            Range("Srun6").Font.Name = "Arial Narrow"
            Range("Srun6").Value = SrArray(nowinning) & "x"
        Case 7
            Range("Srun7").Font.Name = "Arial Narrow"
            Range("Srun7").Value = SrArray(nowinning) & "x"
        Case 8
            Range("Srun8").Font.Name = "Arial Narrow"
            Range("Srun8").Value = SrArray(nowinning) & "x"
        End Select
        Range("Sline").Interior.Color = back_Interior
        returnflg = 1
        Call MemberCellReset
        ball = 0
        strike = 0
        out = 0
        Call ボール点灯
        Call ストライク点灯
        Call アウト点灯
        Call ScoreCellReset
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    Else
        Select Case inning
        Case 0
            Range("Krun9").Font.Name = "Arial Narrow"
            Range("Krun9").Value = KrArray(nowinning) & "x"
        Case 1
            Range("Krun1").Font.Name = "Arial Narrow"
            Range("Krun1").Value = KrArray(nowinning) & "x"
        Case 2
            Range("Krun2").Font.Name = "Arial Narrow"
            Range("Krun2").Value = KrArray(nowinning) & "x"
        Case 3
            Range("Krun3").Font.Name = "Arial Narrow"
            Range("Krun3").Value = KrArray(nowinning) & "x"
        Case 4
            Range("Krun4").Font.Name = "Arial Narrow"
            Range("Krun4").Value = KrArray(nowinning) & "x"
        Case 5
            Range("Krun5").Font.Name = "Arial Narrow"
            Range("Krun5").Value = KrArray(nowinning) & "x"
        Case 6
            Range("Krun6").Font.Name = "Arial Narrow"
            Range("Krun6").Value = KrArray(nowinning) & "x"
        Case 7
            Range("Krun7").Font.Name = "Arial Narrow"
            Range("Krun7").Value = KrArray(nowinning) & "x"
        Case 8
            Range("Krun8").Font.Name = "Arial Narrow"
            Range("Krun8").Value = KrArray(nowinning) & "x"
        End Select
        Range("Kline").Interior.Color = back_Interior
        returnflg = 1
        Call MemberCellReset
        ball = 0
        strike = 0
        out = 0
        Call ボール点灯
        Call ストライク点灯
        Call アウト点灯
        Call ScoreCellReset
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    End If
    Range("BP46:DO108,GJ46:II108,BP117:DO131,GJ117:II131").Borders.LineStyle = -4142
    Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Interior.Color = back_Interior
    Range("BP46:CP97,GJ46:HJ97,BP117:CZ125,GJ117:HT125").Value = ""
    If nowinning = 5 And nowattack = 2 And SrArray(0) < KrArray(0) - KrArray(nowinning) Then
        Range("CW28:HY28").Font.Color = RGB(255, 255, 0)
    End If
    If nowinning >= 6 Then
        If SrArray(0) - SrArray(nowinning) < KrArray(0) Then
            Range("CW28:HY28").Font.Color = RGB(255, 255, 0)
        ElseIf SrArray(0) - SrArray(nowinning) > KrArray(0) Then
            Range("CW17:HY17").Font.Color = RGB(255, 255, 0)
        End If
    End If
    ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)

End Sub
Sub ballminus()
    If pitch > 0 Then
        pitch = pitch - 1
    End If
    If inpitch > 0 Then
        inpitch = inpitch - 1
    End If
    If ball > 0 Then
        ball = ball - 1
    End If
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
    End If
    Call ボール点灯
End Sub
Sub strikeminus()
    If pitch > 0 Then
        pitch = pitch - 1
    End If
    If inpitch > 0 Then
        inpitch = inpitch - 1
    End If
    If strike > 0 Then
        strike = strike - 1
    End If
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
    End If
    Call ストライク点灯
End Sub
Sub outminus()
    If out > 0 Then
        out = out - 1
    End If
    Call アウト点灯
End Sub
Sub minus()
    If pitch > 0 Then
        pitch = pitch - 1
    End If
    If inpitch > 0 Then
        inpitch = inpitch - 1
    End If
    If nowattack = 1 Then
        Range("Kseie1").Value = CStr(pitch)
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Sseie1").Value = CStr(pitch)
        Range("Sseie4").Value = CStr(inpitch)
    End If
End Sub
Sub inpitchminus()
    If inpitch > 0 Then
        inpitch = inpitch - 1
    End If
    If nowattack = 1 Then
        Range("Kseie4").Value = CStr(inpitch)
    Else
        Range("Kseie4").Value = CStr(inpitch)
    End If
End Sub
Sub タイブレーク()
    If nowattack = 1 Then
        If nowbatter >= 3 Then
            Call returncells(23 + ((nowbatter - 2) * 11), 13, runner_pos_Interior, runner_pos_Font, _
                             runner_pos_Borders, runner_pos_Color1, runner_pos_Color2)
            Call returncells(23 + ((nowbatter - 2) * 11), 24, runner_name_Interior, runner_name_Font, _
                             runner_name_Borders, runner_name_Color1, runner_name_Color2)
            Call returncells(23 + ((nowbatter - 3) * 11), 13, runner_pos_Interior, runner_pos_Font, _
                             runner_pos_Borders, runner_pos_Color1, runner_pos_Color2)
            Call returncells(23 + ((nowbatter - 3) * 11), 24, runner_name_Interior, runner_name_Font, _
                             runner_name_Borders, runner_name_Color1, runner_name_Color2)
        Else
            If nowbatter = 1 Then
                Call returnrange("Spos8", runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returnrange("Sname8", runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            Else
                Call returnrange("Spos1", runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returnrange("Sname1", runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            Call returnrange("Spos9", runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                             runner_pos_Color1, runner_pos_Color2)
            Call returnrange("Sname9", runner_name_Interior, runner_name_Font, runner_name_Borders, _
                             runner_name_Color1, runner_name_Color2)
        End If
    Else
        If nowbatter >= 3 Then
            Call returncells(23 + ((nowbatter - 2) * 11), 249, runner_pos_Interior, runner_pos_Font, _
                             runner_pos_Borders, runner_pos_Color1, runner_pos_Color2)
            Call returncells(23 + ((nowbatter - 2) * 11), 260, runner_name_Interior, runner_name_Font, _
                             runner_name_Borders, runner_name_Color1, runner_name_Color2)
            Call returncells(23 + ((nowbatter - 3) * 11), 249, runner_pos_Interior, runner_pos_Font, _
                             runner_pos_Borders, runner_pos_Color1, runner_pos_Color2)
            Call returncells(23 + ((nowbatter - 3) * 11), 260, runner_name_Interior, runner_name_Font, _
                             runner_name_Borders, runner_name_Color1, runner_name_Color2)
        Else
            If nowbatter = 1 Then
                Call returnrange("Kpos8", runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returnrange("Kname8", runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            Else
                Call returnrange("Kpos1", runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returnrange("Kname1", runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            Call returnrange("Kpos9", runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                             runner_pos_Color1, runner_pos_Color2)
            Call returnrange("Kname9", runner_name_Interior, runner_name_Font, runner_name_Borders, _
                             runner_name_Color1, runner_name_Color2)
        End If
    End If
End Sub
Sub RunScore()

    If (nowinning Mod 9) = 1 Then
        If nowattack = 1 Then
            Call returnrange("Krun9", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun1").Interior.Color = RGB(255, 0, 0)
            Range("Srun1").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun1", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun1").Interior.Color = RGB(255, 0, 0)
            Range("Krun1").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 2 Then
        If nowattack = 1 Then
            Call returnrange("Krun1", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun2").Interior.Color = RGB(255, 0, 0)
            Range("Srun2").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun2", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun2").Interior.Color = RGB(255, 0, 0)
            Range("Krun2").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 3 Then
        If nowattack = 1 Then
            Call returnrange("Krun2", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun3").Interior.Color = RGB(255, 0, 0)
            Range("Srun3").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun3", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun3").Interior.Color = RGB(255, 0, 0)
            Range("Krun3").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 4 Then
        If nowattack = 1 Then
            Call returnrange("Krun3", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun4").Interior.Color = RGB(255, 0, 0)
            Range("Srun4").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun4", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun4").Interior.Color = RGB(255, 0, 0)
            Range("Krun4").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 5 Then
        If nowattack = 1 Then
            Call returnrange("Krun4", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun5").Interior.Color = RGB(255, 0, 0)
            Range("Srun5").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun5", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun5").Interior.Color = RGB(255, 0, 0)
            Range("Krun5").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 6 Then
        If nowattack = 1 Then
            Call returnrange("Krun5", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun6").Interior.Color = RGB(255, 0, 0)
            Range("Srun6").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun6", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun6").Interior.Color = RGB(255, 0, 0)
            Range("Krun6").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 7 Then
        If nowattack = 1 Then
            Call returnrange("Krun6", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun7").Interior.Color = RGB(255, 0, 0)
            Range("Srun7").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun7", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun7").Interior.Color = RGB(255, 0, 0)
            Range("Krun7").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 8 Then
        If nowattack = 1 Then
            Call returnrange("Krun7", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun8").Interior.Color = RGB(255, 0, 0)
            Range("Srun8").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun8", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun8").Interior.Color = RGB(255, 0, 0)
            Range("Krun8").Font.Color = RGB(255, 255, 255)
        End If
    ElseIf (nowinning Mod 9) = 0 Then
        If nowattack = 1 Then
            Call returnrange("Krun8", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Srun9").Interior.Color = RGB(255, 0, 0)
            Range("Srun9").Font.Color = RGB(255, 255, 255)
        ElseIf nowattack = 2 Then
            Call returnrange("Srun9", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                             score_entity_Color1, score_entity_Color2)
            Range("Krun9").Interior.Color = RGB(255, 0, 0)
            Range("Krun9").Font.Color = RGB(255, 255, 255)
        End If
    End If
End Sub

Sub 修正()
        
    Dim koumoku As String
    
    koumoku = InputBox("修正する項目を選択してください" & vbCr & _
                       "得点なら1を、安打なら2を、失策なら3を、残塁なら4を半角で入力してください")
    
    If koumoku = Null Then
        Exit Sub
    End If
    
    If koumoku = 1 Then
        Call 得点修正
    ElseIf koumoku = 2 Then
        Call 安打修正
    ElseIf koumoku = 3 Then
        Call 失策修正
    Else
        Call 残塁修正
    End If
                
End Sub
Sub 得点修正()
    
    Dim team As String
    Dim inning_after As Integer
    Dim value_after As Integer
    
    team = InputBox("記録を修正するチームは？" & vbCr & _
                    "先攻チームなら1を、後攻チームなら2を半角で入力")
    inning_after = InputBox("得点を修正するイニングを半角で入力してください")
    value_after = InputBox("修正後の数値を半角で入力してください")
    
    If team = Null Or inning_after = Null Or value_after = Null Then
        Exit Sub
    End If
    
    '得点修正
    If team = 1 Then
        SrArray(inning_after) = value_after
        intcnt = Len(CStr(SrArray(inning_after)))
        Call IntFont
    Else
        KrArray(inning_after) = value_after
        intcnt = Len(CStr(KrArray(inning_after)))
        Call IntFont
    End If

    '得点表示
    Dim a As Integer
    Dim inning As Integer
    inning = inning_after Mod 9
    If team = 1 Then
        If inning = 1 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun1").Font.Name = fontname
            Range("Srun1").Value = CStr(SrArray(inning_after))
        ElseIf inning = 2 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun2").Font.Name = fontname
            Range("Srun2").Value = CStr(SrArray(inning_after))
        ElseIf inning = 3 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun3").Font.Name = fontname
            Range("Srun3").Value = CStr(SrArray(inning_after))
        ElseIf inning = 4 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun4").Font.Name = fontname
            Range("Srun4").Value = CStr(SrArray(inning_after))
        ElseIf inning = 5 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun5").Font.Name = fontname
            Range("Srun5").Value = CStr(SrArray(inning_after))
        ElseIf inning = 6 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun6").Font.Name = fontname
            Range("Srun6").Value = CStr(SrArray(inning_after))
        ElseIf inning = 7 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun7").Font.Name = fontname
            Range("Srun7").Value = CStr(SrArray(inning_after))
        ElseIf inning = 8 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun8").Font.Name = fontname
            Range("Srun8").Value = CStr(SrArray(inning_after))
        ElseIf inning = 0 Then
            intcnt = Len(CStr(SrArray(inning_after)))
            Call IntFont
            Range("Srun9").Font.Name = fontname
            Range("Srun9").Value = CStr(SrArray(inning_after))
        End If
        
        SrArray(0) = 0
        a = 1
        Do While a < UBound(SrArray)
            SrArray(0) = SrArray(0) + SrArray(a)
            a = a + 1
        Loop
        intcnt = Len(CStr(SrArray(0)))
        Call IntFont
        Range("Srunr").Font.Name = fontname
        Range("Srunr").Value = CStr(SrArray(0))
        
    Else
        If inning_after = 1 Or Right(inning_after, 1) = 0 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun1").Font.Name = fontname
            Range("Krun1").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 2 Or Right(inning_after, 1) = 1 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun2").Font.Name = fontname
            Range("Krun2").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 3 Or Right(inning_after, 1) = 2 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun3").Font.Name = fontname
            Range("Krun3").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 4 Or Right(inning_after, 1) = 3 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun4").Font.Name = fontname
            Range("Krun4").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 5 Or Right(inning_after, 1) = 4 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun5").Font.Name = fontname
            Range("Krun5").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 6 Or Right(inning_after, 1) = 5 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun6").Font.Name = fontname
            Range("Krun6").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 7 Or Right(inning_after, 1) = 6 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun7").Font.Name = fontname
            Range("Krun7").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 8 Or Right(inning_after, 1) = 7 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun8").Font.Name = fontname
            Range("Krun8").Value = CStr(KrArray(inning_after))
        ElseIf inning_after = 9 Or Right(inning_after, 1) = 8 Then
            intcnt = Len(CStr(KrArray(inning_after)))
            Call IntFont
            Range("Krun9").Font.Name = fontname
            Range("Krun9").Value = CStr(KrArray(inning_after))
        End If
        
        KrArray(0) = 0
        a = 1
        Do While a < UBound(KrArray)
            KrArray(0) = KrArray(0) + KrArray(a)
            a = a + 1
        Loop
        intcnt = Len(CStr(KrArray(0)))
        Call IntFont
        Range("Krunr").Font.Name = fontname
        Range("Krunr").Value = CStr(KrArray(0))
    End If
End Sub
Sub 安打修正()
    
    Dim team As String
    Dim value_after As Integer
    
    team = InputBox("記録を修正するチームは？" & vbCr & _
                    "先攻チームなら1を、後攻チームなら2を半角で入力")
    value_after = InputBox("修正後の数値を半角で入力してください")
    
    If team = Null Or value_after = Null Then
        Exit Sub
    End If
    
    If team = 1 Then
        Shit = value_after
        intcnt = Len(CStr(Shit))
        Call IntFont
        Range("Shit").Font.Name = fontname
        Range("Shit").Value = CStr(Shit)
    Else
        Khit = value_after
        intcnt = Len(CStr(Khit))
        Call IntFont
        Range("Khit").Font.Name = fontname
        Range("Khit").Value = CStr(Khit)
    End If

End Sub
Sub 失策修正()
    
    Dim team As String
    Dim value_after As Integer
    
    team = InputBox("記録を修正するチームは？" & vbCr & _
                    "先攻チームなら1を、後攻チームなら2を半角で入力")
    value_after = InputBox("修正後の数値を半角で入力してください")
    
    If team = Null Or value_after = Null Then
        Exit Sub
    End If
    
    If team = 1 Then
        Serror = value_after
        intcnt = Len(CStr(Serror))
        Call IntFont
        Range("Serror").Font.Name = fontname
        Range("Serror").Value = CStr(Serror)
    Else
        Kerror = value_after
        intcnt = Len(CStr(Kerror))
        Call IntFont
        Range("Kerror").Font.Name = fontname
        Range("Kerror").Value = CStr(Kerror)
    End If

End Sub
Sub 残塁修正()
    
    Dim team As String
    Dim value_after As Integer
    
    team = InputBox("記録を修正するチームは？" & vbCr & _
                    "先攻チームなら1を、後攻チームなら2を半角で入力")
    value_after = InputBox("修正後の数値を半角で入力してください")
    
    If team = Null Or value_after = Null Then
        Exit Sub
    End If
    
    If team = 1 Then
        Slob = value_after
        intcnt = Len(CStr(Slob))
        Call IntFont
        Range("Slob").Font.Name = fontname
        Range("Slob").Value = CStr(Slob)
    Else
        Klob = value_after
        intcnt = Len(CStr(Klob))
        Call IntFont
        Range("Klob").Font.Name = fontname
        Range("Klob").Value = CStr(Klob)
    End If

End Sub

Sub 移動()
    Dim a As Integer
    Dim b As Integer
    Dim max As Integer
    Dim team As String
    
    team = Worksheets("成績仮置き場").Range("N2").Value
    a = 5
    b = 4
    
    With Worksheets(team)
    max = Cells(Rows.count, 2).End(xlUp).Row
    Maxrow = .Cells(Rows.count, 3).End(xlUp).Row
        For a = 5 To max
            For b = 4 To Maxrow
                If CStr(Cells(a, 2)) = CStr(.Cells(b, 2).Value) Then
                    '打数
                    If Cells(a, 7).Value = "-" Then
                        .Cells(b, 5).Value = 0
                    Else
                        .Cells(b, 5).Value = Cells(a, 7).Value
                    End If
                    '安打
                    If Cells(a, 8).Value = "-" Then
                        .Cells(b, 6).Value = 0
                    Else
                        .Cells(b, 6).Value = Cells(a, 8).Value
                    End If
                    '二塁打
                    If Cells(a, 9).Value = "-" Then
                        .Cells(b, 7).Value = 0
                    Else
                        .Cells(b, 7).Value = Cells(a, 9).Value
                    End If
                    '三塁打
                    If Cells(a, 10).Value = "-" Then
                        .Cells(b, 8).Value = 0
                    Else
                        .Cells(b, 8).Value = Cells(a, 10).Value
                    End If
                    '本塁打
                    If Cells(a, 11).Value = "-" Then
                        .Cells(b, 9).Value = 0
                    Else
                        .Cells(b, 9).Value = Cells(a, 11).Value
                    End If
                    '打点
                    If Cells(a, 13).Value = "-" Then
                        .Cells(b, 10).Value = 0
                    Else
                        .Cells(b, 10).Value = Cells(a, 13).Value
                    End If
                    '四球
                    If Cells(a, 16).Value = "-" Then
                        .Cells(b, 13).Value = 0
                    Else
                        .Cells(b, 13).Value = Cells(a, 16).Value
                    End If
                    '死球
                    If Cells(a, 17).Value = "-" Then
                        .Cells(b, 14).Value = 0
                    Else
                        .Cells(b, 14).Value = Cells(a, 17).Value
                    End If
                    '犠飛
                    If Cells(a, 19).Value = "-" Then
                        .Cells(b, 15).Value = 0
                    Else
                        .Cells(b, 15).Value = Cells(a, 19).Value
                    End If

                End If
            Next b
        Next a
    End With

End Sub

Sub 移動消去()
    
    Worksheets("成績仮置き場").Range("B5:AA75").Value = ""

End Sub

Sub reverse()
    
    Dim nextlead As Integer
    If Worksheets("試合情報").Range("F3").Value <> "" Then
        With Worksheets("試合情報")
            nowinning = .Range("F3").Value
            nowattack = .Range("F4").Value
            nowbatter = .Range("F5").Value
            ball = .Range("F6").Value
            strike = .Range("F7").Value
            out = .Range("F8").Value
            runner1 = .Range("H6").Value
            runner2 = .Range("H7").Value
            runner3 = .Range("H8").Value
            kiroku = .Range("H3").Value
            daten = .Range("H4").Value
        End With
        
        a = 23
        
        If nowattack = 1 Then
            Range("Sline").Interior.Color = RGB(255, 0, 0)
            Range("Kline").Interior.Color = back_Interior
            returnflg = 1
            Call MemberCellReset
            Range("BP117:CZ125").Interior.Color = back_Interior
            Range("BP117:DO131").Borders.LineStyle = -4142
            Range("Stoday1,Stoday2,Stoday3,Stoday4,Stoday5,Stoday6").Value = ""
            Range("GJ117:HT125").Interior.Color = back_Interior
            Range("GJ117:II131").Borders.LineStyle = -4142
            Range("Ktoday1,Ktoday2,Ktoday3,Ktoday4,Ktoday5,Ktoday6").Value = ""
            Call returncells(23 + ((nowbatter - 1) * 11), 13, batter_pos_Interior, batter_pos_Font, batter_pos_Borders, _
                             batter_pos_Color1, batter_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 24, batter_name_Interior, batter_name_Font, batter_name_Borders, _
                             batter_name_Color1, batter_name_Color2)
            If runner1 >= 1 Then
                Call returncells(23 + ((runner1 - 1) * 11), 13, runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returncells(23 + ((runner1 - 1) * 11), 24, runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            If runner2 >= 1 Then
                Call returncells(23 + ((runner2 - 1) * 11), 13, runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returncells(23 + ((runner2 - 1) * 11), 24, runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            If runner3 >= 1 Then
                Call returncells(23 + ((runner3 - 1) * 11), 13, runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returncells(23 + ((runner3 - 1) * 11), 24, runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            nextlead = Worksheets("試合情報").Range("C8").Value
            Call returncells(23 + ((nextlead - 1) * 11), 249, next_pos_Interior, next_pos_Font, next_pos_Borders, _
                             next_pos_Color1, next_pos_Color2)
            Call returncells(23 + ((nextlead - 1) * 11), 260, next_name_Interior, next_name_Font, next_name_Borders, _
                             next_name_Color1, next_name_Color2)
        Else
            Range("Sline").Interior.Color = back_Interior
            Range("Kline").Interior.Color = RGB(255, 0, 0)
            returnflg = 1
            Call MemberCellReset
            Range("BP117:CZ125").Interior.Color = back_Interior
            Range("BP117:DO131").Borders.LineStyle = -4142
            Range("Stoday1,Stoday2,Stoday3,Stoday4,Stoday5,Stoday6").Value = ""
            Range("GJ117:HT125").Interior.Color = back_Interior
            Range("GJ117:II131").Borders.LineStyle = -4142
            Range("Ktoday1,Ktoday2,Ktoday3,Ktoday4,Ktoday5,Ktoday6").Value = ""
            Call returncells(23 + ((nowbatter - 1) * 11), 249, batter_pos_Interior, batter_pos_Font, batter_pos_Borders, _
                             batter_pos_Color1, batter_pos_Color2)
            Call returncells(23 + ((nowbatter - 1) * 11), 260, batter_name_Interior, batter_name_Font, batter_name_Borders, _
                             batter_name_Color1, batter_name_Color2)
            If runner1 >= 1 Then
                Call returncells(23 + ((runner1 - 1) * 11), 249, runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returncells(23 + ((runner1 - 1) * 11), 260, runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            If runner2 >= 1 Then
                Call returncells(23 + ((runner2 - 1) * 11), 249, runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returncells(23 + ((runner2 - 1) * 11), 260, runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            If runner3 >= 1 Then
                Call returncells(23 + ((runner3 - 1) * 11), 249, runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                                 runner_pos_Color1, runner_pos_Color2)
                Call returncells(23 + ((runner3 - 1) * 11), 260, runner_name_Interior, runner_name_Font, runner_name_Borders, _
                                 runner_name_Color1, runner_name_Color2)
            End If
            nextlead = Worksheets("試合情報").Range("B8").Value
            Call returncells(23 + ((nextlead - 1) * 11), 13, next_pos_Interior, next_pos_Font, next_pos_Borders, _
                             next_pos_Color1, next_pos_Color2)
            Call returncells(23 + ((nextlead - 1) * 11), 24, next_name_Interior, next_name_Font, next_name_Borders, _
                             next_name_Color1, next_name_Color2)
        End If
        
        With Worksheets("試合情報")
            If .Range("H3").Value <> "" Then
                If nowattack = 1 Then
                    Maxrow = .Cells(Rows.count, nowbatter).End(xlUp).Row
                    If .Cells((Maxrow), nowbatter).Value = kiroku Then
                        .Cells((Maxrow), nowbatter).Value = ""
                    End If
                Else
                    Maxrow = .Cells(Rows.count, (nowbatter + 9)).End(xlUp).Row
                    If .Cells((Maxrow), (nowbatter + 9)).Value = kiroku Then
                        .Cells((Maxrow), (nowbatter + 9)).Value = ""
                    End If
                End If
            End If
        End With
           
        Call ボール点灯
        Call ストライク点灯
        Call アウト点灯
        Call 成績表示
        Call 投手成績表示
        With Worksheets("試合情報")
            hian = .Range("H5").Value
            pitch = .Range("J3").Value
            k = .Range("J4").Value
            si = .Range("J5").Value
            pitchinning = .Range("J6").Value
            inpitch = .Range("J7").Value
        End With
        If nowattack = 1 Then
            Range("Kseie1").Value = pitch
            Range("Kseie2").Value = k
            Range("Kseie3").Value = hian
            Range("Kseie4").Value = inpitch
            Range("Kseie5").Value = si
            If pitchinning Mod 3 = 0 Then
                Range("Kseie6").Value = pitchinning / 3
            Else
                Range("Kseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
            seban = CStr(Cells(143 + ((nowbatter - 1) * 4), 10).Value)
        Else
            Range("Sseie1").Value = pitch
            Range("Sseie2").Value = k
            Range("Sseie3").Value = hian
            Range("Sseie4").Value = inpitch
            Range("Sseie5").Value = si
            If pitchinning Mod 3 = 0 Then
                Range("Sseie6").Value = pitchinning / 3
            Else
                Range("Sseie6").Value = pitchinning \ 3 & "." & pitchinning Mod 3
            End If
            seban = CStr(Cells(143 + ((nowbatter - 1) * 4), 281).Value)
        End If
        With Worksheets("試合情報")
        If seikasan = 1 Then
            If kiroku <> "" And InStr(kiroku, "球") < 1 And InStr(kiroku, "遠") < 1 And _
               InStr(kiroku, "妨") < 1 Then
                If nowattack = 1 Then
                    Call minusinfo(Steam, seban, 5, 1)
                Else
                    Call minusinfo(Kteam, seban, 5, 1)
                End If
            End If
            If InStr(kiroku, "安") >= 1 Then
                If nowattack = 1 Then
                    Call minusinfo(Steam, seban, 6, 1)
                Else
                    Call minusinfo(Kteam, seban, 6, 1)
                End If
            End If
            If InStr(kiroku, "２") >= 1 Then
                If nowattack = 1 Then
                    Call minusinfo(Steam, seban, 6, 1)
                    Call minusinfo(Steam, seban, 7, 1)
                Else
                    Call minusinfo(Kteam, seban, 6, 1)
                    Call minusinfo(Kteam, seban, 7, 1)
                End If
            End If
            If InStr(kiroku, "３") >= 1 Then
                If nowattack = 1 Then
                    Call minusinfo(Steam, seban, 6, 1)
                    Call minusinfo(Steam, seban, 8, 1)
                Else
                    Call minusinfo(Kteam, seban, 6, 1)
                    Call minusinfo(Kteam, seban, 8, 1)
                End If
            End If
            If InStr(kiroku, "四") >= 1 Then
                If nowattack = 1 Then
                    Call minusinfo(Steam, seban, 13, 1)
                Else
                    Call minusinfo(Kteam, seban, 13, 1)
                End If
            End If
            If InStr(kiroku, "死") >= 1 Then
                If nowattack = 1 Then
                    Call minusinfo(Steam, seban, 14, 1)
                Else
                    Call minusinfo(Kteam, seban, 14, 1)
                End If
            End If
            If InStr(kiroku, "犠") >= 1 And .Range("A1").Value = "1" Then
                If nowattack = 1 Then
                    Call minusinfo(Steam, seban, 15, 1)
                Else
                    Call minusinfo(Kteam, seban, 15, 1)
                End If
            End If
            daten = .Range("H4").Value
            If nowattack = 1 Then
                Call minusinfo(Steam, seban, 10, daten)
            Else
                Call minusinfo(Kteam, seban, 10, daten)
            End If
        End If
        Call 成績表示
        Call 本日結果表示
        Call ScoreCellReset
        Call RunScore
            .Range("A1,F3:F8,H3:H8,J3:J5").Value = ""
        End With
        
    End If
End Sub

Sub 一塁()
    If ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38) Then
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(255, 0, 0)
    Else
        ActiveSheet.Shapes("一塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    End If
End Sub

Sub 二塁()
    If ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38) Then
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(255, 0, 0)
    Else
        ActiveSheet.Shapes("二塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    End If
End Sub

Sub 三塁()
    If ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38) Then
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(255, 0, 0)
    Else
        ActiveSheet.Shapes("三塁").Fill.ForeColor.RGB = RGB(38, 38, 38)
    End If
End Sub

Sub 設定取得()
    '標準
    Set basic_pos_Interior = Range("basicpos").Interior
    If Range("basicpos").Interior.ColorIndex = -4142 Then
        basic_pos_Interior.Color = -4142
    End If
    basic_pos_Font = Range("basicpos").Font.Color
    Set basic_pos_Borders = Range("basicpos").MergeArea.Borders
    
    Set basic_name_Interior = Range("basicname").Interior
    If Range("basicname").Interior.ColorIndex = -4142 Then
        basic_name_Interior.Color = -4142
    End If
    basic_name_Font = Range("basicname").Font.Color
    Set basic_name_Borders = Range("basicname").MergeArea.Borders
    
    'グラデーション
    If Range("basicpos").Interior.Pattern <> 1 Then
        basic_pos_Color1 = basic_pos_Interior.Gradient.ColorStops.Item(1).Color
        basic_pos_Color2 = basic_pos_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("basicname").Interior.Pattern <> 1 Then
        basic_name_Color1 = basic_name_Interior.Gradient.ColorStops.Item(1).Color
        basic_name_Color2 = basic_name_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '打者
    Set batter_pos_Interior = Range("batterpos").Interior
    If Range("batterpos").Interior.ColorIndex = -4142 Then
        batter_pos_Interior.Color = -4142
    End If
    batter_pos_Font = Range("batterpos").Font.Color
    Set batter_pos_Borders = Range("batterpos").MergeArea.Borders
    
    Set batter_name_Interior = Range("battername").Interior
    If Range("battername").Interior.ColorIndex = -4142 Then
        batter_name_Interior.Color = -4142
    End If
    batter_name_Font = Range("battername").Font.Color
    Set batter_name_Borders = Range("battername").MergeArea.Borders
    
    'グラデーション
    If Range("batterpos").Interior.Pattern <> 1 Then
        batter_pos_Color1 = batter_pos_Interior.Gradient.ColorStops.Item(1).Color
        batter_pos_Color2 = batter_pos_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("battername").Interior.Pattern <> 1 Then
        batter_name_Color1 = batter_name_Interior.Gradient.ColorStops.Item(1).Color
        batter_name_Color2 = batter_name_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '走者
    Set runner_pos_Interior = Range("runnerpos").Interior
    If Range("runnerpos").Interior.ColorIndex = -4142 Then
        runner_pos_Interior.Color = -4142
    End If
    runner_pos_Font = Range("runnerpos").Font.Color
    Set runner_pos_Borders = Range("runnerpos").MergeArea.Borders
    
    Set runner_name_Interior = Range("runnername").Interior
    If Range("runnername").Interior.ColorIndex = -4142 Then
        runner_name_Interior.Color = -4142
    End If
    runner_name_Font = Range("runnername").Font.Color
    Set runner_name_Borders = Range("runnername").MergeArea.Borders
    
    'グラデーション
    If Range("runnerpos").Interior.Pattern <> 1 Then
        runner_pos_Color1 = runner_pos_Interior.Gradient.ColorStops.Item(1).Color
        runner_pos_Color2 = runner_pos_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("runnername").Interior.Pattern <> 1 Then
        runner_name_Color1 = runner_name_Interior.Gradient.ColorStops.Item(1).Color
        runner_name_Color2 = runner_name_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '次回先頭
    Set next_pos_Interior = Range("nextpos").Interior
    If Range("nextpos").Interior.ColorIndex = -4142 Then
        next_pos_Interior.Color = -4142
    End If
    next_pos_Font = Range("nextpos").Font.Color
    Set next_pos_Borders = Range("nextpos").MergeArea.Borders
    
    Set next_name_Interior = Range("nextname").Interior
    If Range("nextname").Interior.ColorIndex = -4142 Then
        next_name_Interior.Color = -4142
    End If
    next_name_Font = Range("nextname").Font.Color
    Set next_name_Borders = Range("nextname").MergeArea.Borders
    
    'グラデーション
    If Range("nextpos").Interior.Pattern <> 1 Then
        next_pos_Color1 = next_pos_Interior.Gradient.ColorStops.Item(1).Color
        next_pos_Color2 = next_pos_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("nextname").Interior.Pattern <> 1 Then
        next_name_Color1 = next_name_Interior.Gradient.ColorStops.Item(1).Color
        next_name_Color2 = next_name_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '交代(消滅)
    Set dis_pos_Interior = Range("dispos").Interior
    If Range("dispos").Interior.ColorIndex = -4142 Then
        dis_pos_Interior.Color = -4142
    End If
    dis_pos_Font = Range("dispos").Font.Color
    Set dis_pos_Borders = Range("dispos").MergeArea.Borders
    
    Set dis_name_Interior = Range("disname").Interior
    If Range("disname").Interior.ColorIndex = -4142 Then
        dis_name_Interior.Color = -4142
    End If
    dis_name_Font = Range("disname").Font.Color
    Set dis_name_Borders = Range("disname").MergeArea.Borders
    
    'グラデーション
    If Range("dispos").Interior.Pattern <> 1 Then
        dis_pos_Color1 = dis_pos_Interior.Gradient.ColorStops.Item(1).Color
        dis_pos_Color2 = dis_pos_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("disname").Interior.Pattern <> 1 Then
        dis_name_Color1 = dis_name_Interior.Gradient.ColorStops.Item(1).Color
        dis_name_Color2 = dis_name_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '交代(点灯)
    Set app_pos_Interior = Range("apppos").Interior
    If Range("apppos").Interior.ColorIndex = -4142 Then
        app_pos_Interior.Color = -4142
    End If
    app_pos_Font = Range("apppos").Font.Color
    Set app_pos_Borders = Range("apppos").MergeArea.Borders
    
    Set app_name_Interior = Range("appname").Interior
    If Range("appname").Interior.ColorIndex = -4142 Then
        app_name_Interior.Color = -4142
    End If
    app_name_Font = Range("appname").Font.Color
    Set app_name_Borders = Range("appname").MergeArea.Borders
    
    'グラデーション
    If Range("apppos").Interior.Pattern <> 1 Then
        app_pos_Color1 = app_pos_Interior.Gradient.ColorStops.Item(1).Color
        app_pos_Color2 = app_pos_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("appname").Interior.Pattern <> 1 Then
        app_name_Color1 = app_name_Interior.Gradient.ColorStops.Item(1).Color
        app_name_Color2 = app_name_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '成績
    Set grade_item_Interior = Range("gradeitem").Interior
    If Range("gradeitem").Interior.ColorIndex = -4142 Then
        grade_item_Interior.Color = -4142
    End If
    grade_item_Font = Range("gradeitem").Font.Color
    Set grade_item_Borders = Range("gradeitem").MergeArea.Borders
    
    Set grade_entity_Interior = Range("gradeent").Interior
    If Range("gradeent").Interior.ColorIndex = -4142 Then
        grade_entity_Interior.Color = -4142
    End If
    grade_entity_Font = Range("gradeent").Font.Color
    Set grade_entity_Borders = Range("gradeent").MergeArea.Borders
    
    'グラデーション
    If Range("gradeitem").Interior.Pattern <> 1 Then
        grade_item_Color1 = grade_item_Interior.Gradient.ColorStops.Item(1).Color
        grade_item_Color2 = grade_item_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("gradeent").Interior.Pattern <> 1 Then
        grade_entity_Color1 = grade_entity_Interior.Gradient.ColorStops.Item(1).Color
        grade_entity_Color2 = grade_entity_Interior.Gradient.ColorStops.Item(2).Color
    End If
        
    'スコア
    Set score_inning_Interior = Range("scoreinin").Interior
    If Range("scoreinin").Interior.ColorIndex = -4142 Then
        score_inning_Interior.Color = -4142
    End If
    score_inning_Font = Range("scoreinin").Font.Color
    Set score_inning_Borders = Range("scoreinin").MergeArea.Borders
    
    Set score_entity_Interior = Range("scoreent").Interior
    If Range("scoreent").Interior.ColorIndex = -4142 Then
        score_entity_Interior.Color = -4142
    End If
    score_entity_Font = Range("scoreent").Font.Color
    Set score_entity_Borders = Range("scoreent").MergeArea.Borders
    
    'グラデーション
    If Range("scoreinin").Interior.Pattern <> 1 Then
        score_inning_Color1 = score_inning_Interior.Gradient.ColorStops.Item(1).Color
        score_inning_Color2 = score_inning_Interior.Gradient.ColorStops.Item(2).Color
    End If
    If Range("scoreent").Interior.Pattern <> 1 Then
        score_entity_Color1 = score_entity_Interior.Gradient.ColorStops.Item(1).Color
        score_entity_Color2 = score_entity_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '審判
    Set umpire_Interior = Range("umpire").Interior
    If Range("umpire").Interior.ColorIndex = -4142 Then
        umpire_Interior.Color = -4142
    End If
    umpire_Font = Range("umpire").Font.Color
    Set umpire_Borders = Range("umpire").MergeArea.Borders
    
    'グラデーション
    If Range("umpire").Interior.Pattern <> 1 Then
        umpire_Color1 = umpire_Interior.Gradient.ColorStops.Item(1).Color
        umpire_Color2 = umpire_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    'チーム名
    Set team_Interior = Range("teamname").Interior
    If Range("teamname").Interior.ColorIndex = -4142 Then
        team_Interior.Color = -4142
    End If
    team_Font = Range("teamname").Font.Color
    Set team_Borders = Range("teamname").MergeArea.Borders
    
    'グラデーション
    If Range("teamname").Interior.Pattern <> 1 Then
        team_Color1 = team_Interior.Gradient.ColorStops.Item(1).Color
        team_Color2 = team_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '安打
    Set result_hit_Interior = Range("resulthit").Interior
    If Range("resulthit").Interior.ColorIndex = -4142 Then
        result_hit_Interior.Color = -4142
    End If
    result_hit_Font = Range("resulthit").Font.Color
    Set result_hit_Borders = Range("resulthit").MergeArea.Borders
    
    'グラデーション
    If Range("resulthit").Interior.Pattern <> 1 Then
        result_hit_Color1 = result_hit_Interior.Gradient.ColorStops.Item(1).Color
        result_hit_Color2 = result_hit_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '安打
    Set result_hit_Interior = Range("resulthit").Interior
    If Range("resulthit").Interior.ColorIndex = -4142 Then
        result_hit_Interior.Color = -4142
    End If
    result_hit_Font = Range("resulthit").Font.Color
    Set result_hit_Borders = Range("resulthit").MergeArea.Borders
    
    'グラデーション
    If Range("resulthit").Interior.Pattern <> 1 Then
        result_hit_Color1 = result_hit_Interior.Gradient.ColorStops.Item(1).Color
        result_hit_Color2 = result_hit_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    '凡打
    Set result_out_Interior = Range("resultout").Interior
    If Range("resultout").Interior.ColorIndex = -4142 Then
        result_out_Interior.Color = -4142
    End If
    result_out_Font = Range("resultout").Font.Color
    Set result_out_Borders = Range("resultout").MergeArea.Borders
    
    'グラデーション
    If Range("resultout").Interior.Pattern <> 1 Then
        result_out_Color1 = result_out_Interior.Gradient.ColorStops.Item(1).Color
        result_out_Color2 = result_out_Interior.Gradient.ColorStops.Item(2).Color
    End If
    
    'その他
    Set result_other_Interior = Range("resultother").Interior
    If Range("resultother").Interior.ColorIndex = -4142 Then
        result_other_Interior.Color = -4142
    End If
    result_other_Font = Range("resultother").Font.Color
    Set result_other_Borders = Range("resultother").MergeArea.Borders
    
    'グラデーション
    If Range("resultother").Interior.Pattern <> 1 Then
        result_other_Color1 = result_other_Interior.Gradient.ColorStops.Item(1).Color
        result_other_Color2 = result_other_Interior.Gradient.ColorStops.Item(2).Color
    End If

    '背景色
    If Range("backcolor").Interior.ColorIndex = -4142 Then
        back_Interior = -4142
    Else
        back_Interior = Range("backcolor").Interior.Color
    End If
End Sub

Function returnrange(ByVal cellname As String, interior_property As Interior, font_property As Long, borders_property As Borders, color1 As Long, color2 As Long)
    If interior_property.Pattern = 1 Then
        Range(cellname).Interior.Color = interior_property.Color
    Else
        Range(cellname).Interior.Pattern = interior_property.Pattern
        Range(cellname).Interior.Gradient.Degree = interior_property.Gradient.Degree
        Range(cellname).Interior.Gradient.ColorStops.Clear
        Range(cellname).Interior.Gradient.ColorStops.Add(0).Color = color1
        Range(cellname).Interior.Gradient.ColorStops.Add(1).Color = color2
    End If
    Range(cellname).Font.Color = font_property
    If IsNull(borders_property.LineStyle) Then
        Range(cellname).MergeArea.Borders(xlEdgeLeft).Color = borders_property(xlEdgeLeft).Color
        Range(cellname).MergeArea.Borders(xlEdgeLeft).Weight = borders_property(xlEdgeLeft).Weight
        Range(cellname).MergeArea.Borders(xlEdgeLeft).LineStyle = borders_property(xlEdgeLeft).LineStyle
        Range(cellname).MergeArea.Borders(xlEdgeRight).Color = borders_property(xlEdgeRight).Color
        Range(cellname).MergeArea.Borders(xlEdgeRight).Weight = borders_property(xlEdgeRight).Weight
        Range(cellname).MergeArea.Borders(xlEdgeRight).LineStyle = borders_property(xlEdgeRight).LineStyle
        Range(cellname).MergeArea.Borders(xlEdgeTop).Color = borders_property(xlEdgeTop).Color
        Range(cellname).MergeArea.Borders(xlEdgeTop).Weight = borders_property(xlEdgeTop).Weight
        Range(cellname).MergeArea.Borders(xlEdgeTop).LineStyle = borders_property(xlEdgeTop).LineStyle
        Range(cellname).MergeArea.Borders(xlEdgeBottom).Color = borders_property(xlEdgeBottom).Color
        Range(cellname).MergeArea.Borders(xlEdgeBottom).Weight = borders_property(xlEdgeBottom).Weight
        Range(cellname).MergeArea.Borders(xlEdgeBottom).LineStyle = borders_property(xlEdgeBottom).LineStyle
    ElseIf returnflg = 1 Then
        If IsNull(Range(cellname).Borders.LineStyle) Or Range(cellname).Borders.LineStyle <> -4142 Then
            Range(cellname).Borders.Color = borders_property.Color
            Range(cellname).Borders.Weight = borders_property.Weight
            Range(cellname).Borders.LineStyle = borders_property.LineStyle
        End If
    ElseIf returnflg = 0 Then
        If borders_property.LineStyle <> -4142 Or Range(cellname).MergeArea.Borders.LineStyle <> -4142 _
        Or IsNull(Range(cellname).MergeArea.Borders.LineStyle) Then
            Range(cellname).MergeArea.Borders.Color = borders_property.Color
            Range(cellname).MergeArea.Borders.Weight = borders_property.Weight
            Range(cellname).MergeArea.Borders.LineStyle = borders_property.LineStyle
        End If
    End If
    returnflg = 0
End Function

Function returncells(ByVal cellrow As Integer, cellcol As Integer, interior_property As Interior, font_property As Long, borders_property As Borders, color1 As Long, color2 As Long)
    If interior_property.Pattern = 1 Then
        Cells(cellrow, cellcol).Interior.Color = interior_property.Color
    Else
        Cells(cellrow, cellcol).Interior.Pattern = interior_property.Pattern
        Cells(cellrow, cellcol).Interior.Gradient.Degree = interior_property.Gradient.Degree
        Cells(cellrow, cellcol).Interior.Gradient.ColorStops.Clear
        Cells(cellrow, cellcol).Interior.Gradient.ColorStops.Add(0).Color = color1
        Cells(cellrow, cellcol).Interior.Gradient.ColorStops.Add(1).Color = color2
    End If
    Cells(cellrow, cellcol).Font.Color = font_property
    If IsNull(borders_property.LineStyle) Then
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeLeft).Color = borders_property(xlEdgeLeft).Color
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeLeft).Weight = borders_property(xlEdgeLeft).Weight
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeLeft).LineStyle = borders_property(xlEdgeLeft).LineStyle
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeRight).Color = borders_property(xlEdgeRight).Color
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeRight).Weight = borders_property(xlEdgeRight).Weight
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeRight).LineStyle = borders_property(xlEdgeRight).LineStyle
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeTop).Color = borders_property(xlEdgeTop).Color
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeTop).Weight = borders_property(xlEdgeTop).Weight
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeTop).LineStyle = borders_property(xlEdgeTop).LineStyle
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeBottom).Color = borders_property(xlEdgeBottom).Color
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeBottom).Weight = borders_property(xlEdgeBottom).Weight
        Cells(cellrow, cellcol).MergeArea.Borders(xlEdgeBottom).LineStyle = borders_property(xlEdgeBottom).LineStyle
    ElseIf returnflg = 1 Then
        If IsNull(Cells(cellrow, cellcol).Borders.LineStyle) Or Cells(cellrow, cellcol).Borders.LineStyle <> -4142 Then
            Cells(cellrow, cellcol).Borders.Color = borders_property.Color
            Cells(cellrow, cellcol).Borders.Weight = borders_property.Weight
            Cells(cellrow, cellcol).Borders.LineStyle = borders_property.LineStyle
        End If
    ElseIf returnflg = 0 Then
        If borders_property.LineStyle <> -4142 Or Cells(cellrow, cellcol).MergeArea.Borders.LineStyle <> -4142 _
        Or IsNull(Cells(cellrow, cellcol).MergeArea.Borders.LineStyle) Then
            Cells(cellrow, cellcol).MergeArea.Borders.Color = borders_property.Color
            Cells(cellrow, cellcol).MergeArea.Borders.Weight = borders_property.Weight
            Cells(cellrow, cellcol).MergeArea.Borders.LineStyle = borders_property.LineStyle
        End If
    End If
    returnflg = 0
End Function

Function getinfo(team As String, seban As String)
    
    Dim arr(15)
    i = 4
    sw = 0
    With Worksheets(team)
        Maxrow = .Cells(Rows.count, 3).End(xlUp).Row
        Do While i <= Maxrow And sw = 0
            If CStr(.Cells(i, 2).Value) = CStr(seban) Then
                arr(0) = .Cells(i, 3).Value
                arr(1) = .Cells(i, 4).Value
                arr(2) = .Cells(i, 5).Value
                arr(3) = .Cells(i, 6).Value
                arr(4) = .Cells(i, 7).Value
                arr(5) = .Cells(i, 8).Value
                arr(6) = .Cells(i, 9).Value
                arr(7) = .Cells(i, 10).Value
                arr(8) = .Cells(i, 11).Value
                arr(9) = .Cells(i, 12).Value
                arr(10) = .Cells(i, 13).Value
                arr(11) = .Cells(i, 14).Value
                arr(12) = .Cells(i, 15).Value
                arr(13) = .Cells(i, 16).Value
                arr(14) = .Cells(i, 17).Value
                sw = 1
            End If
            i = i + 1
        Loop
    End With
    
    getinfo = arr()
    
End Function

Function updateinfo(team As String, seban As String, col As Integer, num As Integer)
    
    i = 4
    sw = 0
    With Worksheets(team)
        Maxrow = .Cells(Rows.count, 3).End(xlUp).Row
        Do While i <= Maxrow And sw = 0
            If CStr(.Cells(i, 2).Value) = CStr(seban) Then
                .Cells(i, col).Value = .Cells(i, col).Value + num
                sw = 1
            End If
            i = i + 1
        Loop
    End With

End Function

Function minusinfo(team As String, seban As String, col As Integer, num As Integer)
    
    i = 4
    sw = 0
    With Worksheets(team)
        Maxrow = .Cells(Rows.count, 3).End(xlUp).Row
        Do While i <= Maxrow And sw = 0
            If CStr(.Cells(i, 2).Value) = CStr(seban) Then
                .Cells(i, col).Value = .Cells(i, col).Value - num
                sw = 1
            End If
            i = i + 1
        Loop
    End With

End Function

Function returnmember(team As String, pos As String, seban As String, poscell As String, _
                            namecell As String)
                            
    Dim arr()
    arr() = getinfo(team, seban)
    Dim count As Integer
    i = 1
    changeflg = 0
    Dim nextlead As Integer
    
    If gameflg = 0 Then
    
        Call returnrange(poscell, basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                     basic_pos_Color1, basic_pos_Color2)
        Call returnrange(namecell, basic_name_Interior, basic_name_Font, basic_name_Borders, _
                     basic_name_Color1, basic_name_Color2)
        If team = "セ" Or team = "パ" Then
            If Len(arr(1)) = LenB(StrConv(arr(1), vbFromUnicode)) Then
                strcnt = Len(arr(1))
                Call Fonta
            Else
                strcnt = Len(arr(1))
                Call Font
            End If
            Range(namecell).Font.Name = fontname
            Range(poscell).Value = pos
            Range(namecell).Value = arr(1)
        Else
            If Len(arr(0)) = LenB(StrConv(arr(0), vbFromUnicode)) Then
                strcnt = Len(arr(0))
                Call Fonta
            Else
                Range(namecell).HorizontalAlignment = xlDistributed
                strcnt = Len(arr(0))
                Call Font
            End If
            Range(namecell).Font.Name = fontname
            Range(poscell).Value = pos
            Range(namecell).Value = arr(0)
        End If
    Else
        Call returnrange(poscell, dis_pos_Interior, dis_pos_Font, dis_pos_Borders, _
                         dis_pos_Color1, dis_pos_Color2)
        Call returnrange(namecell, dis_name_Interior, dis_name_Font, dis_name_Borders, _
                         dis_name_Color1, dis_name_Color2)
        Application.Wait [now()+"0:00:00.50"]
        If team = "セ" Or team = "パ" Then
            If Len(arr(1)) = LenB(StrConv(arr(1), vbFromUnicode)) Then
                strcnt = Len(arr(1))
                Call Fonta
            Else
                strcnt = Len(arr(1))
                Call Font
            End If
            Range(namecell).Font.Name = fontname
            Range(poscell).Value = pos
            Range(namecell).Value = arr(1)
        Else
            If Len(arr(0)) = LenB(StrConv(arr(0), vbFromUnicode)) Then
                strcnt = Len(arr(0))
                Call Fonta
            Else
                strcnt = Len(arr(0))
                Call Font
            End If
            Range(namecell).Font.Name = fontname
            Range(poscell).Value = pos
            Range(namecell).Value = arr(0)
        End If
        For i = 1 To 5
            If i Mod 2 <> 0 Then
                Call returnrange(poscell, app_pos_Interior, app_pos_Font, app_pos_Borders, _
                             app_pos_Color1, app_pos_Color2)
                Call returnrange(namecell, app_name_Interior, app_name_Font, app_name_Borders, _
                             app_name_Color1, app_name_Color2)
                Application.Wait [now()+"0:00:00.50"]
            Else
                Call returnrange(poscell, dis_pos_Interior, dis_pos_Font, dis_pos_Borders, _
                             dis_pos_Color1, dis_pos_Color2)
                Call returnrange(namecell, dis_name_Interior, dis_name_Font, dis_name_Borders, _
                             dis_name_Color1, dis_name_Color2)
                Application.Wait [now()+"0:00:00.50"]
            End If
        Next i
        
        If pos = "H" Then
            Call returnrange(poscell, batter_pos_Interior, batter_pos_Font, batter_pos_Borders, _
                         batter_pos_Color1, batter_pos_Color2)
            Call returnrange(namecell, batter_name_Interior, batter_name_Font, batter_name_Borders, _
                         batter_name_Color1, batter_name_Color2)
            Call 成績表示
            Range("BP117:CZ125").Interior.Color = back_Interior
            Range("BP117:DO131").Borders.LineStyle = -4142
            Range("Stoday1,Stoday2,Stoday3,Stoday4,Stoday5,Stoday6").Value = ""
            Range("GJ117:HT125").Interior.Color = back_Interior
            Range("GJ117:II131").Borders.LineStyle = -4142
            Range("Ktoday1,Ktoday2,Ktoday3,Ktoday4,Ktoday5,Ktoday6").Value = ""
        ElseIf pos = "R" Then
            Call returnrange(poscell, runner_pos_Interior, runner_pos_Font, runner_pos_Borders, _
                         runner_pos_Color1, runner_pos_Color2)
            Call returnrange(namecell, runner_name_Interior, runner_name_Font, runner_name_Borders, _
                         runner_name_Color1, runner_name_Color2)
        Else
            Call returnrange(poscell, basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
            Call returnrange(namecell, basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
            If nowattack = 1 Then
                nextlead = Worksheets("試合情報").Range("C8").Value
                Call returncells(23 + ((nextlead - 1) * 11), 249, next_pos_Interior, next_pos_Font, _
                                 next_pos_Borders, next_pos_Color1, next_pos_Color2)
                Call returncells(23 + ((nextlead - 1) * 11), 260, next_name_Interior, next_name_Font, _
                                 next_name_Borders, next_name_Color1, next_name_Color2)
            Else
                nextlead = Worksheets("試合情報").Range("B8").Value
                Call returncells(23 + ((nextlead - 1) * 11), 13, next_pos_Interior, next_pos_Font, _
                                 next_pos_Borders, next_pos_Color1, next_pos_Color2)
                Call returncells(23 + ((nextlead - 1) * 11), 24, next_name_Interior, next_name_Font, _
                                 next_name_Borders, next_name_Color1, next_name_Color2)
            End If
        End If
        If pos = "1" Then
            Call 投手成績リセット
        End If
        changeflg = 1
    End If
        
End Function

Function 打席結果リセット(col As Integer)
    i = 11
    With Worksheets("試合情報")
        Maxrow = .Cells(Rows.count, col).End(xlUp).Row
        For i = 11 To Maxrow
            .Cells(i, col).Value = ""
        Next i
    End With
End Function

Sub test()
    Dim arr()
    '先攻 - 1番
    If Target.Address = "text11" Then
        Exit Sub
    Else
        arr() = getinfo(Steam, Range("text11").Text)
        If gameflg = 0 Then
            Range("Spos1").Value = Range("text1").Value
            Range("Sname1").Value = arr(0)
            Call returnrange("Spos1", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                        basic_pos_Color1, basic_pos_Color2)
            Call returnrange("Sname1", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                        basic_name_Color1, basic_name_Color2)
        Else
            Call memberchange(Steam, Range("text1").Value, Range("text11").Value, "Spos1", "Sname1")
        End If
    End If
End Sub

Sub S1return()
    pos = Range("text1").Value
    seban = Range("text11").Value
    Call returnmember(Steam, pos, seban, "Spos1", "Sname1")
    If changeflg = 1 Then
        Call 打席結果リセット(1)
    End If
End Sub

Sub S2return()
    pos = Range("text2").Value
    seban = Range("text12").Value
    Call returnmember(Steam, pos, seban, "Spos2", "Sname2")
    If changeflg = 1 Then
        Call 打席結果リセット(2)
    End If
End Sub

Sub S3return()
    pos = Range("text3").Value
    seban = Range("text13").Value
    Call returnmember(Steam, pos, seban, "Spos3", "Sname3")
    If changeflg = 1 Then
        Call 打席結果リセット(3)
    End If
End Sub

Sub S4return()
    pos = Range("text4").Value
    seban = Range("text14").Value
    Call returnmember(Steam, pos, seban, "Spos4", "Sname4")
    If changeflg = 1 Then
        Call 打席結果リセット(4)
    End If
End Sub

Sub S5return()
    pos = Range("text5").Value
    seban = Range("text15").Value
    Call returnmember(Steam, pos, seban, "Spos5", "Sname5")
    If changeflg = 1 Then
        Call 打席結果リセット(5)
    End If
End Sub

Sub S6return()
    pos = Range("text6").Value
    seban = Range("text16").Value
    Call returnmember(Steam, pos, seban, "Spos6", "Sname6")
    If changeflg = 1 Then
        Call 打席結果リセット(6)
    End If
End Sub

Sub S7return()
    pos = Range("text7").Value
    seban = Range("text17").Value
    Call returnmember(Steam, pos, seban, "Spos7", "Sname7")
    If changeflg = 1 Then
        Call 打席結果リセット(7)
    End If
End Sub

Sub S8return()
    pos = Range("text8").Value
    seban = Range("text18").Value
    Call returnmember(Steam, pos, seban, "Spos8", "Sname8")
    If changeflg = 1 Then
        Call 打席結果リセット(8)
    End If
End Sub

Sub S9return()
    pos = Range("text9").Value
    seban = Range("text19").Value
    Call returnmember(Steam, pos, seban, "Spos9", "Sname9")
    If changeflg = 1 Then
        Call 打席結果リセット(9)
    End If
End Sub

Sub S10return()
    pos = Range("text10").Value
    seban = Range("text20").Value
    Call returnmember(Steam, pos, seban, "Spos10", "Sname10")
End Sub

Sub K1return()
    pos = Range("text21").Value
    seban = Range("text31").Value
    Call returnmember(Kteam, pos, seban, "Kpos1", "Kname1")
    If changeflg = 1 Then
        Call 打席結果リセット(10)
    End If
End Sub

Sub K2return()
    pos = Range("text22").Value
    seban = Range("text32").Value
    Call returnmember(Kteam, pos, seban, "Kpos2", "Kname2")
    If changeflg = 1 Then
        Call 打席結果リセット(11)
    End If
End Sub

Sub K3return()
    pos = Range("text23").Value
    seban = Range("text33").Value
    Call returnmember(Kteam, pos, seban, "Kpos3", "Kname3")
    If changeflg = 1 Then
        Call 打席結果リセット(12)
    End If
End Sub

Sub K4return()
    pos = Range("text24").Value
    seban = Range("text34").Value
    Call returnmember(Kteam, pos, seban, "Kpos4", "Kname4")
    If changeflg = 1 Then
        Call 打席結果リセット(13)
    End If
End Sub

Sub K5return()
    pos = Range("text25").Value
    seban = Range("text35").Value
    Call returnmember(Kteam, pos, seban, "Kpos5", "Kname5")
    If changeflg = 1 Then
        Call 打席結果リセット(14)
    End If
End Sub

Sub K6return()
    pos = Range("text26").Value
    seban = Range("text36").Value
    Call returnmember(Kteam, pos, seban, "Kpos6", "Kname6")
    If changeflg = 1 Then
        Call 打席結果リセット(15)
    End If
End Sub

Sub K7return()
    pos = Range("text27").Value
    seban = Range("text37").Value
    Call returnmember(Kteam, pos, seban, "Kpos7", "Kname7")
    If changeflg = 1 Then
        Call 打席結果リセット(16)
    End If
End Sub

Sub K8return()
    pos = Range("text28").Value
    seban = Range("text38").Value
    Call returnmember(Kteam, pos, seban, "Kpos8", "Kname8")
    If changeflg = 1 Then
        Call 打席結果リセット(17)
    End If
End Sub

Sub K9return()
    pos = Range("text29").Value
    seban = Range("text39").Value
    Call returnmember(Kteam, pos, seban, "Kpos9", "Kname9")
    If changeflg = 1 Then
        Call 打席結果リセット(18)
    End If
End Sub

Sub K10return()
    pos = Range("text30").Value
    seban = Range("text40").Value
    Call returnmember(Kteam, pos, seban, "Kpos10", "Kname10")
End Sub

Sub S1poschange()
    pos = Range("text1").Value
    seban = Range("text11").Value
    Call returnmember(Steam, pos, seban, "Spos1", "Sname1")
End Sub

Sub S2poschange()
    pos = Range("text2").Value
    seban = Range("text12").Value
    Call returnmember(Steam, pos, seban, "Spos2", "Sname2")
End Sub

Sub S3poschange()
    pos = Range("text3").Value
    seban = Range("text13").Value
    Call returnmember(Steam, pos, seban, "Spos3", "Sname3")
End Sub

Sub S4poschange()
    pos = Range("text4").Value
    seban = Range("text14").Value
    Call returnmember(Steam, pos, seban, "Spos4", "Sname4")
End Sub

Sub S5poschange()
    pos = Range("text5").Value
    seban = Range("text15").Value
    Call returnmember(Steam, pos, seban, "Spos5", "Sname5")
End Sub

Sub S6poschange()
    pos = Range("text6").Value
    seban = Range("text16").Value
    Call returnmember(Steam, pos, seban, "Spos6", "Sname6")
End Sub

Sub S7poschange()
    pos = Range("text7").Value
    seban = Range("text17").Value
    Call returnmember(Steam, pos, seban, "Spos7", "Sname7")
End Sub

Sub S8poschange()
    pos = Range("text8").Value
    seban = Range("text18").Value
    Call returnmember(Steam, pos, seban, "Spos8", "Sname8")
End Sub

Sub S9poschange()
    pos = Range("text9").Value
    seban = Range("text19").Value
    Call returnmember(Steam, pos, seban, "Spos9", "Sname9")
End Sub

Sub S10poschange()
    Range("Spos10").MergeArea.Borders.LineStyle = -4142
    Range("Sname10").MergeArea.Borders.LineStyle = -4142
    Range("Spos10,Sname10").Interior.Color = back_Interior
    Range("Spos10,Sname10").Value = ""
    Call returnrange("Spos9", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                     basic_pos_Color1, basic_pos_Color2)
    Call returnrange("Sname9", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                     basic_name_Color1, basic_name_Color2)
End Sub

Sub K1poschange()
    pos = Range("text21").Value
    seban = Range("text31").Value
    Call returnmember(Kteam, pos, seban, "Kpos1", "Kname1")
End Sub

Sub K2poschange()
    pos = Range("text22").Value
    seban = Range("text32").Value
    Call returnmember(Kteam, pos, seban, "Kpos2", "Kname2")
End Sub

Sub K3poschange()
    pos = Range("text23").Value
    seban = Range("text33").Value
    Call returnmember(Kteam, pos, seban, "Kpos3", "Kname3")
End Sub

Sub K4poschange()
    pos = Range("text24").Value
    seban = Range("text34").Value
    Call returnmember(Kteam, pos, seban, "Kpos4", "Kname4")
End Sub

Sub K5poschange()
    pos = Range("text25").Value
    seban = Range("text35").Value
    Call returnmember(Kteam, pos, seban, "Kpos5", "Kname5")
End Sub

Sub K6poschange()
    pos = Range("text26").Value
    seban = Range("text36").Value
    Call returnmember(Kteam, pos, seban, "Kpos6", "Kname6")
End Sub

Sub K7poschange()
    pos = Range("text27").Value
    seban = Range("text37").Value
    Call returnmember(Kteam, pos, seban, "Kpos7", "Kname7")
End Sub

Sub K8poschange()
    pos = Range("text28").Value
    seban = Range("text38").Value
    Call returnmember(Kteam, pos, seban, "Kpos8", "Kname8")
End Sub

Sub K9poschange()
    pos = Range("text29").Value
    seban = Range("text39").Value
    Call returnmember(Kteam, pos, seban, "Kpos9", "Kname9")
End Sub

Sub K10poschange()
    Range("Kpos10").MergeArea.Borders.LineStyle = -4142
    Range("Kname10").MergeArea.Borders.LineStyle = -4142
    Range("Kpos10,Kname10").Interior.Color = back_Interior
    Range("Kpos10,Kname10").Value = ""
    Call returnrange("Kpos9", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                     basic_pos_Color1, basic_pos_Color2)
    Call returnrange("Kname9", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                     basic_name_Color1, basic_name_Color2)
End Sub

Sub MemberCellReset()
    If IsNull(Range("basicpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("batterpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("runnerpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("nextpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("dispos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("apppos").MergeArea.Borders.LineStyle) _
    Or Range("basicpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("batterpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("runnerpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("nextpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("dispos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("apppos").MergeArea.Borders.LineStyle <> -4142 Then
        Call returnrange("Spos1", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos2", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos3", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos4", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos5", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos6", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos7", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos8", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Spos9", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
    Else
        returnflg = 1
        Call returnrange("Spos1,Spos2,Spos3,Spos4,Spos5,Spos6,Spos7,Spos8,Spos9", _
                         basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
    End If
    If IsNull(Range("basicname").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("battername").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("runnername").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("nextname").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("disname").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("appname").MergeArea.Borders.LineStyle) _
    Or Range("basicname").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("battername").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("runnername").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("nextname").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("disname").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("appname").MergeArea.Borders.LineStyle <> -4142 Then
        Call returnrange("Sname1", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname2", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname3", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname4", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname5", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname6", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname7", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname8", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Sname9", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
    Else
        returnflg = 1
        Call returnrange("Sname1,Sname2,Sname3,Sname4,Sname5,Sname6,Sname7,Sname8,Sname9", _
                         basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
    End If
    If IsNull(Range("basicpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("batterpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("runnerpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("nextpos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("dispos").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("apppos").MergeArea.Borders.LineStyle) _
    Or Range("basicpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("batterpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("runnerpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("nextpos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("dispos").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("apppos").MergeArea.Borders.LineStyle <> -4142 Then
        Call returnrange("Kpos1", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos2", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos3", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos4", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos5", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos6", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos7", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos8", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
        Call returnrange("Kpos9", basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
    Else
        returnflg = 1
        Call returnrange("Kpos1,Kpos2,Kpos3,Kpos4,Kpos5,Kpos6,Kpos7,Kpos8,Kpos9", _
                         basic_pos_Interior, basic_pos_Font, basic_pos_Borders, _
                         basic_pos_Color1, basic_pos_Color2)
    End If
    If IsNull(Range("basicname").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("battername").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("runnername").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("nextname").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("disname").MergeArea.Borders.LineStyle) _
    Or IsNull(Range("appname").MergeArea.Borders.LineStyle) _
    Or Range("basicname").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("battername").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("runnername").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("nextname").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("disname").MergeArea.Borders.LineStyle <> -4142 _
    Or Range("appname").MergeArea.Borders.LineStyle <> -4142 Then
        Call returnrange("Kname1", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname2", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname3", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname4", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname5", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname6", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname7", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname8", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
        Call returnrange("Kname9", basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
    Else
        returnflg = 1
        Call returnrange("Kname1,Kname2,Kname3,Kname4,Kname5,Kname6,Kname7,Kname8,Kname9", _
                         basic_name_Interior, basic_name_Font, basic_name_Borders, _
                         basic_name_Color1, basic_name_Color2)
    End If
End Sub

Sub ScoreCellReset()
    If IsNull(Range("scoreinin").MergeArea.Borders.LineStyle) Or IsNull(Range("scoreent").MergeArea.Borders.LineStyle) _
    Or Range("scoreinin").MergeArea.Borders.LineStyle <> -4142 Or Range("scoreent").MergeArea.Borders.LineStyle <> -4142 Then
        Call returnrange("Srun1", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun2", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun3", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun4", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun5", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun6", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun7", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun8", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srun9", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Srunr", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Shit", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Serror", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Slob", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun1", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun2", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun3", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun4", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun5", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun6", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun7", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun8", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krun9", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Krunr", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Khit", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Kerror", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
        Call returnrange("Klob", score_entity_Interior, score_entity_Font, score_entity_Borders, _
                         score_entity_Color1, score_entity_Color2)
    Else
        returnflg = 1
        Call returnrange("Srun1,Srun2,Srun3,Srun4,Srun5,Srun6,Srun7,Srun8,Srun9,Srunr,Shit,Serror,Slob", _
                         score_entity_Interior, score_entity_Font, score_entity_Borders, score_entity_Color1, score_entity_Color2)
        returnflg = 1
        Call returnrange("Krun1,Krun2,Krun3,Krun4,Krun5,Krun6,Krun7,Krun8,Krun9,Krunr,Khit,Kerror,Klob", _
                         score_entity_Interior, score_entity_Font, score_entity_Borders, score_entity_Color1, score_entity_Color2)
    End If
End Sub

Sub test_2()
    MsgBox (basic_pos_Borders.Color)
End Sub

Sub test_3()
    MsgBox (LenB(StrConv("大瀬良 大地", vbFromUnicode)))
End Sub

Sub test_4()
    MsgBox (Range("scoreent").Borders.LineStyle)
End Sub

Sub test_5()
    MsgBox (Range("basicpos").Borders.LineStyle)
End Sub

Sub test_6()
    MsgBox (Range("Spos1").MergeArea.Borders(xlEdgeTop).LineStyle)
    MsgBox (Range("Spos1").MergeArea.Borders(xlEdgeBottom).LineStyle)
    MsgBox (Range("Spos1").MergeArea.Borders(xlEdgeLeft).LineStyle)
    MsgBox (Range("Spos1").MergeArea.Borders(xlEdgeRight).LineStyle)
    MsgBox (Range("Sname1").MergeArea.Borders(xlEdgeTop).LineStyle)
    MsgBox (Range("Sname1").MergeArea.Borders(xlEdgeBottom).LineStyle)
    MsgBox (Range("Sname1").MergeArea.Borders(xlEdgeLeft).LineStyle)
    MsgBox (Range("Sname1").MergeArea.Borders(xlEdgeRight).LineStyle)
End Sub

Sub test_7()
    MsgBox (Range("batterpos").MergeArea.Borders.LineStyle)
    MsgBox (Range("battername").MergeArea.Borders.LineStyle)
End Sub

Sub test_8()
    MsgBox (batter_pos_Borders.LineStyle)
    MsgBox (batter_name_Borders.LineStyle)
End Sub

Sub test_9()
    MsgBox (Range("Spos1").Borders.LineStyle)
    MsgBox (Range("Sname1").Borders.LineStyle)
End Sub

Sub test_10()
    If Range("Spos3").Borders(xlEdgeRight).Color = Range("runnerpos").Borders(xlEdgeRight).Color Then
        MsgBox ("yeah")
    End If
End Sub

Sub test_11()
    If IsNull(Range("batterpos").MergeArea.Borders.LineStyle) Then
        MsgBox ("yeah")
    End If
End Sub

Sub test_12()
    MsgBox (Range("backcolor").Interior.ColorIndex)
End Sub

