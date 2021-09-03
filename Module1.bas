Attribute VB_Name = "Module1"
Option Explicit

'NE用データcsv作成
Public Sub GetdataNE()
 
    If Application.OperatingSystem Like "*Mac*" Then
      ' Mac 向けの処理
      Open ThisWorkbook.Path & "/NE.csv" For Output As #1
    Else
      ' Windows 向けの処理
      'ファイルを書き込みで開く(無ければ新規作成される、あれば上書き)
      Open Replace(ThisWorkbook.FullName, Dir(ThisWorkbook.FullName), "") & "NE.csv" For Output As #1
    End If
 
     Dim i As Long
 
    '開いたファイルに書き込む
    For i = Cells(Rows.Count, 1).End(xlUp).Row To 5 Step -1
        If Cells(i, 1).Value <> "" Then
            If Cells(i, 2).Value = "楽天" Then
                Print #1, Cells(i, 1).Value
            ElseIf Cells(i, 2).Value = "Yahoo" Then
                Print #1, Cells(i, 1).Value
            End If
        End If
    Next
 
    '開いたファイルを閉じる
    Close #1
 
    '終わったのが分かるようにメッセージを出す
    MsgBox "完了！"
 
End Sub



'ECCUBE用データcsv作成
Public Sub GetdataECCUBE()

    Const COL_DENPYONO As Integer = 7
    Const COL_OKUARISAKI_NAME As Integer = 9
    Const COL_OKUARISAKI_POST_CD As Integer = 10
    Const COL_OKUARISAKI_ADDRESS As Integer = 11
    Const COL_HAISO_DENPYONO As Integer = 12
    Const COL_SHUKKABI As Integer = 13
    Const COL_JUCHUNO As Integer = 15
    Const COL_HAISO_DENPYONO_M As Integer = 17
    Const COL_MIN_SHUKKA_DATE As Integer = 18
    Const COL_TEKIYOU As Integer = 20
    Const COL_CHUMONID As Integer = 22
    Const COL_HAISOID As Integer = 23
    Const COL_SHOHIN_CD As Integer = 24
    Const COL_HAISOSAKI_NAME1 As Integer = 25
    Const COL_HAISOSAKI_NAME2 As Integer = 26
    Const COL_HAISOSAKI_POST_CD As Integer = 27
    Const COL_HAISOSAKI_TODOFUKEN As Integer = 28
    Const COL_HAISOSAKI_JUSHO1 As Integer = 29
    Const COL_HAISOSAKI_JUSHO2 As Integer = 30
 
    Dim chumonId As String       'EC注文ID
    Dim juchuNo As String       'NE受注番号
    
    Dim csvdata As String        'CSVデータ
    Dim csvdataHaiso As String   'CSV出力用配送ID
    Dim csvdataDenpyo As String   'CSV出力用配送伝票番号
    Dim csvdataShukka As String   'CSV出力用出荷日
    
    Dim ws As Worksheet
    Dim wscsv As Worksheet
    
    Dim i, j As Long
    
    Set ws = ThisWorkbook.Worksheets("ECCUBE用参照データ")
    Set wscsv = ThisWorkbook.Worksheets("ECCUBE用データ抽出")
    
    'ECCUBEの注文ID(V列)があるところまでループ(V～AD列までのデータ参照にiを使ってます)
    For i = 2 To ws.Cells(Rows.Count, COL_CHUMONID).End(xlUp).Row Step 1
    
        chumonId = ws.Cells(i, COL_CHUMONID).Value  'ECCUBE注文ID(V列)
        
        'ECCUBEの注文ID(V列)の最終行になると終了(セルに数式入れてるので2万行回さないために)
        If chumonId = "" Then Exit For
       
        'NEの受注番号(O列)があるところまでループ(G～T列までのデータ参照にjを使ってます)
        For j = 2 To ws.Cells(Rows.Count, COL_JUCHUNO).End(xlUp).Row Step 1

            juchuNo = ws.Cells(j, COL_JUCHUNO).Value  'NE受注番号(O列)
            
            'NEの注文ID(O列)の最終行になると終了
            If juchuNo = "" Then Exit For
            
                'ECCUBEの値とサンリッチ参照値とが一致したら
                'csvデータ(配送ID(23)、配送伝票番号(12)、出荷日(13))を作成
                '注文IDと受注IDの一致
                '名前の一致
                '郵便番号の一致
                '住所の一致
                '出荷日データがあること
                
                '比較データ用変数
                Dim okurisaki_name, haisousaki_name As String
                Dim okurisaki_post_cd, haisousaki_post_cd, okurisaki_address, haisousaki_address As String
                Dim shukkabi As String
                
                okurisaki_name = ws.Cells(j, COL_OKUARISAKI_NAME).Value
                haisousaki_name = ws.Cells(i, COL_HAISOSAKI_NAME1).Value & ws.Cells(i, COL_HAISOSAKI_NAME2).Value
                okurisaki_post_cd = ws.Cells(j, COL_OKUARISAKI_POST_CD).Value
                haisousaki_post_cd = ws.Cells(i, COL_HAISOSAKI_POST_CD).Value
                okurisaki_address = ws.Cells(j, COL_OKUARISAKI_ADDRESS).Value
                haisousaki_address = ws.Cells(i, COL_HAISOSAKI_TODOFUKEN).Value & ws.Cells(i, COL_HAISOSAKI_JUSHO1).Value & ws.Cells(i, COL_HAISOSAKI_JUSHO2).Value
                shukkabi = Format(ws.Cells(j, COL_SHUKKABI).Value, "yyyy-mm-dd")
                
                If chumonId = juchuNo And _
                    okurisaki_name = haisousaki_name And _
                    okurisaki_post_cd = haisousaki_post_cd And _
                    okurisaki_address = haisousaki_address And _
                    Len(shukkabi) > 0 Then
                    
                    csvdataHaiso = ws.Cells(i, COL_HAISOID).Value
                    
                    If Len(csvdataHaiso) > 0 Then
                        csvdataDenpyo = ws.Cells(j, COL_HAISO_DENPYONO).Value
                        csvdata = csvdataHaiso & ",""" & csvdataDenpyo & """," & shukkabi
                        ws.Cells(i, 31).Value = csvdata
                        wscsv.Cells(i, 1).Value = csvdata
                    
                    End If
                    
                    
                End If
                
            Next

    Next
    
    '重複データの削除(空白データも削除）
    Dim lastRow As Long
    lastRow = wscsv.Range("A" & Rows.Count).End(xlUp).Row
    wscsv.Range("A1:A" & lastRow).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    wscsv.Range("A1:A" & lastRow).RemoveDuplicates (Array(1))
    
    
    If Application.OperatingSystem Like "*Mac*" Then
      ' Mac 向けの処理
      Open ThisWorkbook.Path & "/ECCUBE.csv" For Output As #2
    Else
      ' Windows 向けの処理
      'ファイルを書き込みで開く(無ければ新規作成される、あれば上書き)
      Open Replace(ThisWorkbook.FullName, Dir(ThisWorkbook.FullName), "") & "ECCUBE.csv" For Output As #2
    End If
    
    '開いたファイルに書き込む
    lastRow = wscsv.Range("A" & Rows.Count).End(xlUp).Row
    Print #2, "出荷ID,お問い合わせ番号,出荷日"    'ヘッダー書き込み
    For i = lastRow To 2 Step -1
        Print #2, Cells(i, 1).Value
    Next
    
    '開いたファイルを閉じる
    Close #2
 
    '終わったのが分かるようにメッセージを出す
    MsgBox "完了！"
 
End Sub
 




