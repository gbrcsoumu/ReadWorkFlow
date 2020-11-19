Imports Microsoft.VisualBasic.FileIO

Module Module1
    Public LoginID As String                                ' FileMakerのログインID
    Public LoginPassWord As String                          ' FileMakerのログインパスワード
    Public Const DataBaseName As String = "退勤管理test01"  ' FileMakerのデータベース名
    Public Const MemberNameTable As String = "職員一覧"     ' 職員名簿のテーブル名
    Public Const MemberLogTable As String = "出退勤記録"      ' 退勤管理のテーブル名
    Public Const DateLogTable As String = "出退勤一覧"      ' 出退勤一覧のテーブル名
    Public Const CodeTable1 As String = "出勤コード一覧"     ' 出退コード一覧のテーブル名
    Public Const CodeTable2 As String = "退勤コード一覧"     ' 退勤コード一覧のテーブル名
    Public Const GooutTable As String = "外出記録"          ' 外出記録のテーブル名
    Public Const ReturnTable As String = "戻り記録"          ' 戻り記録のテーブル名
    Public Const CardMasterKeyString = "GBRC 2020"          ' Felicaカードの暗号化のキー
    Public Const ClosePassWord As String = "exit"
    Public Const ListReadTime As String = "03：00：00"            ' 職員一覧を読み込む時刻

    Public Const Path1 As String = "C:\CSV\workflow1"
    Public Const Path2 As String = "C:\CSV\workflow2"
    Public Const Path3 As String = "C:\CSV\workflow3"

    Public Const outPath1 As String = "C:\CSV2\workflow1"
    Public Const outPath2 As String = "C:\CSV2\workflow2"
    Public Const outPath3 As String = "C:\CSV2\workflow3"
    Sub Main()

        If System.IO.Directory.Exists(outPath1) = False Then System.IO.Directory.CreateDirectory(outPath1)
        If System.IO.Directory.Exists(outPath2) = False Then System.IO.Directory.CreateDirectory(outPath2)
        If System.IO.Directory.Exists(outPath3) = False Then System.IO.Directory.CreateDirectory(outPath3)

        'Dim File1 As String() = ReadCSV(Path1)
        'Dim File2 As String() = ReadCSV(Path2)
        ''Dim File3 As String() = ReadCSV(Path3)

        If ReadWorkFlow(Path1) = True Then

        End If

    End Sub

    Function ReadWorkFlow(ByVal path As String) As Boolean

        ReadWorkFlow = False
        Dim WildCard1 As String
        'Dim Count As Integer = 0
        Dim ff() As String    ', flag() As Boolean

        WildCard1 = "*.csv"

        Dim nn As Integer = 0

        ff = System.IO.Directory.GetFiles(path, WildCard1, System.IO.SearchOption.AllDirectories)
        nn = ff.Length
        If nn > 0 Then
            Dim data As String()()
            For i As Integer = 0 To nn - 1
                data = ReadCsv(ff(i), False, False)
                If data.Length > 0 Then
                    If data.Length > 1 Then

                    End If

                    Dim file2 As String = outPath1 + "\" + System.IO.Path.GetFileName(ff(i))
                    System.IO.File.Move(ff(i), file2)
                End If
            Next
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CSVファイルの読込処理
    ''' </summary>
    ''' <param name="astrFileName">ファイル名
    ''' <param name="ablnTab">区切りの指定(True:タブ区切り, False:カンマ区切り)
    ''' <param name="ablnQuote">引用符フラグ(True:引用符で囲まれている, False:囲まれていない)
    ''' <returns>読込結果の文字列の2次元配列</returns>
    ''' -----------------------------------------------------------------------------
    Public Function ReadCsv(ByVal astrFileName As String,
                         ByVal ablnTab As Boolean,
                         ByVal ablnQuote As Boolean) As String()()
        ReadCsv = Nothing
        'ファイルStreamReader
        Dim parser As Microsoft.VisualBasic.FileIO.TextFieldParser = Nothing
        Try
            'Shift-JISエンコードで変換できない場合は「?」文字の設定
            Dim encFallBack As System.Text.DecoderReplacementFallback = New System.Text.DecoderReplacementFallback("?")
            Dim enc As System.Text.Encoding =
            System.Text.Encoding.GetEncoding("shift_jis", System.Text.EncoderFallback.ReplacementFallback, encFallBack)

            'TextFieldParserクラス
            parser = New Microsoft.VisualBasic.FileIO.TextFieldParser(astrFileName, enc)

            '区切りの指定
            parser.TextFieldType = FieldType.Delimited
            If ablnTab = False Then
                'カンマ区切り
                parser.SetDelimiters(",")
            Else
                'タブ区切り
                parser.SetDelimiters(vbTab)
            End If

            If ablnQuote = True Then
                'フィールドが引用符で囲まれているか
                parser.HasFieldsEnclosedInQuotes = True
            End If

            'フィールドの空白トリム設定
            parser.TrimWhiteSpace = False

            Dim strArr()() As String = Nothing
            Dim nLine As Integer = 0
            'ファイルの終端までループ
            While Not parser.EndOfData
                'フィールドを読込
                Dim strDataArr As String() = parser.ReadFields()

                '戻り値領域の拡張
                ReDim Preserve strArr(nLine)

                '退避
                strArr(nLine) = strDataArr
                nLine += 1
            End While

            '正常終了
            Return strArr

        Catch ex As Exception
            'エラー
            MsgBox(ex.Message)
        Finally
            '閉じる
            If parser IsNot Nothing Then
                parser.Close()
            End If
        End Try
    End Function

End Module
