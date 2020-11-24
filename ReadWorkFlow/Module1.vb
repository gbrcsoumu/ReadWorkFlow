Imports Microsoft.VisualBasic.FileIO

Module Module1
    Public LoginID As String                                    ' FileMakerのログインID
    Public LoginPassWord As String                              ' FileMakerのログインパスワード
    Public Const DataBaseName As String = "届出管理"            ' FileMakerのデータベース名
    Public Const HolidayTable As String = "休暇等届"            ' 休暇等届のテーブル名
    Public Const BussinessTripTable As String = "出張命令書"    ' 出張命令書のテーブル名
    Public Const HolidayWorkTable As String = "休日出勤命令書"  ' 休日出勤命令書のテーブル名

    Public Const Path1 As String = "C:\CSV\workflow1"
    Public Const Path2 As String = "C:\CSV\workflow2"
    Public Const Path3 As String = "C:\CSV\workflow3"

    Public Const outPath1 As String = "C:\CSV2\workflow1"
    Public Const outPath2 As String = "C:\CSV2\workflow2"
    Public Const outPath3 As String = "C:\CSV2\workflow3"

    Public Const delPath1 As String = "C:\CSV3\workflow1"
    Public Const delPath2 As String = "C:\CSV3\workflow2"
    Public Const delPath3 As String = "C:\CSV3\workflow3"

    Public Const Version As String = "Ver 1.00"

    Sub Main()

        If System.IO.Directory.Exists(outPath1) = False Then System.IO.Directory.CreateDirectory(outPath1)
        If System.IO.Directory.Exists(outPath2) = False Then System.IO.Directory.CreateDirectory(outPath2)
        If System.IO.Directory.Exists(outPath3) = False Then System.IO.Directory.CreateDirectory(outPath3)

        If System.IO.Directory.Exists(delPath1) = False Then System.IO.Directory.CreateDirectory(delPath1)
        If System.IO.Directory.Exists(delPath2) = False Then System.IO.Directory.CreateDirectory(delPath2)
        If System.IO.Directory.Exists(delPath3) = False Then System.IO.Directory.CreateDirectory(delPath3)

        'Dim File1 As String() = ReadCSV(Path1)
        'Dim File2 As String() = ReadCSV(Path2)
        ''Dim File3 As String() = ReadCSV(Path3)

        If ReadWorkFlow1(Path1) = True Then

        End If

    End Sub

    Function ReadWorkFlow1(ByVal path As String) As Boolean
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String

        ReadWorkFlow1 = False
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

                        db.Connect()


                        For j As Integer = 1 To data.Length - 1
                            Dim aa As String = "", bb As String = ""

                            Dim No As String
                            No = data(j)(21)
                            aa += """職員番号"""
                            bb += "'" + No + "'"
                            aa += ","
                            bb += ","

                            Dim Name As String
                            Name = data(j)(5)
                            aa += """職員名"""
                            bb += "'" + Name + "'"
                            aa += ","
                            bb += ","

                            Dim DateTime1 As String
                            DateTime1 = data(j)(7).Replace("/", "-") + ":00"
                            aa += """申請日"""
                            bb += "TIMESTAMP '" + DateTime1 + "'"
                            aa += ","
                            bb += ","

                            Dim Cat As String
                            Cat = data(j)(11).Replace("（変更前の日付を備考に記載）", "")
                            aa += """申請区分"""
                            bb += "'" + Cat + "'"
                            aa += ","
                            bb += ","

                            Dim Kind1 As String, Kind2 As String, Kind As String
                            Kind1 = data(j)(12).Replace("（備考欄に詳細記載）", "").Replace("/", vbCrLf)
                            Kind2 = data(j)(13).Replace("（備考欄に詳細記載）", "").Replace("/", vbCrLf)
                            If Kind1 <> "" Then
                                Kind = Kind1
                            Else
                                Kind = Kind2
                            End If
                            aa += """種類"""
                            bb += "'" + Kind + "'"
                            aa += ","
                            bb += ","

                            Dim StDate As String, EdDate As String
                            Dim StTime As String, EdTime As String
                            StDate = data(j)(14).Replace("/", "-")
                            aa += """開始日"""
                            bb += "DATE '" + StDate + "'"
                            aa += ","
                            bb += ","

                            StTime = data(j)(15) + ":00"
                            aa += """開始時間"""
                            bb += "TIME '" + StTime + "'"
                            aa += ","
                            bb += ","

                            EdDate = data(j)(17).Replace("/", "-")
                            aa += """終了日"""
                            bb += "DATE '" + EdDate + "'"
                            aa += ","
                            bb += ","

                            EdTime = data(j)(16) + ":00"
                            aa += """終了時間"""
                            bb += "TIME '" + EdTime + "'"
                            aa += ","
                            bb += ","

                            Dim DayCount As String
                            DayCount = data(j)(18)
                            If DayCount = "" Then DayCount = "0"
                            aa += """今回休暇日数"""
                            bb += DayCount
                            aa += ","
                            bb += ","

                            Dim TotalDayCount As String
                            TotalDayCount = data(j)(19)
                            If TotalDayCount = "" Then TotalDayCount = "0"
                            aa += """有給休暇累計"""
                            bb += TotalDayCount
                            aa += ","
                            bb += ","

                            Dim ReMarks As String
                            ReMarks = data(j)(20).Replace("'", "''")
                            aa += """備考"""
                            bb += "'" + ReMarks + "'"
                            aa += ","
                            bb += ","

                            'Sql_Command = "INSERT INTO """ + HolidayTable + """"
                            'Sql_Command += " (""職員番号"", ""職員名"", ""申請日"", ""申請区分"", ""種類"", ""開始日"", ""開始時間"", ""終了日"", ""終了時間"", ""今回休暇日数"", ""有給休暇累計"", ""備考"", ""入力"")"
                            'Sql_Command += " VALUES ('" + No + "','" + Name + "',TIMESTAMP '" + DateTime1 + "','" + Cat + "','" + Kind + "',DATE '" + StDate + "',TIME '" + StTime + "',DATE '" + EdDate + "',TIME '" + EdTime + "'"
                            'Sql_Command += "," + DayCount + "," + TotalDayCount + ",'" + ReMarks + "','未入力')"


                            aa += """入力"""
                            bb += "'未入力'"
                            aa += ","
                            bb += ","

                            aa += """バージョン"""
                            bb += "'" + Version + "'"
                            'aa += ","
                            'bb += ","

                            Sql_Command = "INSERT INTO """ + HolidayTable + """ (" + aa + ") VALUES (" + bb + ")"

                            tb = db.ExecuteSql(Sql_Command)


                        Next

                        db.Disconnect()

                        Dim file2 As String = outPath1 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)
                    Else
                        Dim file2 As String = delPath1 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)

                    End If


                End If
            Next
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' CSVファイルの読込処理
    ''' param name="astrFileName">ファイル名
    ''' param name="ablnTab"区切りの指定(True:タブ区切り, False:カンマ区切り)
    ''' param name="ablnQuote"引用符フラグ(True:引用符で囲まれている, False:囲まれていない)
    ''' return読込結果の文字列の2次元配列returns
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
