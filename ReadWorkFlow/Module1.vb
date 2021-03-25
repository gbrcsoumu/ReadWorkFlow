Imports Microsoft.VisualBasic.FileIO

Module Module1
    Public LoginID As String                                    ' FileMakerのログインID
    Public LoginPassWord As String                              ' FileMakerのログインパスワード
    Public Const DataBaseName As String = "届出管理"            ' FileMakerのデータベース名
    Public Const HolidayTable As String = "休暇等届"            ' 休暇等届のテーブル名
    Public Const BussinessTripTable As String = "出張命令書"    ' 出張命令書のテーブル名
    Public Const HolidayWorkTable As String = "休日出勤命令書"  ' 休日出勤命令書のテーブル名
    Public Const HomeWorkTable As String = "在宅勤務許可申請"   ' 在宅勤務許可申請のテーブル名

    Public Const Path1 As String = "C:\CSV\workflow1"
    Public Const Path2 As String = "C:\CSV\workflow2"
    Public Const Path3 As String = "C:\CSV\workflow3"
    Public Const Path4 As String = "C:\CSV\workflow4"

    Public Const outPath1 As String = "C:\CSV2\workflow1"
    Public Const outPath2 As String = "C:\CSV2\workflow2"
    Public Const outPath3 As String = "C:\CSV2\workflow3"
    Public Const outPath4 As String = "C:\CSV2\workflow4"

    Public Const delPath1 As String = "C:\CSV3\workflow1"
    Public Const delPath2 As String = "C:\CSV3\workflow2"
    Public Const delPath3 As String = "C:\CSV3\workflow3"
    Public Const delPath4 As String = "C:\CSV3\workflow4"

    Public Const Version As String = "Ver 1.00"

    Sub Main()

        If System.IO.Directory.Exists(outPath1) = False Then System.IO.Directory.CreateDirectory(outPath1)
        If System.IO.Directory.Exists(outPath2) = False Then System.IO.Directory.CreateDirectory(outPath2)
        If System.IO.Directory.Exists(outPath3) = False Then System.IO.Directory.CreateDirectory(outPath3)
        If System.IO.Directory.Exists(outPath4) = False Then System.IO.Directory.CreateDirectory(outPath4)

        If System.IO.Directory.Exists(delPath1) = False Then System.IO.Directory.CreateDirectory(delPath1)
        If System.IO.Directory.Exists(delPath2) = False Then System.IO.Directory.CreateDirectory(delPath2)
        If System.IO.Directory.Exists(delPath3) = False Then System.IO.Directory.CreateDirectory(delPath3)
        If System.IO.Directory.Exists(delPath4) = False Then System.IO.Directory.CreateDirectory(delPath4)

        'Dim File1 As String() = ReadCSV(Path1)
        'Dim File2 As String() = ReadCSV(Path2)
        ''Dim File3 As String() = ReadCSV(Path3)

        If ReadWorkFlow1(Path1) = 0 Then

        End If
        If ReadWorkFlow2(Path2) = 0 Then

        End If
        If ReadWorkFlow3(Path3) = 0 Then

        End If
        If ReadWorkFlow4(Path4) = 0 Then

        End If

    End Sub

    Function ReadWorkFlow1(ByVal path As String) As Integer
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String

        ReadWorkFlow1 = -1
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

                        'Sql_Command2 = "SELECT * FROM """ + DateLogTable + """ WHERE (""職員番号"" = '" & value & "' AND ""日付"" = DATE '" + D1 + " ')"
                        'tb2 = db.ExecuteSql(Sql_Command2)

                        For j As Integer = 1 To data.Length - 1
                            Dim aa As String = "", bb As String = ""

                            Dim No As String = data(j)(21)
                            aa += """職員番号"""
                            bb += "'" + No + "'"
                            aa += ","
                            bb += ","

                            Dim Name As String = data(j)(5)
                            aa += """職員名"""
                            bb += "'" + Name + "'"
                            aa += ","
                            bb += ","

                            Dim DateTime1 As String = data(j)(7).Replace("/", "-") + ":00"
                            aa += """申請日"""
                            bb += "TIMESTAMP '" + DateTime1 + "'"
                            aa += ","
                            bb += ","

                            Dim Cat As String = data(j)(11).Replace("（変更前の日付を備考に記載）", "")
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

                            If data(j)(14) <> "" Then
                                StDate = data(j)(14).Replace("/", "-")
                                aa += """開始日"""
                                bb += "DATE '" + StDate + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(15) <> "" Then
                                StTime = data(j)(15) + ":00"
                                aa += """開始時間"""
                                bb += "TIME '" + StTime + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(17) <> "" Then
                                EdDate = data(j)(17).Replace("/", "-")
                                aa += """終了日"""
                                bb += "DATE '" + EdDate + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(16) <> "" Then
                                EdTime = data(j)(16) + ":00"
                                aa += """終了時間"""
                                bb += "TIME '" + EdTime + "'"
                                aa += ","
                                bb += ","
                            End If

                            Dim DayCount As String = data(j)(18)
                            If DayCount = "" Then DayCount = "0"
                            aa += """今回休暇日数"""
                            bb += DayCount
                            aa += ","
                            bb += ","

                            Dim TotalDayCount As String = data(j)(19)
                            If TotalDayCount = "" Then TotalDayCount = "0"
                            aa += """有給休暇累計"""
                            bb += TotalDayCount
                            aa += ","
                            bb += ","

                            Dim TimeCount As String = data(j)(24)
                            aa += """今回休暇時間"""
                            bb += "'" + TimeCount + "'"
                            aa += ","
                            bb += ","

                            Dim ReMarks As String = data(j)(20).Replace("'", "''")
                            aa += """備考"""
                            bb += "'" + ReMarks + "'"
                            aa += ","
                            bb += ","

                            aa += """処理"""
                            bb += "'未処理'"
                            aa += ","
                            bb += ","

                            aa += """バージョン"""
                            bb += "'" + Version + "'"
                            'aa += ","
                            'bb += ","

                            Dim Sql_Command2 As String = "SELECT ""職員番号"" FROM """ + HolidayTable +
                                """ WHERE (""職員番号"" = '" + No + "' AND ""申請日"" = TIMESTAMP '" + DateTime1 + "' AND ""開始日"" = DATE '" + StDate + " ')"
                            Dim tb2 As DataTable = db.ExecuteSql(Sql_Command2)
                            Dim n2 As Integer = tb2.Rows.Count

                            If n2 = 0 Then
                                Sql_Command = "INSERT INTO """ + HolidayTable + """ (" + aa + ") VALUES (" + bb + ")"
                                tb = db.ExecuteSql(Sql_Command)
                            End If
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
            ReadWorkFlow1 = 0

        End If

    End Function


    Function ReadWorkFlow2(ByVal path As String) As Integer
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String

        ReadWorkFlow2 = -1
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

                            Dim No As String = data(j)(11)
                            aa += """職員番号"""
                            bb += "'" + No + "'"
                            aa += ","
                            bb += ","

                            Dim Name As String = data(j)(5)
                            aa += """職員名"""
                            bb += "'" + Name + "'"
                            aa += ","
                            bb += ","

                            Dim DateTime1 As String = data(j)(7).Replace("/", "-") + ":00"
                            aa += """申請日"""
                            bb += "TIMESTAMP '" + DateTime1 + "'"
                            aa += ","
                            bb += ","

                            Dim Cat As String = data(j)(12).Replace("（変更・中止：内容等を備考欄に入力）", "")
                            aa += """申請区分"""
                            bb += "'" + Cat + "'"
                            aa += ","
                            bb += ","



                            Dim StDate As String, EdDate As String
                            Dim StTime As String, EdTime As String

                            If data(j)(15) <> "" Then
                                StDate = data(j)(15).Replace("/", "-")
                                aa += """出張開始日"""
                                bb += "DATE '" + StDate + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(16) <> "" Then
                                StTime = data(j)(16) + ":00"
                                aa += """出発時間"""
                                bb += "TIME '" + StTime + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(17) <> "" Then
                                EdDate = data(j)(17).Replace("/", "-")
                                aa += """出張終了日"""
                                bb += "DATE '" + EdDate + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(18) <> "" Then
                                EdTime = data(j)(18) + ":00"
                                aa += """帰着時間"""
                                bb += "TIME '" + EdTime + "'"
                                aa += ","
                                bb += ","
                            End If

                            Dim Zenpaku As String = data(j)(19)
                            aa += """前泊・後泊"""
                            bb += "'" + Zenpaku + "'"
                            aa += ","
                            bb += ","


                            Dim Dest1 As String = data(j)(21)
                            aa += """行先＿所内"""
                            bb += "'" + Dest1 + "'"
                            aa += ","
                            bb += ","

                            Dim Dest2 As String = data(j)(22)
                            aa += """行先＿所外"""
                            bb += "'" + Dest2 + "'"
                            aa += ","
                            bb += ","

                            Dim Address As String = data(j)(23)
                            aa += """行先＿所在地"""
                            bb += "'" + Address + "'"
                            aa += ","
                            bb += ","

                            Dim Method As String = data(j)(24)
                            aa += """移動手段"""
                            bb += "'" + Method + "'"
                            aa += ","
                            bb += ","


                            Dim Job As String = data(j)(20)
                            aa += """用務"""
                            bb += "'" + Job + "'"
                            aa += ","
                            bb += ","

                            Dim ReMarks As String = data(j)(28).Replace("'", "''")
                            aa += """備考"""
                            bb += "'" + ReMarks + "'"
                            aa += ","
                            bb += ","

                            Dim Member As String = data(j)(31)
                            aa += """同行者の有無"""
                            bb += "'" + Member + "'"
                            aa += ","
                            bb += ","

                            Dim No1 As String = data(j)(14)
                            aa += """同行者職員番号1"""
                            bb += "'" + No1 + "'"
                            aa += ","
                            bb += ","

                            Dim Name1 As String = data(j)(13)
                            aa += """同行者氏名1"""
                            bb += "'" + Name1 + "'"
                            aa += ","
                            bb += ","

                            Dim No2 As String = data(j)(32)
                            aa += """同行者職員番号2"""
                            bb += "'" + No2 + "'"
                            aa += ","
                            bb += ","

                            Dim Name2 As String = data(j)(34)
                            aa += """同行者氏名2"""
                            bb += "'" + Name2 + "'"
                            aa += ","
                            bb += ","

                            Dim No3 As String = data(j)(33)
                            aa += """同行者職員番号3"""
                            bb += "'" + No3 + "'"
                            aa += ","
                            bb += ","

                            Dim Name3 As String = data(j)(35)
                            aa += """同行者氏名3"""
                            bb += "'" + Name3 + "'"
                            aa += ","
                            bb += ","


                            Dim CostExist As String = data(j)(30)
                            aa += """費用の有無"""
                            bb += "'" + CostExist + "'"
                            aa += ","
                            bb += ","

                            Dim CostTerm As String = data(j)(27)
                            aa += """費用の内容"""
                            bb += "'" + CostTerm + "'"
                            aa += ","
                            bb += ","

                            Dim Cost As String = data(j)(25)
                            If Cost = "" Then Cost = "0"
                            aa += """費用"""
                            bb += Cost
                            aa += ","
                            bb += ","

                            aa += """処理"""
                            bb += "'未処理'"
                            aa += ","
                            bb += ","

                            aa += """バージョン"""
                            bb += "'" + Version + "'"
                            'aa += ","
                            'bb += ","

                            Dim Sql_Command2 As String = "SELECT ""職員番号"" FROM """ + BussinessTripTable +
                                """ WHERE (""職員番号"" = '" + No + "' AND ""申請日"" = TIMESTAMP '" + DateTime1 + "' AND ""出張開始日"" = DATE '" + StDate + " ')"
                            Dim tb2 As DataTable = db.ExecuteSql(Sql_Command2)
                            Dim n2 As Integer = tb2.Rows.Count

                            If n2 = 0 Then
                                Sql_Command = "INSERT INTO """ + BussinessTripTable + """ (" + aa + ") VALUES (" + bb + ")"

                                tb = db.ExecuteSql(Sql_Command)
                            End If

                        Next

                        db.Disconnect()

                        Dim file2 As String = outPath2 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)
                    Else
                        Dim file2 As String = delPath2 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)

                    End If


                End If
            Next
            ReadWorkFlow2 = 0
        End If

    End Function


    Function ReadWorkFlow3(ByVal path As String) As Integer
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String

        ReadWorkFlow3 = -1
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
                            Dim WorkDate1 As String

                            Dim No As String = data(j)(11)
                            aa += """職員番号"""
                            bb += "'" + No + "'"
                            aa += ","
                            bb += ","

                            Dim Name As String = data(j)(5)
                            aa += """職員名"""
                            bb += "'" + Name + "'"
                            aa += ","
                            bb += ","

                            Dim DateTime1 As String = data(j)(7).Replace("/", "-") + ":00"
                            aa += """申請日"""
                            bb += "TIMESTAMP '" + DateTime1 + "'"
                            aa += ","
                            bb += ","

                            Dim Cat As String = data(j)(12).Replace("（変更・中止：理由を備考欄に入力）", "")
                            aa += """申請区分"""
                            bb += "'" + Cat + "'"
                            aa += ","
                            bb += ","

                            If data(j)(13) <> "" Then
                                WorkDate1 = data(j)(13).Replace("/", "-")
                                aa += """休日出勤日1"""
                                bb += "DATE '" + WorkDate1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(14) <> "" Then
                                Dim StTime1 As String = data(j)(14) + ":00"
                                aa += """開始時間1"""
                                bb += "TIME '" + StTime1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(15) <> "" Then
                                Dim EdTime1 As String = data(j)(15) + ":00"
                                aa += """終了時間1"""
                                bb += "TIME '" + EdTime1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(16) <> "" Then
                                Dim SubDate1 As String = data(j)(16).Replace("/", "-")
                                aa += """振替休日1"""
                                bb += "DATE '" + SubDate1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            Dim DayLength1 As String = data(j)(36)
                            aa += """振替休日の長さ1"""
                            bb += "'" + DayLength1 + "'"
                            aa += ","
                            bb += ","

                            If data(j)(17) <> "" Then
                                Dim WorkDate2 As String = data(j)(17).Replace("/", "-")
                                aa += """休日出勤日2"""
                                bb += "DATE '" + WorkDate2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(18) <> "" Then
                                Dim StTime2 As String = data(j)(18) + ":00"
                                aa += """開始時間2"""
                                bb += "TIME '" + StTime2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(19) <> "" Then
                                Dim EdTime2 As String = data(j)(19) + ":00"
                                aa += """終了時間2"""
                                bb += "TIME '" + EdTime2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(20) <> "" Then
                                Dim SubDate2 As String = data(j)(20).Replace("/", "-")
                                aa += """振替休日2"""
                                bb += "DATE '" + SubDate2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            Dim DayLength2 As String = data(j)(37)
                            aa += """振替休日の長さ2"""
                            bb += "'" + DayLength2 + "'"
                            aa += ","
                            bb += ","

                            Dim Job As String = data(j)(21)
                            aa += """用務"""
                            bb += "'" + Job + "'"
                            aa += ","
                            bb += ","

                            Dim TriFlag As String = data(j)(22)
                            aa += """出張の有無"""
                            bb += "'" + TriFlag + "'"
                            aa += ","
                            bb += ","

                            Dim Zenpaku As String = data(j)(33)
                            aa += """前泊・後泊"""
                            bb += "'" + Zenpaku + "'"
                            aa += ","
                            bb += ","

                            Dim Dest1 As String = data(j)(23)
                            aa += """行先＿所内"""
                            bb += "'" + Dest1 + "'"
                            aa += ","
                            bb += ","

                            Dim Dest2 As String = data(j)(24)
                            aa += """行先＿所外"""
                            bb += "'" + Dest2 + "'"
                            aa += ","
                            bb += ","

                            Dim Address As String = data(j)(39)
                            aa += """所外所在地"""
                            bb += "'" + Address + "'"
                            aa += ","
                            bb += ","

                            Dim Method As String = data(j)(38)
                            aa += """移動手段"""
                            bb += "'" + Method + "'"
                            aa += ","
                            bb += ","

                            Dim TripStDate1 As String, TripEdDate1 As String
                            Dim TripStTime1 As String, TripEdTime1 As String
                            Dim TripStDate2 As String, TripEdDate2 As String
                            Dim TripStTime2 As String, TripEdTime2 As String

                            If data(j)(25) <> "" Then
                                TripStDate1 = data(j)(25).Replace("/", "-")
                                aa += """開始日1"""
                                bb += "DATE '" + TripStDate1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(29) <> "" Then
                                TripStTime1 = data(j)(29) + ":00"
                                aa += """出発時間1"""
                                bb += "TIME '" + TripStTime1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(26) <> "" Then
                                TripEdDate1 = data(j)(26).Replace("/", "-")
                                aa += """終了日1"""
                                bb += "DATE '" + TripEdDate1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(30) <> "" Then
                                TripEdTime1 = data(j)(30) + ":00"
                                aa += """帰着時間1"""
                                bb += "TIME '" + TripEdTime1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(27) <> "" Then
                                TripStDate2 = data(j)(27).Replace("/", "-")
                                aa += """開始日2"""
                                bb += "DATE '" + TripStDate2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(31) <> "" Then
                                TripStTime2 = data(j)(31) + ":00"
                                aa += """出発時間2"""
                                bb += "TIME '" + TripStTime2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(28) <> "" Then
                                TripEdDate2 = data(j)(28).Replace("/", "-")
                                aa += """終了日2"""
                                bb += "DATE '" + TripEdDate2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(32) <> "" Then
                                TripEdTime2 = data(j)(32) + ":00"
                                aa += """帰着時間2"""
                                bb += "TIME '" + TripEdTime2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            Dim ReMarks As String = data(j)(34).Replace("'", "''")
                            aa += """備考"""
                            bb += "'" + ReMarks + "'"
                            aa += ","
                            bb += ","

                            aa += """処理"""
                            bb += "'未処理'"
                            aa += ","
                            bb += ","

                            aa += """バージョン"""
                            bb += "'" + Version + "'"
                            'aa += ","
                            'bb += ","

                            Dim Sql_Command2 As String = "SELECT ""職員番号"" FROM """ + HolidayWorkTable +
                                """ WHERE (""職員番号"" = '" + No + "' AND ""申請日"" = TIMESTAMP '" + DateTime1 + "' AND ""休日出勤日1"" = DATE '" + WorkDate1 + " ')"

                            'Dim Sql_Command2 As String = "SELECT ""職員番号"" FROM """ + HolidayWorkTable + """ WHERE (""職員番号"" = '" & No & "' AND ""申請日"" = TIMESTAMP '" + DateTime1 + " ')"
                            Dim tb2 As DataTable = db.ExecuteSql(Sql_Command2)
                            Dim n2 As Integer = tb2.Rows.Count

                            If n2 = 0 Then
                                Sql_Command = "INSERT INTO """ + HolidayWorkTable + """ (" + aa + ") VALUES (" + bb + ")"

                                tb = db.ExecuteSql(Sql_Command)
                            End If

                        Next

                        db.Disconnect()

                        Dim file2 As String = outPath3 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)
                    Else
                        Dim file2 As String = delPath3 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)

                    End If


                End If
            Next
            ReadWorkFlow3 = 0
        End If

    End Function

    Function ReadWorkFlow4(ByVal path As String) As Integer
        Dim db As New OdbcDbIf
        Dim tb As DataTable
        Dim Sql_Command As String

        ReadWorkFlow4 = -1
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

                        'Sql_Command2 = "SELECT * FROM """ + DateLogTable + """ WHERE (""職員番号"" = '" & value & "' AND ""日付"" = DATE '" + D1 + " ')"
                        'tb2 = db.ExecuteSql(Sql_Command2)

                        For j As Integer = 1 To data.Length - 1
                            Dim aa As String = "", bb As String = ""

                            Dim No As String = data(j)(13)
                            aa += """職員番号"""
                            bb += "'" + No + "'"
                            aa += ","
                            bb += ","

                            Dim Name As String = data(j)(5)
                            aa += """職員名"""
                            bb += "'" + Name + "'"
                            aa += ","
                            bb += ","

                            Dim DateTime1 As String = data(j)(8).Replace("/", "-") + ":00"
                            aa += """申請日"""
                            bb += "TIMESTAMP '" + DateTime1 + "'"
                            aa += ","
                            bb += ","

                            Dim Cat As String = data(j)(12).Replace("（変更前の日付を備考に記載）", "")
                            aa += """申請区分"""
                            bb += "'" + Cat + "'"
                            aa += ","
                            bb += ","

                            'Dim Kind1 As String, Kind2 As String, Kind As String
                            'Kind1 = data(j)(12).Replace("（備考欄に詳細記載）", "").Replace("/", vbCrLf)
                            'Kind2 = data(j)(13).Replace("（備考欄に詳細記載）", "").Replace("/", vbCrLf)
                            'If Kind1 <> "" Then
                            '    Kind = Kind1
                            'Else
                            '    Kind = Kind2
                            'End If
                            'aa += """種類"""
                            'bb += "'" + Kind + "'"
                            'aa += ","
                            'bb += ","

                            Dim Date1 As String, Date2 As String, Date3 As String, Date4 As String

                            If data(j)(14) <> "" Then
                                Date1 = data(j)(14).Replace("/", "-")
                                aa += """在宅勤務日1"""
                                bb += "DATE '" + Date1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(15) <> "" Then
                                Date2 = data(j)(15).Replace("/", "-")
                                aa += """在宅勤務日2"""
                                bb += "DATE '" + Date2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(16) <> "" Then
                                Date3 = data(j)(16).Replace("/", "-")
                                aa += """在宅勤務日3"""
                                bb += "DATE '" + Date3 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(17) <> "" Then
                                Date4 = data(j)(17).Replace("/", "-")
                                aa += """在宅勤務日4"""
                                bb += "DATE '" + Date4 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(18) <> "" Then
                                Dim 終日以外選択 As String = data(j)(18).Replace("'", "''")
                                aa += """終日以外選択"""
                                bb += "'" + 終日以外選択 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(19) <> "" Then
                                Dim 終日以外の理由 As String = data(j)(19).Replace("'", "''")
                                aa += """終日以外の理由"""
                                bb += "'" + 終日以外の理由 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(20) <> "" Then
                                Dim 連絡先 As String = data(j)(20).Replace("'", "''")
                                aa += """連絡先"""
                                bb += "'" + 連絡先 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(21) <> "" Then
                                Dim 実務する業務 As String = data(j)(21).Replace("'", "''")
                                aa += """実務する業務"""
                                bb += "'" + 実務する業務 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(22) <> "" Then
                                Dim パソコンのOS As String = data(j)(22).Replace("'", "''")
                                aa += """パソコンのOS"""
                                bb += "'" + パソコンのOS + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(23) <> "" Then
                                Dim セキュリティソフト As String = data(j)(23).Replace("'", "''")
                                aa += """セキュリティソフト"""
                                bb += "'" + セキュリティソフト + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(24) <> "" Then
                                Dim 使用するパソコン As String = data(j)(24).Replace("'", "''")
                                aa += """使用するパソコン"""
                                bb += "'" + 使用するパソコン + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(25) <> "" Then
                                Dim リモートデスクトップの方法 As String = data(j)(25).Replace("'", "''")
                                aa += """リモートデスクトップの方法"""
                                bb += "'" + リモートデスクトップの方法 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(26) <> "" Then
                                Dim 自宅PC確認事項1 As String = data(j)(26).Replace("'", "''")
                                aa += """自宅PC確認事項1"""
                                bb += "'" + 自宅PC確認事項1 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(27) <> "" Then
                                Dim 自宅PC確認事項2 As String = data(j)(27).Replace("'", "''")
                                aa += """自宅PC確認事項2"""
                                bb += "'" + 自宅PC確認事項2 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(28) <> "" Then
                                Dim 同意 As String = data(j)(28).Replace("'", "''")
                                aa += """同意"""
                                bb += "'" + 同意 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(29) <> "" Then
                                Dim 持ち出す情報 As String = data(j)(29).Replace("'", "''")
                                aa += """持ち出す情報"""
                                bb += "'" + 持ち出す情報 + "'"
                                aa += ","
                                bb += ","
                            End If

                            If data(j)(30) <> "" Then
                                Dim 所属長コメント As String = data(j)(30).Replace("'", "''")
                                aa += """所属長コメント"""
                                bb += "'" + 所属長コメント + "'"
                                aa += ","
                                bb += ","
                            End If


                            Dim ReMarks As String = data(j)(31).Replace("'", "''")
                            aa += """備考"""
                            bb += "'" + ReMarks + "'"
                            aa += ","
                            bb += ","

                            aa += """処理"""
                            bb += "'未処理'"
                            aa += ","
                            bb += ","

                            aa += """バージョン"""
                            bb += "'" + Version + "'"
                            'aa += ","
                            'bb += ","

                            Dim Sql_Command2 As String = "SELECT ""職員番号"" FROM """ + HomeWorkTable +
                                """ WHERE (""職員番号"" = '" + No + "' AND ""申請日"" = TIMESTAMP '" + DateTime1 + "' AND ""在宅勤務日1"" = DATE '" + Date1 + " ')"
                            Dim tb2 As DataTable = db.ExecuteSql(Sql_Command2)
                            Dim n2 As Integer = tb2.Rows.Count

                            If n2 = 0 Then
                                Sql_Command = "INSERT INTO """ + HomeWorkTable + """ (" + aa + ") VALUES (" + bb + ")"
                                tb = db.ExecuteSql(Sql_Command)
                            End If
                        Next

                        db.Disconnect()

                        Dim file2 As String = outPath4 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)
                    Else
                        Dim file2 As String = delPath4 + "\" + System.IO.Path.GetFileName(ff(i))
                        System.IO.File.Move(ff(i), file2)

                    End If


                End If
            Next
            ReadWorkFlow4 = 0

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
