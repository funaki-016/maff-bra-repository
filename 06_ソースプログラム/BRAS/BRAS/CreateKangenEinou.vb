'------------------------------------------------------------------------------------------
'| REV | 変更年月日 | 変更者             | 変更内容
'------------------------------------------------------------------------------------------
'| 001 | 2021/05/25 | TSP                | 稼働障害　330205/7806対応
'| 002 | 2023/01/25 | Daiko              | 要件No4 バージョン区分追加
'| 003 | 2023/12/11 | Daiko              | 変更要件No.5 項番の用語の解説を追加・修正
'------------------------------------------------------------------------------------------

Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Globalization

Public Class CreateKangenEinou
    Inherits BRA7420R

    Public Class 営農個人


        '集計結果表で「円」となるもの
        Public Shared shukeiEn() As String = {"S010513", "S010523", "S010527", "S010528", "S010529", "S010530", "S010537", "S010538", "S010542", "S031041", "S031046", "S031047", "S031048", "S031049", "S031329",
                                              "S031330", "S031331", "S031332", "S031333", "S031334", "S031335", "S031336", "S031337", "S031338", "S031339", "S031340", "S031341", "S031342", "S031343", "S031344",
                                              "S031345", "S031346", "S031347", "S031348", "S031349", "S031350", "S031351", "S031352"}
        Public Class sheet1
            Public Const WAREKI_CELL As String = "O11"
            Public Const EINOU_CELL As String = "J16"
            Public Const PICTURE_CELL As String = "G19"
            Public Const AISATSU_CELL As String = "E45"
            Public Const SHUKEI_CELL As String = "BA11:BC14"

        End Class

        Public Class sheet2
            Public Const GROUP1 As String = "A65:A72"
            Public Const GROUP2 As String = "E65:F72"

            Public Const KAISETSU1 As String = "角丸四角形吹き出し 21"
            Public Const KAISETSU2 As String = "角丸四角形吹き出し 50"
            Public Const TITLE As String = "角丸四角形 9"
            Public Const GRAPH As String = "テキスト ボックス 1"

            Public Const KAISETSU_START As String = "【費目の解説】"
            Public Shared 費目の解説 As New Dictionary(Of String, String) From {
                {"K011118", "◎雇人費:雇用労賃、交通費、賞与、福利厚生費などです。"},
                {"K011119", "◎種苗費:種もみ、種子、種いも、苗類などの購入費用です。"},
                {"K011120", "◎素畜費:素畜購入費用や種付料などです。"},
                {"K011121", "◎肥料費:化学肥料、有機肥料などの購入費用です。"},
                {"K011122", "◎飼料費:飼料の購入費用、自給飼料の振替額です。"},
                {"K011123", "◎農薬衛生費:農薬・家畜薬品などの購入費用、共同防除の負担金です。"},
                {"K011124", "◎諸材料費:被覆用ビニール、縄、おがくず、鉢などの購入費用です。"},
                {"K011125", "◎動力光熱費:水道料、電気料、燃油代金などです。"},
                {"K011126", "◎修繕費:農業生産に使用した機械、車両、建物などの修繕に要した費用です。"},
                {"K011127", "◎農具費:取得価額が10万円未満又は耐用年数が1年未満の農具の購入費用です。"},
                {"K011128", "◎作業用衣料費:作業服、軍手、長靴、地下足袋などの購入費用です。"},
                {"K011129", "◎地代・賃借料:農地、農機具、農用建物などの使用代金です。"},
                {"K011130", "◎土地改良費:土地改良負担金、客土費用、揚排水施設などの使用維持管理費です。"},
                {"K011131", "◎租税公課:農業に関する固定資産税、印紙税、組合・部会費、支払消費税などです。"},
                {"K011132", "◎利子割引料:借入金の支払利息、手形割引料、債務保証料などです。"},
                {"K011133", "◎荷造運賃手数料:販売に要した資材（ダンボール、袋、紐など）の代金、運送費用などです。"},
                {"K011139", "◎減価償却費:農業生産に使用した農機具、建物などの減価償却費です。"}
                }
        End Class

        Public Class sheet3
            Public Const 農業粗収益 As String = "K010237"
            Public Const 農業経営費 As String = "K010239"
            Public Const 農業所得 As String = "K010236"

            Public Const グラフ As String = "グラフ 15"

            Public Const GROUP1 As String = "L57:Y60"
            Public Const GROUP2 As String = "C40:H43"

            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From {
                {1, "6"}, {2, "3,9"}, {3, "0,6,12"}, {4, "0,3,6,9"}, {5, "0,3,6,9,12"}
                }

        End Class

        Public Class sheet4
            Public Const 経年変化タイトル As String = "Text Box 16"
            Public Const 経年変化 As String = "の経年変化"
            Public Const グラフ As String = "グラフ 19"
            Public Const 単位1 As String = "Text Box 1082"
            Public Const 単位2 As String = "Text Box 1083"
            Public Const GROUP1 As String = "K56:P58"

            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From {
            {1, "3"}, {2, "2,4"}, {3, "1,3,5"}, {4, "1,2,3,4"}, {5, "1,2,3,4,5"}
            }
        End Class


        Public Class sheet5
            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From {
                {1, "3"}, {2, "2,4"}, {3, "1,3,5"}, {4, "1,2,3,4"}, {5, "1,2,3,4,5"}
                }

            Public Const GROUP1 As String = "S63:X67"
            Public Const GROUP2 As String = "S70:X75"
            Public Const 農業粗収益 As String = "K010237"
            Public Const 農業経営費 As String = "K010239"

            Public Const グラフ1 As String = "グラフ 26"
            Public Const グラフ2 As String = "グラフ 33"

            '農業粗収益
            Public Shared soshueki() As String = {"K020906", "K020912", "K020915", "K020919", "K020920", "K020921", "K020927", "K020928", "K020929", "K020931", "K020934", "K020937", "K020942", "K020943" _
                            , "K021205", "K021206", "K021211", "K021214", "K021215", "K021216", "K021219", "K021223"}
            '農業経営費
            Public Shared keieihi() As String = {"K011118", "K011119", "K011120", "K011121", "K011122", "K011123", "K011124", "K011125", "K011126", "K011127", "K011128", "K011129", "K011130" _
                                              , "K011131", "K011132", "K011133", "K011139"}
        End Class

        Public Class sheet6
            Public Shared THIS_YEAR As String = "G5"
            Public Shared PREV_YEAR As String = "H5"
            Public Shared THIS_YEAR_SYU As String = "J7"
            Public Shared PREV_YEAR_SYU As String = "K7"
            Public Const HEIKIN As String = "J5"
            Public Const MENSEKI As String = "J6"

            Public Const SHUEKI_START As Integer = 1
            Public Const SHUEKI_END As Integer = 4
            Public Const KEIEI_START As Integer = 7
            Public Const KEIEI_END As Integer = 14

            Public Const SHUEKI_TITLE As String = "E10:E13"
            Public Const KEIEI_TITLE As String = "E16:E23"
            Public Const KOBETSU_DATA As String = "G9:H29"
            Public Const SHUKEI_DATA As String = "J9:K29"

            Public Const ZENBUN As String = "{0}（１～12月）のお宅の経営の内容{1}の結果を参考に掲載しました。"

            '農業粗収益
            Public Shared soshueki() As String = {"K020906", "K020912", "K020915", "K020919", "K020920", "K020921", "K020927", "K020928", "K020929", "K020931", "K020934", "K020937", "K020942", "K020943" _
                            , "K021205", "K021206", "K021211", "K021214", "K021215", "K021216", "K021219", "K021223"}
            '農業経営費
            Public Shared keieihi() As String = {"K011118", "K011119", "K011120", "K011121", "K011122", "K011123", "K011124", "K011125", "K011126", "K011127", "K011128", "K011129", "K011130" _
                                              , "K011131", "K011132", "K011133", "K011135", "K011139"}

        End Class

        Public Class sheet7
            Public Shared THIS_YEAR As String = "G5"
            Public Shared PREV_YEAR As String = "H5"
            Public Shared THIS_YEAR_SYU As String = "J7"
            Public Shared PREV_YEAR_SYU As String = "K7"
            Public Const HEIKIN As String = "J5"
            Public Const MENSEKI As String = "J6"

            Public Shared GROUP1_TITLE As String = "C9:F15"
            Public Shared KOBETSU_DATA1 As String = "G9:H22"
            Public Shared SHUKEI_DATA1 As String = "J9:K22"
            Public Shared KOBETSU_DATA2 As String = "G9:H15"
            Public Shared SHUKEI_DATA2 As String = "J9:K15"

        End Class

        Public Class sheet8
            Public Shared THIS_YEAR As String = "G5"
            Public Shared PREV_YEAR As String = "H5"
            Public Shared THIS_YEAR_SYU As String = "J7"
            Public Shared PREV_YEAR_SYU As String = "K7"
            Public Const HEIKIN As String = "J5"
            Public Const MENSEKI As String = "J6"

            Public Shared GROUP1_TITLE As String = "B9:F12"
            Public Shared KOBETSU_DATA1 As String = "G9:H21"
            Public Shared SHUKEI_DATA1 As String = "J9:K21"
            Public Shared KOBETSU_DATA2 As String = "G9:H12"
            Public Shared SHUKEI_DATA2 As String = "J9:K12"

        End Class

        Public Class sheet9
            Public Shared THIS_YEAR As String = "D5"
            Public Shared PREV_YEAR As String = "E5"

            Public Shared TITLE As String = "B9:C26"
            Public Shared KOBETSU_DATA As String = "D9:E26"
        End Class

        Public Class sheet10
            Public Shared お問合せ先 As New Dictionary(Of Integer, String) From {
               {ComConst.還元資料.お問合せ先_明細番号.事務局_拠点, "E46"},
               {ComConst.還元資料.お問合せ先_明細番号.住所, "E48"},
               {ComConst.還元資料.お問合せ先_明細番号.電話番号, "AA50"},
               {ComConst.還元資料.お問合せ先_明細番号.FAX, "AA51"},
               {ComConst.還元資料.お問合せ先_明細番号.担当者名, "Z53"}
             }

        End Class
    End Class

    Public Class 営農法人
        Public Class Sheet1
            Public Const WAREKI_CELL As String = "O8"
            Public Const EINOU_CELL As String = "J14"
            Public Const FOLDER_NAME As String = "還元資料_画像"
            Public Const PICTURE_CELL As String = "G18"
            Public Const AISATSU_CELL As String = "E43"
            Public Const SHUKEI_CELL As String = "BA11:BC14"
        End Class

        Public Class Sheet2
            Public Const GROUP1 As String = "A63:A70"
            Public Const GROUP2 As String = "E63:F70"

            Public Const KAISETSU1 As String = "角丸四角形吹き出し 21"
            Public Const KAISETSU2 As String = "角丸四角形吹き出し 4"
            Public Const TITLE As String = "角丸四角形 9"
            Public Const GRAPH As String = "テキスト ボックス 1"

            Public Const KAISETSU_START As String = "【費目の解説】"
            Public Shared 費目の解説 As New Dictionary(Of String, String) From {
                {"K021231", "◎雇人費:雇用労賃、交通費、賞与、福利厚生費などです。"},
                {"K021232", "◎種苗費:種もみ、種子、種いも、苗類などの購入費用です。"},
                {"K021233", "◎素畜費:素畜購入費用や種付料などです。"},
                {"K021234", "◎肥料費:化学肥料、有機肥料などの購入費用です。"},
                {"K021235", "◎飼料費:飼料の購入費用、自給飼料の振替額です。"},
                {"K021236", "◎農薬衛生費:農薬・家畜薬品などの購入費用、共同防除の負担金です。"},
                {"K021237", "◎諸材料費:被覆用ビニール、縄、おがくず、鉢などの購入費用です。"},
                {"K021238", "◎動力光熱費:水道料、電気料、燃油代金などです。"},
                {"K021239", "◎修繕費:農業生産に使用した機械、車両、建物などの修繕に要した費用です。"},
                {"K021240", "◎農具費:取得価額が10万円未満又は耐用年数が1年未満の農具の購入費用です。"},
                {"K021241", "◎作業用衣料費:作業服、軍手、長靴、地下足袋などの購入費用です。"},
                {"K021242", "◎地代・賃借料:農地、農機具、農用建物などの使用代金です。"},
                {"K021243", "◎土地改良費:土地改良負担金、客土費用、揚排水施設などの使用維持管理費です。"},
                {"K021244", "◎租税公課:農業に関する固定資産税、印紙税、組合・部会費、支払消費税などです。"},
                {"K021245", "◎利子割引料:借入金の支払利息、手形割引料、債務保証料などです。"},
                {"K021246", "◎荷造運賃手数料:販売に要した資材（ダンボール、袋、紐など）の代金、運送費用などです。"},
                {"K021249", "◎減価償却費:農業生産に使用した農機具、建物などの減価償却費です。"}
                }
        End Class

        Public Class sheet3
            Public Const 農業収入 As String = "K010833"
            Public Const 農業支出 As String = "K010834"
            Public Const 営業利益 As String = "K010843"

            Public Const グラフ As String = "グラフ 15"

            Public Const GROUP1 As String = "M58:Z61"
            Public Const GROUP2 As String = "B35:G38"

            Public Const KAISETSU As String = "角丸四角形吹き出し 5"
            Public Const KAISETSU_START As String = "【用語の解説】"
            Public Const KAISETSU_1 As String = "◎　農業粗収益とは、農業収入に共済・補助金等の雑収入を加えた金額です。"
            Public Const KAISETSU_2 As String = "◎　営業利益（うち農業）とは、農業収入から農業支出を差し引いた金額です。"
            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From {
                {1, "6"}, {2, "3,9"}, {3, "0,6,12"}, {4, "0,3,6,9"}, {5, "0,3,6,9,12"}
                }

            Public Class 用語の解説
                Public fontPoint As Integer
                Public kaisetsu As String
            End Class

            'REV_003↓
            'Public Shared KAISETSU_3 As New Dictionary(Of String, 用語の解説) From {
            '    {"K010507", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　総資本経常利益率とは、総資本（貸借対照表における負債と純資産の合計）に対する経常利益の割合を示す指標。" & vbLf & "　企業の経常的な活動による業績状態を示すもので、投下した資本に対してどの程度の利益を生み出したのかを知ることができる。この数値が高ければ、企業が効率的に資本運用できているといえる。"}},
            '    {"K010508", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高営業利益率とは、売上高（事業収入）に対する営業利益の割合を示す指標。" & vbLf & "　企業ごとの財務構造からの影響を排除した上で、本業でどのくらい効率的に儲けたのかを知ることができる。この数値が高いほど本業の収益力が高いといえる。"}},
            '    {"K010509", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高経常利益率とは、売上高（事業収入）に対する経常利益の割合を示す指標。" & vbLf & "　財務活動なども含めた企業の事業活動全体における利益率を表すもので、通常、金融収支の良し悪しや資金調達力の違いなどの財務面も含めた総合的な収益性が反映されるため、この数値が高いほど収益性が高い（良い）といえる。"}},
            '    {"K010510", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　売上高当期純利益率とは、売上高（事業収入）に対する当期純利益の割合を示す指標。" & vbLf & "　企業の配当の原資である当期純利益の水準をみるものであり、この数値が高いほど収益性が高い（良い）といえる。通常は、売上高経常利益率と売上高当期純利益率とを比べてみて、その差が大きい場合には特別損益の内容（当期純利益の額が特別損益による突発的なものなのかどうかの確認）にも注意が必要となる。"}},
            '    {"K010511", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　自己資本当期純利益率とは、自己資本（純資産）に対する当期純利益の割合を示す指標。" & vbLf & "　株主持分である自己資本に対してどれだけのリターン（当期純利益）が生み出されているかを示し、一般的に10％を上回ると優良な企業だといわれており、投資価値のある企業だと判断される。"}},
            '    {"K010512", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高費用比率とは、売上高（事業収入）に対する損益計算書上の費用（事業支出）の割合を示す指標。"}},
            '    {"K010513", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　売上高原価比率とは、売上高（事業収入）に対する売上原価の割合を示す指標。" & vbLf & "　一般に企業においてこの比率が上昇している場合、売上原価がかさみ、利益を確保しにくくなっている状況を示す。売上に占める売上原価の構成比が小さいほどこの数値も低くなり、この数値が低いほど、商品の付加価値と販売管理費を賄う収益源が大きくなる。この指標により、その企業の時系列での数値の変化や、同一部門・同規模の企業の数値との比較などをチェックする。"}},
            '    {"K010514", New 用語の解説 With {.fontPoint = 9, .kaisetsu = "◎　売上高販売管理費比率とは、売上高（事業収入）に対する販売及び一般管理費の割合を示す指標。" & vbLf & "　販売及び一般管理費は費用であることから、基本はこの数値が低ければ低いほど良いと判断される。ただし、一概に数値が高いからといっても悪いともいえず、販売及び一般管理費への資本投下が売上高や売上総利益の伸びに効果的に結びついているかどうかが大事であり、販売及び一般管理費が売上や利益に貢献していないのであればそれは削減対象になる。販売及び一般管理費の無駄を省き効率的に抑えることができれば、売上高（事業収入）は変わらずとも営業利益を高められるようになる。"}},
            '    {"K010515", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額とは、企業が一定期間に生み出した利益で、経営向上の程度を示す指標。"}},
            '    {"K010516", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高付加価値率とは、売上高（事業収入）に対する付加価値額の割合を示す指標。" & vbLf & "　この数値が高い場合は、企業が新しく創造した価値の割合が大きいといえる。"}},
            '    {"K010517", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　事業従事者1人当たり給与総額とは、事業従事者がどれだけ効率的に成果を生み出し、それがどの程度人件費に使われているかを示す指標。" & vbLf & "　この数値が高すぎると人件費が高すぎるか収益構造に問題があることが分かり、低すぎるともしかしたら給与水準が低すぎる可能性がある。"}},
            '    {"K010518", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額給与総額率とは、企業の生産活動によって新たに生み出された付加価値に占める人件費の割合を示す指標。" & vbLf & "　この数値が低下する要因は、付加価値が増加したことか人件費などの固定費が減少したことのどちらかである。"}},
            '    {"K010519", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　事業従事者1人当たり付加価値額とは、事業従事者がどれだけ効率的に成果を生み出したかを数値化した指標の一つ。" & vbLf & "　この数値が高い場合は、１人の事業従事者が生み出す付加価値が高いといえる。"}},
            '    {"K010520", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　事業従事者1人当たり売上高とは、事業従事者がどれだけ効率的に成果を生み出したかを数値化した指標の一つ。" & vbLf & "　この数値が高い場合は、１人の事業従事者が生み出す売上高（事業収入）が高いといえる。"}},
            '    {"K010521", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　総資本回転率とは、事業に投資された総資本がどれだけ有効に活用されたかを示す指標。" & vbLf & "　企業に存在する全ての資本が直接的に売上獲得に貢献したと仮定して、売上高（事業収入）が総資本（貸借対照表における負債と純資産の合計）の何倍であるか（何回転しているか）により企業が調達した総資本の活用度合いを示したもの。"}},
            '    {"K010522", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　自己資本比率とは、総資本（貸借対照表における負債と純資産の合計）に対する自己資本の割合を示す指標。" & vbLf & "　一般的にこの数値が50％を超えているとかなり優良であるといわれ、20～30％くらいでも良い印象がある。中小企業の場合は15％くらいが平均とされている。"}},
            '    {"K010523", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　流動比率とは、流動負債に対する流動資産の割合を示し、短期的な支払能力を分析する際に用いる指標。" & vbLf & "　１年以内に現金化できる資産が、１年以内に返済すべき負債をどれだけ上回っているかを表し、一般的にこの数値が120％以上であれば短期的な資金繰りには困らないとされ、100％を下回っていると支払能力に不安があるとされる。"}},
            '    {"K010525", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業依存度とは、事業等の所得に占める農業所得の割合をいい、経済活動による所得のうち、どれだけが農業所得に依存しているかを示す指標。"}},
            '    {"K010526", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業所得率とは、農業粗収益のうち、どれだけが農業所得として実現するかを示す指標。"}},
            '    {"K010527", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額とは、農業粗収益から物財費（雇人費、地代・賃借料及び利子割引料を含まない農業経営費）を差し引いたもので、農業生産により新たに生み出された付加価値額を示す指標。"}},
            '    {"K010528", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業粗収益付加価値率とは、農業粗収益のうち、どれだけが農業生産によって新たに付加価値額として生み出されたものであるかを示す指標。"}},
            '    {"K010529", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり付加価値額とは、農業従事者１人当たりの付加価値額でみた収益性を示す指標。"}},
            '    {"K010530", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり農業粗収益とは、農業従事者１人当たりの農業粗収益でみた収益性を示す指標。"}},
            '    {"K010531", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり農業経営費とは、農業従事者１人当たりの農業経営費でみた収益性を示す指標。"}},
            '    {"K010532", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり農業所得とは、農業従事者１人当たりの農業所得でみた収益性を示す指標。"}},
            '    {"K010533", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　有給役員農業労働１時間当たり農業粗収益とは、投下された有給役員農業労働の単位時間当たりの農業粗収益でみた労働収益性を示す指標。" & vbLf & "この指標により異なる部門間や同一部門での規模間比較が可能。"}},
            '    {"K010534", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　有給役員農業労働１時間当たり農業経営費とは、投下された有給役員農業労働の単位時間当たりの農業経営費でみた労働収益性を示す指標。" & vbLf & "この指標により異なる部門間や同一部門での規模間比較が可能。"}},
            '    {"K010535", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　有給役員農業労働１時間当たり農業所得とは、投下された有給役員農業労働の単位時間当たりの農業所得でみた労働収益性を示す指標。" & vbLf & "この指標により異なる部門間や同一部門での規模間比較が可能。"}},
            '    {"K010536", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　農業固定資産装備率とは、固定資産装備の大きさを示す指標。" & vbLf & "　一般的には労働者１人当たりの固定資産額をいうが、農業の場合は、農業労働に季節性があること等から自営農業労働１時間当たりの固定資産額で示した。" & vbLf & "　注：　「自営農業労働時間」とは、自営農業労働時間と農作業受託に係わる労働時間を合わせたものである。"}},
            '    {"K010537", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業固定資産回転率とは、農業固定資産の運用効率、利用度の状況をみる指標。"}},
            '    {"K010538", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　営農類型規模当たり農業粗収益とは、営農類型規模当たり（例：水田作の場合、水田作作付延べ面積10ａ当たり）の農業粗収益でみた収益性を示す指標。"}},
            '    {"K010539", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　営農類型規模当たり農業経営費とは、営農類型規模当たり（例：水田作の場合、水田作作付延べ面積10ａ当たり）の農業経営費でみた収益性を示す指標。"}},
            '    {"K010540", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　営農類型規模当たり農業所得とは、営農類型規模当たり（例：水田作の場合、水田作作付延べ面積10ａ当たり）の農業所得でみた収益性を示す指標。"}},
            '    {"K010542", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業所得率とは、農業生産関連事業収入のうち、どれだけが農業生産関連事業所得として実現するかを示す指標。"}},
            '    {"K010543", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業1時間当関連事業所得とは、投下された農業生産関連事業労働の単位時間当たりの農業生産関連事業所得でみた労働収益性を示す指標。" & vbLf & "　この指標により異なる農業生産関連事業間や同一農業生産関連事業での規模間比較が可能。"}},
            '    {"K010544", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業従事者1人当関連事業所得とは、農業生産関連事業従事者１人当たりの農業生産関連事業所得でみた収益性を示す指標。"}},
            '    {"K010545", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業従事者1人当関連事業労働時間とは、農業生産関連事業従事者１人当たりの農業生産関連事業投下労働時間でみた投下労働量を示す指標。"}},
            '    {"K010546", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額とは、農業生産関連事業収入から農業生産関連事業物財費（雇人費を含まない農業生産関連事業支出）を差し引いたもので、農業生産関連事業により新たに生み出された付加価値額を示す指標。"}},
            '    {"K010547", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値率とは、農業生産関連事業収入のうち、どれだけが農業生産関連事業によって新たに付加価値額として生み出されたものであるかを示す指標。"}},
            '    {"K010548", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業1時間当付加価値額とは、投下された農業生産関連事業労働の単位時間当たりの農業生産関連事業付加価値額でみた労働収益性を示す指標。"}},
            '    {"K010549", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業従事者1人当付加価値額とは、農業生産関連事業従事者１人当たりの農業生産関連事業付加価値額でみた労働収益性を示す指標。"}},
            '    {"K010551", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業及び関連事業所得率とは、農業収入と農業生産関連事業収入の合計額のうち、どれだけが農業所得と農業生産関連事業所得の合計額として実現するかを示す指標。"}},
            '    {"K010552", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業及び関連事業付加価値額とは、農業により新たに生み出された付加価値額と農業生産関連事業により新たに生み出された付加価値額の合計額を示す指標。"}},
            '    {"K010553", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業及び関連事業付加価値率とは、農業粗収益と農業生産関連事業収入の合計額のうち、どれだけが農業と農業生産関連事業によって新たに付加価値額として生み出されたものであるかを示す指標。"}}
            '    }
            Public Shared KAISETSU_3 As New Dictionary(Of String, 用語の解説) From {
                {"K010507", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　総資本経常利益率とは、総資本（貸借対照表における負債と純資産の合計）に対する経常利益の割合を示す指標。" & vbLf & "　企業の経常的な活動による業績状態を示すもので、投下した資本に対してどの程度の利益を生み出したのかを知ることができる。この数値が高ければ、企業が効率的に資本運用できているといえる。"}},
                {"K010508", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高営業利益率とは、売上高（事業収入）に対する営業利益の割合を示す指標。" & vbLf & "　企業ごとの財務構造からの影響を排除した上で、本業でどのくらい効率的に儲けたのかを知ることができる。この数値が高いほど本業の収益力が高いといえる。"}},
                {"K010509", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高経常利益率とは、売上高（事業収入）に対する経常利益の割合を示す指標。" & vbLf & "　財務活動なども含めた企業の事業活動全体における利益率を表すもので、通常、金融収支の良し悪しや資金調達力の違いなどの財務面も含めた総合的な収益性が反映されるため、この数値が高いほど収益性が高い（良い）といえる。"}},
                {"K010510", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　売上高当期純利益率とは、売上高（事業収入）に対する当期純利益の割合を示す指標。" & vbLf & "　企業の配当の原資である当期純利益の水準をみるものであり、この数値が高いほど収益性が高い（良い）といえる。通常は、売上高経常利益率と売上高当期純利益率とを比べてみて、その差が大きい場合には特別損益の内容（当期純利益の額が特別損益による突発的なものなのかどうかの確認）にも注意が必要となる。"}},
                {"K010511", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　自己資本当期純利益率とは、自己資本（純資産）に対する当期純利益の割合を示す指標。" & vbLf & "　株主持分である自己資本に対してどれだけのリターン（当期純利益）が生み出されているかを示す。一般的に、優良な企業は本指標が10％を上回る。"}},
                {"K010512", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高費用比率とは、売上高（事業収入）に対する損益計算書上の費用（事業支出）の割合を示す指標。"}},
                {"K010513", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　売上高原価比率とは、売上高（事業収入）に対する売上原価の割合を示す指標。" & vbLf & "　一般に企業においてこの比率が上昇している場合、売上原価がかさみ、利益を確保しにくくなっている状況を示す。売上に占める売上原価の構成比が小さいほどこの数値も低くなり、この数値が低いほど、商品の付加価値と販売管理費を賄う収益源が大きくなる。この指標により、その企業の時系列での数値の変化や、同一部門・同規模の企業の数値との比較などをチェックする。"}},
                {"K010514", New 用語の解説 With {.fontPoint = 9, .kaisetsu = "◎　売上高販売管理費比率とは、売上高（事業収入）に対する販売及び一般管理費の割合を示す指標。" & vbLf & "　販売及び一般管理費は費用であることから、基本はこの数値が低ければ低いほど良い。ただし、一概に数値が高いからといっても悪いともいえず、販売及び一般管理費への資本投下が売上高や売上総利益の伸びに効果的に結びついているかどうかが大事であり、販売及び一般管理費が売上や利益に貢献していないのであればそれは削減対象になる。販売及び一般管理費の無駄を省き効率的に抑えることができれば、売上高（事業収入）は変わらずとも営業利益を高められるようになる。"}},
                {"K010515", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額とは、企業が一定期間に生み出した利益で、経営向上の程度を示す指標。"}},
                {"K010516", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　売上高付加価値率とは、売上高（事業収入）に対する付加価値額の割合を示す指標。" & vbLf & "　この数値が高い場合は、企業が新しく創造した価値の割合が大きいといえる。"}},
                {"K010517", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　事業従事者1人当たり給与総額とは、事業従事者がどれだけ効率的に成果を生み出し、それがどの程度人件費に使われているかを示す指標。" & vbLf & "　この数値が高すぎると人件費が高すぎるか収益構造に問題があることが分かり、低すぎるともしかしたら給与水準が低すぎる可能性がある。"}},
                {"K010518", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額給与総額率とは、企業の生産活動によって新たに生み出された付加価値に占める人件費の割合を示す指標。" & vbLf & "　この数値が低下する要因は、付加価値が増加したことか人件費などの固定費が減少したことのどちらかである。"}},
                {"K010519", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　事業従事者1人当たり付加価値額とは、事業従事者がどれだけ効率的に成果を生み出したかを数値化した指標の一つ。" & vbLf & "　この数値が高い場合は、１人の事業従事者が生み出す付加価値が高いといえる。"}},
                {"K010520", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　事業従事者1人当たり売上高とは、事業従事者がどれだけ効率的に成果を生み出したかを数値化した指標の一つ。" & vbLf & "　この数値が高い場合は、１人の事業従事者が生み出す売上高（事業収入）が高いといえる。"}},
                {"K010521", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　総資本回転率とは、事業に投資された総資本がどれだけ有効に活用されたかを示す指標。" & vbLf & "　企業に存在する全ての資本が直接的に売上獲得に貢献したと仮定して、売上高（事業収入）が総資本（貸借対照表における負債と純資産の合計）の何倍であるか（何回転しているか）により企業が調達した総資本の活用度合いを示したもの。"}},
                {"K010522", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　自己資本比率とは、総資本（貸借対照表における負債と純資産の合計）に対する自己資本の割合を示す指標。" & vbLf & "　一般的に、かなり優良な企業は、この数値が50％を超えている。20％以上あれば良い印象の企業で、中小企業の平均は約15％。"}},
                {"K010523", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　流動比率とは、流動負債に対する流動資産の割合を示し、短期的な支払能力を分析する際に用いる指標。" & vbLf & "　１年以内に現金化できる資産が、１年以内に返済すべき負債をどれだけ上回っているかを表す。一般的に、短期的な資金繰りには困らない企業はこの数値が120％以上、支払能力に不安がある企業は100％を下回る。"}},
                {"K010525", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業依存度とは、事業等の所得に占める農業所得の割合をいい、経済活動による所得のうち、どれだけが農業所得に依存しているかを示す指標。"}},
                {"K010526", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業所得率とは、農業粗収益のうち、どれだけが農業所得として実現するかを示す指標。"}},
                {"K010527", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額とは、農業粗収益から物財費（雇人費、地代・賃借料及び利子割引料を含まない農業経営費）を差し引いたもので、農業生産により新たに生み出された付加価値額を示す指標。"}},
                {"K010528", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業粗収益付加価値率とは、農業粗収益のうち、どれだけが農業生産によって新たに付加価値額として生み出されたものであるかを示す指標。"}},
                {"K010529", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり付加価値額とは、農業従事者１人当たりの付加価値額でみた収益性を示す指標。"}},
                {"K010530", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり農業粗収益とは、農業従事者１人当たりの農業粗収益でみた収益性を示す指標。"}},
                {"K010531", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり農業経営費とは、農業従事者１人当たりの農業経営費でみた収益性を示す指標。"}},
                {"K010532", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業従事者1人当たり農業所得とは、農業従事者１人当たりの農業所得でみた収益性を示す指標。"}},
                {"K010533", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　有給役員農業労働１時間当たり農業粗収益とは、投下された有給役員農業労働の単位時間当たりの農業粗収益でみた労働収益性を示す指標。" & vbLf & "この指標により異なる部門間や同一部門での規模間比較が可能。"}},
                {"K010534", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　有給役員農業労働１時間当たり農業経営費とは、投下された有給役員農業労働の単位時間当たりの農業経営費でみた労働収益性を示す指標。" & vbLf & "この指標により異なる部門間や同一部門での規模間比較が可能。"}},
                {"K010535", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　有給役員農業労働１時間当たり農業所得とは、投下された有給役員農業労働の単位時間当たりの農業所得でみた労働収益性を示す指標。" & vbLf & "この指標により異なる部門間や同一部門での規模間比較が可能。"}},
                {"K010536", New 用語の解説 With {.fontPoint = 10, .kaisetsu = "◎　農業固定資産装備率とは、固定資産装備の大きさを示す指標。" & vbLf & "　一般的には労働者１人当たりの固定資産額をいうが、農業の場合は、農業労働に季節性があること等から自営農業労働１時間当たりの固定資産額で示した。" & vbLf & "　注：　「自営農業労働時間」とは、自営農業労働時間と農作業受託に係わる労働時間を合わせたものである。"}},
                {"K010537", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業固定資産回転率とは、農業固定資産の運用効率、利用度の状況をみる指標。"}},
                {"K010538", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　経営耕地10a当たり農業所得とは、経営耕地面積10a当たりの農業所得でみた収益性を示す指標。"}},
                {"K010539", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　営農類型規模当たり農業所得とは、営農類型規模当たり（例：水田作の場合、水田作作付延べ面積10ａ当たり）の農業所得でみた収益性を示す指標。"}},
                {"K010540", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　営農類型規模当たり付加価値額とは、営農類型規模当たり（例：水田作の場合、水田作作付延べ面積10ａ当たり）の付加価値額でみた収益性を示す指標。"}},
                {"K010542", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業所得率とは、農業生産関連事業収入のうち、どれだけが農業生産関連事業所得として実現するかを示す指標。"}},
                {"K010543", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業1時間当関連事業所得とは、投下された農業生産関連事業労働の単位時間当たりの農業生産関連事業所得でみた労働収益性を示す指標。" & vbLf & "　この指標により異なる農業生産関連事業間や同一農業生産関連事業での規模間比較が可能。"}},
                {"K010544", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業従事者1人当関連事業所得とは、農業生産関連事業従事者１人当たりの農業生産関連事業所得でみた収益性を示す指標。"}},
                {"K010545", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業従事者1人当関連事業労働時間とは、農業生産関連事業従事者１人当たりの農業生産関連事業投下労働時間でみた投下労働量を示す指標。"}},
                {"K010546", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値額とは、農業生産関連事業収入から農業生産関連事業物財費（雇人費を含まない農業生産関連事業支出）を差し引いたもので、農業生産関連事業により新たに生み出された付加価値額を示す指標。"}},
                {"K010547", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　付加価値率とは、農業生産関連事業収入のうち、どれだけが農業生産関連事業によって新たに付加価値額として生み出されたものであるかを示す指標。"}},
                {"K010548", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業1時間当付加価値額とは、投下された農業生産関連事業労働の単位時間当たりの農業生産関連事業付加価値額でみた労働収益性を示す指標。"}},
                {"K010549", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　関連事業従事者1人当付加価値額とは、農業生産関連事業従事者１人当たりの農業生産関連事業付加価値額でみた労働収益性を示す指標。"}},
                {"K010551", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業及び関連事業所得率とは、農業収入と農業生産関連事業収入の合計額のうち、どれだけが農業所得と農業生産関連事業所得の合計額として実現するかを示す指標。"}},
                {"K010552", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業及び関連事業付加価値額とは、農業により新たに生み出された付加価値額と農業生産関連事業により新たに生み出された付加価値額の合計額を示す指標。"}},
                {"K010553", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業及び関連事業付加価値率とは、農業粗収益と農業生産関連事業収入の合計額のうち、どれだけが農業と農業生産関連事業によって新たに付加価値額として生み出されたものであるかを示す指標。"}},
                {"K010554", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業固定資産千円当たり農業所得とは、投下された固定資産の単位金額当たりの農業所得でみた資本収益性を示す指標。"}},
                {"K010555", New 用語の解説 With {.fontPoint = 11, .kaisetsu = "◎　農業固定資産千円当たり付加価値額とは、投下された固定資本の単位金額当たりの付加価値額でみた資本生産性を示す指標。"}}
                }
            'REV_003↑
        End Class

        Public Class sheet4
            Public Const 経年変化タイトル As String = "Text Box 16"
            Public Const 経年変化 As String = "の経年変化"
            Public Const グラフ As String = "グラフ 19"
            Public Const 単位1 As String = "Text Box 1082"
            Public Const 単位2 As String = "Text Box 1083"
            Public Const GROUP1 As String = "O56:T58"

            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From {
            {1, "3"}, {2, "2,4"}, {3, "1,3,5"}, {4, "1,2,3,4"}, {5, "1,2,3,4,5"}
            }
        End Class

        Public Class sheet5
            Public Shared 経年変化開始位置 As New Dictionary(Of Integer, String) From {
                {1, "3"}, {2, "2,4"}, {3, "1,3,5"}, {4, "1,2,3,4"}, {5, "1,2,3,4,5"}
                }

            Public Const GROUP1 As String = "A61:F65"
            Public Const GROUP2 As String = "A67:F72"
            Public Const 農業収入 As String = "K010833"
            Public Const 農業支出 As String = "K021230"

            Public Const グラフ1 As String = "グラフ 26"
            Public Const グラフ2 As String = "グラフ 33"

            '農業収入
            Public Shared Shunyu() As String = {"K020906", "K020912", "K020915", "K020919", "K020920", "K020921", "K020927", "K020928", "K020929", "K020931", "K020934", "K020937", "K020942",
                                                  "K020943", "K021205", "K021206", "K021210", "K021213", "K021214", "K021215", "K021218", "K021219"}
            '農業支出
            Public Shared shishutsu() As String = {"K021231", "K021232", "K021233", "K021234", "K021235", "K021236", "K021237", "K021238", "K021239", "K021240", "K021241", "K021242", "K021243", "K021244",
                                                   "K021245", "K021246", "K021247", "K021249"}
        End Class

        Public Class sheet6
            Public Shared THIS_YEAR As String = "J5"
            Public Shared PREV_YEAR As String = "K5"
            Public Shared THIS_YEAR_SYU As String = "M7"
            Public Shared PREV_YEAR_SYU As String = "N7"
            Public Const HEIKIN As String = "M5"
            Public Const MENSEKI As String = "M6"

            Public Const SHUNYU_START As Integer = 3
            Public Const SHUNYU_END As Integer = 6
            Public Const GENKA_START As Integer = 15
            Public Const GENKA_END As Integer = 20
            Public Const HANBAIHI_START As Integer = 22
            Public Const HANBAIHI_END As Integer = 23

            Public Const SHUNYU_TITLE As String = "G12:H15"
            Public Const GENKA_TITLE As String = "H24:H29"
            Public Const KANRI_TITLE As String = "H31:H32"

            Public Const KOBETSU_DATA As String = "J9:K35"
            Public Const SHUKEI_DATA As String = "M9:N35"

            Public Const ZENBUN As String = "{0}調査の貴法人の経営内容{1}の結果を参考に掲載しました。"

            '農業収入
            Public Shared Shunyu() As String = {"K020906", "K020912", "K020915", "K020919", "K020920", "K020921", "K020927", "K020928", "K020929", "K020931", "K020934", "K020937", "K020942",
                                                  "K020943", "K021205", "K021206", "K021210", "K021213", "K021214", "K021215", "K021218", "K021219"}

            '生産原価
            Public Shared seisanGenka() As String = {"K031435", "K031436", "K031437", "K031438", "K031439", "K031440", "K031441", "K031442", "K031443", "K031444", "K031445", "K031446", "K031447",
                                                     "K031448", "K031449", "K031450", "K031451", "K031452", "K031453", "K031454", "K031455", "K031456", "K031457", "K031458", "K031459", "K031460", "K031461"
                                            }
            '販売費および一般管理費
            Public Shared hanbaiKanri() As String = {"K031635", "K031636", "K031637", "K031638", "K031639", "K031640", "K031641", "K031642", "K031643", "K031644", "K031645", "K031646", "K031647"}

            Public Const SUM_KOBETSU_RIEKI_DATA As String = "J9:K9"
            Public Const SUM_SHUKEI_RIEKI_DATA As String = "M9:N9"
            Public Const SUM_KOBETSU_SHISHUTSU_DATA As String = "J20:K20"
            Public Const SUM_SHUKEI_SHISHUTSU_DATA As String = "J20:K20"

        End Class

        Public Class sheet7
            Public Shared THIS_YEAR As String = "G5"
            Public Shared PREV_YEAR As String = "H5"
            Public Shared THIS_YEAR_SYU As String = "J7"
            Public Shared PREV_YEAR_SYU As String = "K7"
            Public Const HEIKIN As String = "J5"
            Public Const MENSEKI As String = "J6"

            Public Shared GROUP1_TITLE As String = "C21:F26"
            Public Shared KOBETSU_DATA1 As String = "G9:H33"
            Public Shared SHUKEI_DATA1 As String = "J9:K33"
            Public Shared KOBETSU_DATA2 As String = "G21:H26"
            Public Shared SHUKEI_DATA2 As String = "J21:K26"

        End Class

        Public Class sheet8
            Public Shared THIS_YEAR As String = "G5"
            Public Shared PREV_YEAR As String = "H5"
            Public Shared THIS_YEAR_SYU As String = "J7"
            Public Shared PREV_YEAR_SYU As String = "K7"
            Public Const HEIKIN As String = "J5"
            Public Const MENSEKI As String = "J6"

            Public Shared GROUP1_TITLE As String = "B9:F12"
            Public Shared KOBETSU_DATA1 As String = "G9:H21"
            Public Shared SHUKEI_DATA1 As String = "J9:K21"
            Public Shared KOBETSU_DATA2 As String = "G9:H12"
            Public Shared SHUKEI_DATA2 As String = "J9:K12"

        End Class

        Public Class sheet9
            Public Shared THIS_YEAR As String = "C6"
            Public Shared PREV_YEAR As String = "D6"

            Public Shared TITLE As String = "A10:B38"
            Public Shared KOBETSU_DATA As String = "C10:D38"
        End Class

        Public Class sheet10
            Public Shared お問合せ先 As New Dictionary(Of Integer, String) From {
               {ComConst.還元資料.お問合せ先_明細番号.事務局_拠点, "E46"},
               {ComConst.還元資料.お問合せ先_明細番号.住所, "E48"},
               {ComConst.還元資料.お問合せ先_明細番号.電話番号, "AA50"},
               {ComConst.還元資料.お問合せ先_明細番号.FAX, "AA51"},
               {ComConst.還元資料.お問合せ先_明細番号.担当者名, "Z53"}
             }

        End Class
    End Class

    Public Sub New(pCheckBox As ArrayList, pPrintFlg As Boolean, pOutputFlg As Boolean, pOutputPath As String, pSetting As DataTable,
               pKobetsu As Dictionary(Of String, Dictionary(Of String, DAOKobetsuKekkahyo.個別結果表項目)),
               pShukei As Dictionary(Of String, Dictionary(Of String, DAOSyukeiKekkahyo.集計結果表項目)),
               pMaster As DataTable, pTaiouMaster As DataTable, pKoumokuMaster As DataTable, pChosaNenKey As ArrayList,
               pSeidoUketoriItem As DataTable) ' REV 003

        ' REV_002↓
        'MyBase.New(pCheckBox, pPrintFlg, pOutputFlg, pOutputPath, pSetting, pKobetsu, pShukei, pMaster, pTaiouMaster, pKoumokuMaster, pChosaNenKey)
        MyBase.New(pCheckBox, pPrintFlg, pOutputFlg, pOutputPath, pSetting, pKobetsu, pShukei, pMaster, pTaiouMaster, pKoumokuMaster, pChosaNenKey, pSeidoUketoriItem)
        ' REV_002↑
    End Sub

    ''' <summary>
    ''' 作成実行
    ''' </summary>
    ''' <param name="pCensusNo">センサス番号</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Overrides Sub ReportEdit(xlSheets As Excel.Sheets)
        Dim xlSheet As Excel.Worksheet = Nothing

        '不要なシートのリスト
        Dim noSheetName As New ArrayList

        '1シート目
        xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)
        If CType(msheetList(0), Boolean) = True Then
            CreateSheet1(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '2シート目
        xlSheet = DirectCast(xlSheets.Item(2), Excel.Worksheet)
        If CType(msheetList(1), Boolean) = True Then
            CreateSheet2(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '3シート目
        xlSheet = DirectCast(xlSheets.Item(3), Excel.Worksheet)
        If CType(msheetList(2), Boolean) = True Then
            CreateSheet3(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '4シート目
        xlSheet = DirectCast(xlSheets.Item(4), Excel.Worksheet)
        If CType(msheetList(3), Boolean) = True Then
            CreateSheet4(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '5シート目
        xlSheet = DirectCast(xlSheets.Item(5), Excel.Worksheet)
        If CType(msheetList(4), Boolean) = True Then
            CreateSheet5(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '6シート目
        xlSheet = DirectCast(xlSheets.Item(6), Excel.Worksheet)
        If CType(msheetList(5), Boolean) = True Then
            CreateSheet6(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '7シート目
        xlSheet = DirectCast(xlSheets.Item(7), Excel.Worksheet)
        If CType(msheetList(6), Boolean) = True Then
            CreateSheet7(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '8シート目
        xlSheet = DirectCast(xlSheets.Item(8), Excel.Worksheet)
        If CType(msheetList(7), Boolean) = True Then
            CreateSheet8(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '9シート目
        xlSheet = DirectCast(xlSheets.Item(9), Excel.Worksheet)
        If CType(msheetList(8), Boolean) = True Then
            CreateSheet9(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '10シート目
        xlSheet = DirectCast(xlSheets.Item(10), Excel.Worksheet)
        If CType(msheetList(9), Boolean) = True Then
            CreateSheet10(xlSheet)
        Else
            noSheetName.Add(xlSheet.Name)
        End If
        ReleaseComObject(xlSheet)
        xlSheet = Nothing

        '不要なシートの削除
        If noSheetName.Count <> 0 Then
            For Each sheetName As String In noSheetName
                xlSheet = DirectCast(xlSheets.Item(sheetName), Excel.Worksheet)
                xlSheet.Delete()
                ReleaseComObject(xlSheet)
                xlSheet = Nothing
            Next
        End If

        'フッタの設定
        Dim firstPageFlg As Boolean = True
        Dim cencusNo As String

        cencusNo = mKobetsu(CStr(mChosanenKey(0))).Item(CENCUS).値
        cencusNo = cencusNo.Substring(0, 2) & Space(1) & cencusNo.Substring(2, 3) & Space(1) & cencusNo.Substring(5, 2) & Space(1) & cencusNo.Substring(7, 3) & Space(1) &
                    cencusNo.Substring(10, 3) & Space(1) & cencusNo.Substring(13, 3)

        For i As Integer = 1 To xlSheets.Count
            xlSheet = DirectCast(xlSheets.Item(i), Excel.Worksheet)
            xlSheet.PageSetup.RightFooter = cencusNo
            If xlSheet.Name <> "営個還１" AndAlso xlSheet.Name <> "営法還１" Then
                If firstPageFlg = True Then
                    xlSheet.PageSetup.FirstPageNumber = 1
                    firstPageFlg = False
                End If
            End If
        Next

    End Sub

    Private Sub CreateSheet1(pXlsheet As Excel.Worksheet)
        Dim rng As Excel.Range = Nothing

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 OrElse CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Dim warekiCell As String = String.Empty
            Dim ruikeiName As String = String.Empty
            Dim pictureCell As String = String.Empty
            Dim aisatsuCell As String = String.Empty
            Dim shukeiCell As String = String.Empty

            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                warekiCell = 営農個人.sheet1.WAREKI_CELL
                ruikeiName = 営農個人.sheet1.EINOU_CELL
                pictureCell = 営農個人.sheet1.PICTURE_CELL
                aisatsuCell = 営農個人.sheet1.AISATSU_CELL
                shukeiCell = 営農個人.sheet1.SHUKEI_CELL
            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                warekiCell = 営農法人.Sheet1.WAREKI_CELL
                ruikeiName = 営農法人.Sheet1.EINOU_CELL
                pictureCell = 営農法人.Sheet1.PICTURE_CELL
                aisatsuCell = 営農法人.Sheet1.AISATSU_CELL
                shukeiCell = 営農法人.Sheet1.SHUKEI_CELL
            End If
            '和暦の設定
            rng = pXlsheet.Range(warekiCell)
            rng.Value = getNengo(mChosanenKey(0).ToString, False)
            ReleaseComObject(rng)
            rng = Nothing

            '営農類型の設定

            rng = pXlsheet.Range(ruikeiName)
            rng.Value = ComUtil.getEinouName(mKobetsu(CStr(mChosanenKey(0))))
            ReleaseComObject(rng)
            rng = Nothing

            '画像の設定
            Dim einou As String = mKobetsu(CStr(mChosanenKey(0))).Item(ComConst.個別結果表.営農類型(CommonInfo.Chosakubun)).値
            Dim filePath As String = IniFileInfo.ExcelReportPath() & "\" & FOLDER_NAME & "\" & EINOU_PICTURE_FILE(einou)

            Dim xlShapes As Excel.Shapes
            rng = pXlsheet.Range(pictureCell)
            xlShapes = pXlsheet.Shapes
            Dim xlShape As Excel.Shape

            xlShape = xlShapes.AddPicture(Filename:=filePath,
               LinkToFile:=Microsoft.Office.Core.MsoTriState.msoFalse, SaveWithDocument:=Microsoft.Office.Core.MsoTriState.msoTrue,
               Left:=CSng(rng.Left), Top:=CSng(rng.Top), Width:=391, Height:=261.6)

            '挨拶文の設定
            rng = pXlsheet.Range(aisatsuCell)
            Dim tmp As String = CStr(rng.Value)
            tmp = tmp.Replace("年", getNengo(mChosanenKey(0).ToString, False))
            rng.Value = tmp

            '集計結果表の設定
            Dim syuukeiList(3, 2) As Object
            Dim d1 As DataRow()
            d1 = mSetteing.Select("項目番号 = '" & SHUKEIKEKKA & "'")

            For Each d2 As DataRow In d1
                If String.IsNullOrEmpty(CStr(d2.Item("設定値"))) Then
                    Continue For
                End If

                Dim value() As String = d2.Item("設定値").ToString.Split("_"c)
                If value.Count = 4 Then
                    syuukeiList(CType(d2.Item("明細番号"), Integer) - 1, 0) = value(1)
                    syuukeiList(CType(d2.Item("明細番号"), Integer) - 1, 1) = value(2)
                    syuukeiList(CType(d2.Item("明細番号"), Integer) - 1, 2) = value(3)
                End If
            Next

            rng = pXlsheet.Range(shukeiCell)
            rng.Value = syuukeiList
            ReleaseComObject(rng)
            rng = Nothing
        End If
    End Sub

    Private Sub CreateSheet2(pXlsheet As Excel.Worksheet)
        Dim rng As Excel.Range = Nothing

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '配列の作成 
            Dim TitleList(7, 0) As Object
            Dim DataList(7, 1) As Object
            '解説の作成
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder

            sb.Append(営農個人.sheet2.KAISETSU_START & vbCrLf)

            '個別結果表設定テーブルの取得
            Dim d1 As DataRow()
            Dim shukeiNo As String = ""
            Dim seisanHeikin As String = "0"
            d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_営農個人.営個還2_レーダーチャート & "'")

            Dim i As Integer = 0
            For Each d2 As DataRow In d1
                If String.IsNullOrEmpty(d2.Item("設定値").ToString) Then
                    Continue For
                End If

                'タイトルを配列に設定する
                TitleList(i, 0) = getKoumokuName(d2.Item("設定値").ToString, Name)
                DataList(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("設定値").ToString))
                DataList(i, 1) = ShukeiDataHenkan(d2.Item("設定値").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                sb.Append(営農個人.sheet2.費目の解説(d2.Item("設定値").ToString) & vbCrLf)
                i += 1
            Next

            Dim range As Excel.Range = Nothing

            range = pXlsheet.Range(営農個人.sheet2.GROUP1)
            range.Value = TitleList
            ReleaseComObject(range)
            range = Nothing

            range = pXlsheet.Range(営農個人.sheet2.GROUP2)
            range.Value = DataList
            ReleaseComObject(range)
            range = Nothing

            pXlsheet.Shapes.Item(営農個人.sheet2.KAISETSU2).TextEffect.Text = sb.ToString

            shukeiNo = SHUKEI_CHIIKI
            Dim Heikin As String = GetShukeiKekkahyoValue(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１, seisanHeikin, shukeiNo)
            Dim heikinName As String = String.Empty
            If Not IsNothing(Heikin) Then
                heikinName = ComConst.地域.リスト(Heikin).名称
                heikinName = heikinName.Replace("平均", String.Empty)
                heikinName = heikinName & "平均"
            End If

            Dim HyoujiTani As String = String.Empty
            Dim kibokaisou As String = String.Empty
            Dim comment As String = String.Empty

            d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.集計結果表 & " AND 明細番号 = " & ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
            If d1.Count <> 0 Then
                Dim values() As String = d1(0).Item("設定値").ToString.Split("_"c)
                kibokaisou = values(values.Count - 1)
            End If

            If mKobetsu(mChosanenKey(0).ToString).Item("K000006").値 = kibokaisou Then
                comment = "同規模" & " " & heikinName
            Else
                comment = GetHyoujiTani(True) & " " & heikinName
            End If
            pXlsheet.Shapes.Item(営農個人.sheet2.GRAPH).TextEffect.Text = comment

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '配列の作成 
            Dim TitleList(7, 0) As Object
            Dim DataList(7, 1) As Object
            '解説の作成
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder

            sb.Append(営農法人.Sheet2.KAISETSU_START & vbCrLf)

            '個別結果表設定テーブルの取得
            Dim d1 As DataRow()
            d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_営農法人.営法還2_レーダーチャート & "'")

            Dim i As Integer = 0
            For Each d2 As DataRow In d1
                If String.IsNullOrEmpty(d2.Item("設定値").ToString) Then
                    Continue For
                End If

                'タイトルを配列に設定する
                TitleList(i, 0) = getKoumokuName(d2.Item("設定値").ToString, Name)
                DataList(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(d2.Item("設定値").ToString))
                DataList(i, 1) = ShukeiDataHenkan(d2.Item("設定値").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                sb.Append(営農法人.Sheet2.費目の解説(d2.Item("設定値").ToString) & vbCrLf)
                i += 1
            Next

            Dim range As Excel.Range = Nothing

            range = pXlsheet.Range(営農法人.Sheet2.GROUP1)
            range.Value = TitleList
            ReleaseComObject(range)
            range = Nothing

            range = pXlsheet.Range(営農法人.Sheet2.GROUP2)
            range.Value = DataList
            ReleaseComObject(range)
            range = Nothing

            pXlsheet.Shapes.Item(営農法人.Sheet2.KAISETSU2).TextEffect.Text = sb.ToString

            Dim shukeiNo As String = ""
            Dim seisanHeikin As String = "0"

            shukeiNo = SHUKEI_CHIIKI
            Dim Heikin As String = GetShukeiKekkahyoValue(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１, seisanHeikin, shukeiNo)
            Dim heikinName As String = String.Empty
            If Not IsNothing(Heikin) Then
                heikinName = ComConst.地域.リスト(Heikin).名称
                heikinName = heikinName.Replace("平均", String.Empty)
                heikinName = heikinName & "平均"
            End If

            Dim HyoujiTani As String = String.Empty
            Dim kibokaisou As String = String.Empty
            Dim comment As String = String.Empty

            d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.集計結果表 & " AND 明細番号 = " & ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
            If d1.Count <> 0 Then
                Dim values() As String = d1(0).Item("設定値").ToString.Split("_"c)
                kibokaisou = values(values.Count - 1)
            End If

            If mKobetsu(mChosanenKey(0).ToString).Item("K000007").値 = kibokaisou Then
                comment = "同規模" & " " & heikinName
            Else
                comment = GetHyoujiTani(True) & " " & heikinName
            End If
            pXlsheet.Shapes.Item(営農法人.Sheet2.GRAPH).TextEffect.Text = comment

        End If
    End Sub

    Private Sub CreateSheet3(pXlSheet As Excel.Worksheet)
        '出力範囲の定義
        Dim keinenHenka(3, 13) As Object
        Dim start() As String
        Dim d1 As DataRow()
        Dim bunseki(3, 5) As Object
        Dim chartObject As Excel.ChartObject = Nothing
        Dim chart As Excel.Chart = Nothing

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '農業粗収益、経営費、所得の経年変化
            start = 営農個人.sheet3.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)


            For i As Integer = 1 To start.Count
                keinenHenka(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                keinenHenka(1, CInt(start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農個人.sheet3.農業粗収益).値
                keinenHenka(2, CInt(start(i - 1)) + 1) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農個人.sheet3.農業所得).値
                keinenHenka(3, CInt(start(i - 1)) + 1) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農個人.sheet3.農業経営費).値
            Next

            Dim range As Excel.Range = Nothing
            range = pXlSheet.Range(営農個人.sheet3.GROUP1)
            range.Value = keinenHenka
            ReleaseComObject(range)
            range = Nothing

            '分析指標の経年変化
            d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_営農個人.営個還3_分析指標の経年変化 & "'")


            Dim j As Integer = 1
            For Each d2 As DataRow In d1
                If String.IsNullOrEmpty(d2.Item("設定値").ToString) Then
                    Continue For
                End If

                '名称の設定
                bunseki(j, 0) = getKoumokuName(d2.Item("設定値").ToString, Name) & " (" & getKoumokuName(d2.Item("設定値").ToString, tani) & ")"
                For k As Integer = 1 To mChosanenKey.Count
                    bunseki(0, k) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - k)), True)
                    bunseki(j, k) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - k))).Item(d2.Item("設定値").ToString))
                Next
                j = j + 1
            Next

            range = Nothing
            range = pXlSheet.Range(営農個人.sheet3.GROUP2)
            range.Value = bunseki
            SetCellFormat(range, bunseki)
            ReleaseComObject(range)
            range = Nothing

            'グラフの設定
            chartObject = DirectCast(pXlSheet.ChartObjects(営農個人.sheet3.グラフ), Excel.ChartObject)
            chart = chartObject.Chart

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '農業収入、農業支出、農業利益の経年変化


            start = 営農法人.sheet3.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)


            For i As Integer = 1 To start.Count
                keinenHenka(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                keinenHenka(1, CInt(start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農法人.sheet3.農業収入).値
                keinenHenka(2, CInt(start(i - 1)) + 1) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農法人.sheet3.営業利益).値
                keinenHenka(3, CInt(start(i - 1)) + 1) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農法人.sheet3.農業支出).値
            Next

            Dim range As Excel.Range = Nothing
            range = pXlSheet.Range(営農法人.sheet3.GROUP1)
            range.Value = keinenHenka
            ReleaseComObject(range)
            range = Nothing

            '分析指標の経年変化
            d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_営農法人.営法還3_分析指標の経年変化 & "'")


            Dim j As Integer = 1
            For Each d2 As DataRow In d1
                If String.IsNullOrEmpty(d2.Item("設定値").ToString) Then
                    Continue For
                End If

                '名称の設定
                bunseki(j, 0) = getKoumokuName(d2.Item("設定値").ToString, Name) & " (" & getKoumokuName(d2.Item("設定値").ToString, tani) & ")"
                For k As Integer = 1 To mChosanenKey.Count
                    bunseki(0, k) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - k)), True)
                    bunseki(j, k) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - k))).Item(d2.Item("設定値").ToString))
                Next
                j = j + 1
            Next

            range = Nothing
            range = pXlSheet.Range(営農法人.sheet3.GROUP2)
            range.Value = bunseki
            SetCellFormat(range, bunseki)
            ReleaseComObject(range)
            range = Nothing

            '解説の設定
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder

            sb.Append(営農法人.sheet3.KAISETSU_START & vbCrLf)
            sb.Append(営農法人.sheet3.KAISETSU_1 & vbCrLf)
            sb.Append(営農法人.sheet3.KAISETSU_2 & vbCrLf)

            If d1.Count <> 0 Then
                sb.Append(営農法人.sheet3.KAISETSU_3.Item(d1(0).Item("設定値").ToString).kaisetsu)
            End If

            pXlSheet.Shapes.Item(営農法人.sheet3.KAISETSU).TextEffect.Text = sb.ToString
            pXlSheet.Shapes.Item(営農法人.sheet3.KAISETSU).TextEffect.FontSize = 営農法人.sheet3.KAISETSU_3.Item(d1(0).Item("設定値").ToString).fontPoint


            'グラフの設定
            chartObject = DirectCast(pXlSheet.ChartObjects(営農法人.sheet3.グラフ), Excel.ChartObject)
            chart = chartObject.Chart

        End If



        '系列名の設定
        Dim xlSeries1 As Excel.Series
        Dim xlSeries2 As Excel.Series
        Dim xlSeries3 As Excel.Series
        xlSeries1 = DirectCast(chart.SeriesCollection(1), Excel.Series)
        xlSeries2 = DirectCast(chart.SeriesCollection(2), Excel.Series)
        xlSeries3 = DirectCast(chart.SeriesCollection(3), Excel.Series)
        Dim xlPoints As Excel.Point
        Dim xlDataLabel As Excel.DataLabel

        Select Case (mChosanenKey.Count)
            Case 1
                xlPoints = DirectCast(xlSeries1.Points(7), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(8), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(8), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""

            Case 2, 4
                xlPoints = DirectCast(xlSeries1.Points(10), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(11), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(11), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
            Case 3, 5
                xlPoints = DirectCast(xlSeries1.Points(13), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(14), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(14), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
        End Select


    End Sub

    Private Sub CreateSheet4(pXlSheet As Excel.Worksheet)
        'グラフの設定
        Dim chartObject As Excel.ChartObject = DirectCast(pXlSheet.ChartObjects(営農個人.sheet4.グラフ), Excel.ChartObject)
        Dim chart As Excel.Chart = chartObject.Chart
        ''REV 001 MOD START---
        'Dim max As Integer = 0
        'Dim max2 As Integer = 0
        Dim max As Long = 0
        Dim max2 As Long = 0
        ''REV 001 MOD END---
        Dim start() As String
        Dim d1 As DataRow()
        Dim keinen(2, 5) As Object
        Dim axisY1 As Excel.Axis
        Dim axisY2 As Excel.Axis
        Dim xlSeries1 As Excel.Series
        Dim xlSeries2 As Excel.Series
        Dim xlPoints As Excel.Point
        Dim xlDataLabel As Excel.DataLabel

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '出力範囲と出力データの定義

            d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_営農個人.営個還4_グラフ項目_面積_収入 & "'")

            start = 営農個人.sheet4.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)


            For i As Integer = 0 To start.Count
                If i = 0 Then
                    If Not String.IsNullOrEmpty(d1(1).Item("設定値").ToString) Then
                        keinen(1, i) = getKoumokuName(d1(1).Item("設定値").ToString, Name)
                    End If
                    If Not String.IsNullOrEmpty(d1(0).Item("設定値").ToString) Then
                        keinen(2, i) = getKoumokuName(d1(0).Item("設定値").ToString, Name)
                    End If
                Else
                    keinen(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                    If Not String.IsNullOrEmpty(d1(0).Item("設定値").ToString) Then
                        keinen(1, CInt(start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(d1(1).Item("設定値").ToString).値
                        '最大値を取得
                        Dim value As Decimal
                        If Decimal.TryParse(keinen(1, CInt(start(i - 1))).ToString, value) Then
                            ''REV 001 MOD START---
                            'If max < CInt(keinen(1, CInt(start(i - 1)))) Then
                            'max = CInt(keinen(1, CInt(start(i - 1))))
                            If max < CLng(keinen(1, CInt(start(i - 1)))) Then
                                max = CLng(keinen(1, CInt(start(i - 1))))
                                'REV 001 MOD END---
                            End If
                        End If
                    End If
                    If Not String.IsNullOrEmpty(d1(0).Item("設定値").ToString) Then
                        keinen(2, CInt(start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(d1(0).Item("設定値").ToString).値
                        '最大値を取得
                        Dim value As Decimal
                        If Decimal.TryParse(keinen(2, CInt(start(i - 1))).ToString, value) Then
                            ''REV 001 MOD START---
                            'If max2 < CInt(keinen(2, CInt(start(i - 1)))) Then
                            '    max2 = CInt(keinen(2, CInt(start(i - 1))))
                            If max2 < CLng(keinen(2, CInt(start(i - 1)))) Then
                                max2 = CLng(keinen(2, CInt(start(i - 1))))
                                ''REV 001 MOD END---
                            End If
                        End If
                    End If
                End If
            Next

            Dim range As Excel.Range = Nothing
            range = pXlSheet.Range(営農個人.sheet4.GROUP1)
            range.Value = keinen
            ReleaseComObject(range)
            range = Nothing

            'タイトルの設定
            pXlSheet.Shapes.Item(営農個人.sheet4.経年変化タイトル).TextEffect.Text = getKoumokuName(d1(1).Item("設定値").ToString, Name) & "、" &
                                                                                        getKoumokuName(d1(0).Item("設定値").ToString, Name) & 営農個人.sheet4.経年変化
            'グラフの設定
            chartObject = DirectCast(pXlSheet.ChartObjects(営農個人.sheet4.グラフ), Excel.ChartObject)
            chart = chartObject.Chart
            With chart
                '左軸の修正
                axisY1 = DirectCast(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                '最大値/最小値の設定
                ''REV 001 MOD START---
                'axisY1.MaximumScale = CDbl(max) * 2
                axisY1.MaximumScale = CLng(max) * 2
                ''REV 001 MOD END---
                axisY1.MinimumScale = 0

                '単位が「円」の場合、「万円」とする。
                If getKoumokuName(d1(1).Item("設定値").ToString, tani).Equals("円") Then
                    axisY1.DisplayUnit = Excel.XlDisplayUnit.xlTenThousands
                    axisY1.HasDisplayUnitLabel = False
                    .Shapes.Item(営農個人.sheet4.単位1).TextEffect.Text = "(万円)"
                Else
                    .Shapes.Item(営農個人.sheet4.単位1).TextEffect.Text = "(" & getKoumokuName(d1(1).Item("設定値").ToString, tani) & ")"
                End If

                '右軸の設定
                axisY2 = DirectCast(.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary), Excel.Axis)
                '最大値/最小値の設定
                If max2 < 6 Then
                    '最大値が6未満の場合、6を設定
                    axisY2.MaximumScale = 6
                Else
                    axisY2.MaximumScaleIsAuto = True
                End If
                '最小値は0
                axisY2.MinimumScale = 0

                '単位の設定
                .Shapes.Item(営農個人.sheet4.単位2).TextEffect.Text = "(" & getKoumokuName(d1(0).Item("設定値").ToString, tani) & ")"

            End With

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then

            d1 = mSetteing.Select("項目番号 = '" & ComConst.還元資料.還元資料項目_営農法人.営法還4_グラフ項目_面積_収入 & "'")

            start = 営農法人.sheet4.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)
            max = 0

            For i As Integer = 0 To start.Count
                If i = 0 Then
                    If Not String.IsNullOrEmpty(d1(1).Item("設定値").ToString) Then
                        keinen(1, i) = getKoumokuName(d1(1).Item("設定値").ToString, Name)
                    End If
                    If Not String.IsNullOrEmpty(d1(0).Item("設定値").ToString) Then
                        keinen(2, i) = getKoumokuName(d1(0).Item("設定値").ToString, Name)
                    End If
                Else
                    keinen(0, CInt(start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                    If Not String.IsNullOrEmpty(d1(1).Item("設定値").ToString) Then
                        keinen(1, CInt(start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(d1(1).Item("設定値").ToString).値
                        '最大値を取得
                        Dim value As Decimal
                        If Decimal.TryParse(keinen(1, CInt(start(i - 1))).ToString, value) Then
                            ''REV 001 MOD START---
                            'If max < CInt(keinen(1, CInt(start(i - 1)))) Then
                            '    max = CInt(keinen(1, CInt(start(i - 1))))
                            If max < CLng(keinen(1, CInt(start(i - 1)))) Then
                                max = CLng(keinen(1, CInt(start(i - 1))))
                                ''REV 001 MOD END---
                            End If
                        End If
                    End If
                    If Not String.IsNullOrEmpty(d1(0).Item("設定値").ToString) Then
                        keinen(2, CInt(start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(d1(0).Item("設定値").ToString).値
                        '最大値を取得
                        Dim value As Decimal
                        If Decimal.TryParse(keinen(2, CInt(start(i - 1))).ToString, value) Then
                            ''REV 001 MOD START---
                            'If max2 < CInt(keinen(2, CInt(start(i - 1)))) Then
                            '    max2 = CInt(keinen(2, CInt(start(i - 1))))
                            If max2 < CLng(keinen(2, CInt(start(i - 1)))) Then
                                max2 = CLng(keinen(2, CInt(start(i - 1))))
                                ''REV 001 MOD END---
                            End If
                        End If
                    End If
                End If
            Next

            Dim range As Excel.Range = Nothing
            range = pXlSheet.Range(営農法人.sheet4.GROUP1)
            range.Value = keinen
            ReleaseComObject(range)
            range = Nothing

            'タイトルの設定
            pXlSheet.Shapes.Item(営農法人.sheet4.経年変化タイトル).TextEffect.Text = getKoumokuName(d1(1).Item("設定値").ToString, Name) & "、" &
                                                                                        getKoumokuName(d1(0).Item("設定値").ToString, Name) & 営農法人.sheet4.経年変化
            'グラフの設定
            chartObject = DirectCast(pXlSheet.ChartObjects(営農法人.sheet4.グラフ), Excel.ChartObject)
            chart = chartObject.Chart
            With chart
                '左軸の修正
                axisY1 = DirectCast(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                '最大値/最小値の設定
                ''REV 001 MOD START---
                'axisY1.MaximumScale = CDbl(max) * 2
                axisY1.MaximumScale = CLng(max) * 2
                ''REV 001 MOD END---
                axisY1.MinimumScale = 0

                '単位が「円」の場合、「万円」とする。
                If getKoumokuName(d1(1).Item("設定値").ToString, tani).Equals("円") Then
                    axisY1.DisplayUnit = Excel.XlDisplayUnit.xlTenThousands
                    axisY1.HasDisplayUnitLabel = False
                    .Shapes.Item(営農法人.sheet4.単位1).TextEffect.Text = "(万円)"
                Else
                    .Shapes.Item(営農法人.sheet4.単位1).TextEffect.Text = "(" & getKoumokuName(d1(1).Item("設定値").ToString, tani) & ")"
                End If

                '右軸の設定
                axisY2 = DirectCast(.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary), Excel.Axis)
                '最大値/最小値の設定
                If max2 < 6 Then
                    '最大値が6未満の場合、6を設定
                    axisY2.MaximumScale = 6
                Else
                    axisY2.MaximumScaleIsAuto = True
                End If
                '最小値は0
                axisY2.MinimumScale = 0

                '単位の設定
                .Shapes.Item(営農法人.sheet4.単位2).TextEffect.Text = "(" & getKoumokuName(d1(0).Item("設定値").ToString, tani) & ")"

            End With
        End If

        '系列名の設定
        xlSeries1 = DirectCast(chart.SeriesCollection(1), Excel.Series)
        xlSeries2 = DirectCast(chart.SeriesCollection(2), Excel.Series)

        Select Case (mChosanenKey.Count)
            Case 1
                xlPoints = DirectCast(xlSeries1.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
            Case 2, 4
                xlPoints = DirectCast(xlSeries1.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
            Case 3, 5
                xlPoints = DirectCast(xlSeries1.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
        End Select
    End Sub

    Private Sub CreateSheet5(pXlSheet As Excel.Worksheet)
        Dim i As Integer
        Dim soshueki(4, 5) As Object
        Dim Start() As String
        Dim no1 As String = String.Empty
        Dim no2 As String = String.Empty
        Dim no3 As String = String.Empty
        Dim range As Excel.Range = Nothing

        Dim xlSeries1 As Excel.Series
        Dim xlSeries2 As Excel.Series
        Dim xlSeries3 As Excel.Series
        Dim chartObject As Excel.ChartObject
        Dim chart As Excel.Chart
        Dim xlPoints As Excel.Point
        Dim xlDataLabel As Excel.DataLabel

        Dim shunyuList As List(Of 順位)
        Dim shishutsuList As List(Of 順位)

        Dim chartObjectName1 As String = String.Empty
        Dim chartObjectName2 As String = String.Empty

        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '農業粗収益の経年変化
            shunyuList = createLank(営農個人.sheet5.soshueki)

            Start = 営農個人.sheet5.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)

            If shunyuList(0).値 <> 0 Then
                no1 = getKoumokuName(shunyuList(0).項番, Name)
            End If

            If shunyuList(1).値 <> 0 Then
                no2 = getKoumokuName(shunyuList(1).項番, Name)
            End If

            For i = 0 To Start.Count
                If i = 0 Then
                    soshueki(1, i) = no1
                    soshueki(2, i) = no2
                    soshueki(3, i) = "その他"
                Else
                    soshueki(0, CInt(Start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                    soshueki(1, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(0).項番).値
                    soshueki(2, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(1).項番).値
                    soshueki(3, CInt(Start(i - 1))) = changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農個人.sheet5.農業粗収益).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(0).項番).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(1).項番).値)

                    If CStr(soshueki(3, CInt(Start(i - 1)))) = "0" Then
                        soshueki(3, CInt(Start(i - 1))) = Nothing
                    End If
                    soshueki(4, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農個人.sheet5.農業粗収益).値
                End If
            Next


            range = pXlSheet.Range(営農個人.sheet5.GROUP1)
            range.Value = soshueki
            ReleaseComObject(range)
            range = Nothing

            '農業経営費の経年変化
            Dim keieihi(5, 5) As Object
            shishutsuList = createLank(営農個人.sheet5.keieihi)
            Start = 営農個人.sheet5.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)

            no1 = String.Empty
            no2 = String.Empty
            If shishutsuList(0).値 <> 0 Then
                no1 = getKoumokuName(shishutsuList(0).項番, Name)
            End If

            If shishutsuList(1).値 <> 0 Then
                no2 = getKoumokuName(shishutsuList(1).項番, Name)
            End If

            If shishutsuList(2).値 <> 0 Then
                no3 = getKoumokuName(shishutsuList(2).項番, Name)
            End If


            For i = 0 To Start.Count
                If i = 0 Then
                    keieihi(1, i) = no1
                    keieihi(2, i) = no2
                    keieihi(3, i) = no3
                    keieihi(4, i) = "その他"
                    keieihi(5, i) = "合計"
                Else
                    keieihi(0, CInt(Start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                    keieihi(1, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(0).項番).値
                    keieihi(2, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(1).項番).値
                    keieihi(3, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(2).項番).値
                    keieihi(4, CInt(Start(i - 1))) = changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農個人.sheet5.農業経営費).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(0).項番).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(1).項番).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(2).項番).値)

                    If CStr(keieihi(4, CInt(Start(i - 1)))) = "0" Then
                        keieihi(4, CInt(Start(i - 1))) = Nothing
                    End If
                    keieihi(5, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農個人.sheet5.農業経営費).値
                End If
            Next

            range = Nothing
            range = pXlSheet.Range(営農個人.sheet5.GROUP2)
            range.Value = keieihi
            ReleaseComObject(range)
            range = Nothing

            chartObjectName1 = 営農個人.sheet5.グラフ1
            chartObjectName2 = 営農個人.sheet5.グラフ2
        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '農業収入の経年変化
            Start = 営農法人.sheet5.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)
            shunyuList = createLank(営農法人.sheet5.Shunyu)
            If shunyuList(0).値 <> 0 Then
                no1 = getKoumokuName(shunyuList(0).項番, Name)
            End If

            If shunyuList(1).値 <> 0 Then
                no2 = getKoumokuName(shunyuList(1).項番, Name)
            End If

            For i = 0 To Start.Count
                If i = 0 Then
                    soshueki(1, i) = no1
                    soshueki(2, i) = no2
                    soshueki(3, i) = "その他"
                Else
                    soshueki(0, CInt(Start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                    soshueki(1, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(0).項番).値
                    soshueki(2, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(1).項番).値
                    soshueki(3, CInt(Start(i - 1))) = changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農法人.sheet5.農業収入).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(0).項番).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shunyuList(1).項番).値)
                    If CStr(soshueki(3, CInt(Start(i - 1)))) = "0" Then
                        soshueki(3, CInt(Start(i - 1))) = Nothing
                    End If

                    soshueki(4, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農法人.sheet5.農業収入).値
                End If
            Next

            range = pXlSheet.Range(営農法人.sheet5.GROUP1)
            range.Value = soshueki
            ReleaseComObject(range)
            range = Nothing

            '農業支出の経年変化
            Dim keieihi(5, 5) As Object
            Start = 営農法人.sheet5.経年変化開始位置.Item(mChosanenKey.Count).Split(","c)
            shishutsuList = createLank(営農法人.sheet5.shishutsu)

            no1 = String.Empty
            no2 = String.Empty
            If shishutsuList(0).値 <> 0 Then
                no1 = getKoumokuName(shishutsuList(0).項番, Name)
            End If

            If shishutsuList(1).値 <> 0 Then
                no2 = getKoumokuName(shishutsuList(1).項番, Name)
            End If

            If shishutsuList(2).値 <> 0 Then
                no3 = getKoumokuName(shishutsuList(2).項番, Name)
            End If

            For i = 0 To Start.Count
                If i = 0 Then
                    keieihi(1, i) = no1
                    keieihi(2, i) = no2
                    keieihi(3, i) = no3
                    keieihi(4, i) = "その他"
                    keieihi(5, i) = "合計"
                Else
                    keieihi(0, CInt(Start(i - 1))) = getNengo(CStr(mChosanenKey(mChosanenKey.Count - i)), True)
                    keieihi(1, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(0).項番).値
                    keieihi(2, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(1).項番).値
                    keieihi(3, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(2).項番).値
                    keieihi(4, CInt(Start(i - 1))) = changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農法人.sheet5.農業支出).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(0).項番).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(1).項番).値) -
                                                      changeNum(mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(shishutsuList(2).項番).値)
                    If CStr(keieihi(4, CInt(Start(i - 1)))) = "0" Then
                        keieihi(4, CInt(Start(i - 1))) = Nothing
                    End If
                    keieihi(5, CInt(Start(i - 1))) = mKobetsu(CStr(mChosanenKey(mChosanenKey.Count - i))).Item(営農法人.sheet5.農業支出).値
                End If
            Next

            range = Nothing
            range = pXlSheet.Range(営農法人.sheet5.GROUP2)
            range.Value = keieihi
            ReleaseComObject(range)
            range = Nothing

            chartObjectName1 = 営農法人.sheet5.グラフ1
            chartObjectName2 = 営農法人.sheet5.グラフ2

        End If
        'グラフの設定
        chartObject = DirectCast(pXlSheet.ChartObjects(chartObjectName1), Excel.ChartObject)
        chart = chartObject.Chart


        xlSeries1 = DirectCast(chart.SeriesCollection(1), Excel.Series)
        xlSeries2 = DirectCast(chart.SeriesCollection(2), Excel.Series)
        xlSeries3 = DirectCast(chart.SeriesCollection(3), Excel.Series)

        Select Case (mChosanenKey.Count)
            Case 1
                xlPoints = DirectCast(xlSeries1.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
            Case 2, 4
                xlPoints = DirectCast(xlSeries1.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
            Case 3, 5
                xlPoints = DirectCast(xlSeries1.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
        End Select

        chartObject = Nothing
        xlSeries1 = Nothing
        xlSeries2 = Nothing
        xlSeries3 = Nothing

        'グラフの設定
        chartObject = DirectCast(pXlSheet.ChartObjects(営農個人.sheet5.グラフ2), Excel.ChartObject)
        chart = chartObject.Chart

        '系列名の設定

        xlSeries1 = DirectCast(chart.SeriesCollection(1), Excel.Series)
        xlSeries2 = DirectCast(chart.SeriesCollection(2), Excel.Series)
        xlSeries3 = DirectCast(chart.SeriesCollection(3), Excel.Series)
        Dim xlSeries4 As Excel.Series = DirectCast(chart.SeriesCollection(4), Excel.Series)
        xlPoints = Nothing
        xlDataLabel = Nothing

        Select Case (mChosanenKey.Count)
            Case 1
                xlPoints = DirectCast(xlSeries1.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries4.Points(3), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
            Case 2, 4
                xlPoints = DirectCast(xlSeries1.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries4.Points(4), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
            Case 3, 5
                xlPoints = DirectCast(xlSeries1.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries2.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries3.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
                xlPoints = DirectCast(xlSeries4.Points(5), Excel.Point)
                xlDataLabel = DirectCast(xlPoints.DataLabel, Excel.DataLabel)
                xlDataLabel.ShowSeriesName = True
                xlDataLabel.Separator = "" & Chr(13) & ""
        End Select


        chartObject = Nothing
        xlSeries1 = Nothing
        xlSeries2 = Nothing
        xlSeries3 = Nothing
        xlSeries4 = Nothing
        xlPoints = Nothing
        xlDataLabel = Nothing

    End Sub

    Private Sub CreateSheet6(pXlSheet As Excel.Worksheet)
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '説明文の変更
            Dim ShunyuList As New List(Of 順位)
            Dim shishutsuList As New List(Of 順位)

            Dim range As Excel.Range = Nothing
            range = pXlSheet.Range("B4")
            Dim tmp As String

            Dim seisanHeikin As String = "0"
            Dim Heikin As String = GetShukeiKekkahyoValue(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１, seisanHeikin, SHUKEI_CHIIKI)
            Dim heikinName As String = String.Empty

            If Not IsNothing(Heikin) Then
                heikinName = ComConst.地域.リスト(Heikin).名称
                heikinName = heikinName.Replace("平均", String.Empty)
                heikinName = "と" & heikinName & "平均"
            End If

            tmp = 営農個人.sheet6.ZENBUN

            tmp = tmp.Replace("{0}", getNengo(CStr(mChosanenKey(0)), False))
            tmp = tmp.Replace("{1}", heikinName)

            range.Value = tmp
            ReleaseComObject(range)
            range = Nothing

            '農業粗収益の上位4つの名称を出力
            ShunyuList = createLank(営農個人.sheet6.soshueki)
            Ranking_Title(pXlSheet, ShunyuList, 3, 0, 営農個人.sheet6.SHUEKI_TITLE)

            '農業経営費の上位8つの名称を出力
            shishutsuList = createLank(営農個人.sheet6.keieihi)
            Ranking_Title(pXlSheet, shishutsuList, 7, 0, 営農個人.sheet6.KEIEI_TITLE)

            '前年度の判定
            Dim zennen As Boolean = False
            If mChosanenKey.Count <> 1 Then
                zennen = True
            End If

            SetKangenShiryoYear(pXlSheet, 営農個人.sheet6.THIS_YEAR, 営農個人.sheet6.PREV_YEAR)
            SetShukeiHeader(pXlSheet, 営農個人.sheet6.THIS_YEAR_SYU, 営農個人.sheet6.PREV_YEAR_SYU, 営農個人.sheet6.HEIKIN, 営農個人.sheet6.MENSEKI)

            '個別結果表/集計結果表の内容を出力
            Dim kobetsu(20, 1) As Object
            Dim shukei(20, 1) As Object

            '可変箇所の出力
            Dim i As Integer = 0
            Dim j As Integer = 0

            For i = 営農個人.sheet6.SHUEKI_START To 営農個人.sheet6.SHUEKI_END
                If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(0))).Item(ShunyuList(j).項番).値) Then
                    kobetsu(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(ShunyuList(j).項番))

                    If zennen = True Then
                        If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(1))).Item(ShunyuList(j).項番).値) Then
                            kobetsu(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(ShunyuList(j).項番))
                        End If
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                        shukei(i, 0) = ShukeiDataHenkan(ShunyuList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                        shukei(i, 1) = ShukeiDataHenkan(ShunyuList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                    End If
                End If


                j += 1
            Next

            j = 0
            For i = 営農個人.sheet6.KEIEI_START To 営農個人.sheet6.KEIEI_END
                If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(0))).Item(shishutsuList(j).項番).値) Then
                    kobetsu(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(shishutsuList(j).項番))

                    If zennen = True Then
                        If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(1))).Item(shishutsuList(j).項番).値) Then
                            kobetsu(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(shishutsuList(j).項番))
                        End If
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                        shukei(i, 0) = ShukeiDataHenkan(shishutsuList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                        shukei(i, 1) = ShukeiDataHenkan(shishutsuList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                    End If

                    j += 1
                End If
            Next


            ' 固定項目を代入していく
            Dim d1 As DataRow() = mMaster.Select("シート番号 = 6")
            Dim atai As String
            For Each row As DataRow In d1
                atai = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(row.Item("個別結果表項番").ToString))
                kobetsu(CInt(row.Item("ROW")), 0) = atai

                If zennen = True Then
                    atai = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(row.Item("個別結果表項番").ToString))
                    kobetsu(CInt(row.Item("ROW")), 1) = atai
                End If

                If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                    shukei(CInt(row.Item("ROW")), 0) = ShukeiDataHenkan(row.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                End If

                If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                    shukei(CInt(row.Item("ROW")), 1) = ShukeiDataHenkan(row.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                End If

            Next

            range = Nothing
            range = pXlSheet.Range(営農個人.sheet6.KOBETSU_DATA)
            range.Value = kobetsu
            ReleaseComObject(range)

            range = Nothing
            range = pXlSheet.Range(営農個人.sheet6.SHUKEI_DATA)
            range.Value = shukei
            ReleaseComObject(range)

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            Dim shunyuList As New List(Of 順位)
            Dim seisanGenkaList As New List(Of 順位)
            Dim kanrihiList As New List(Of 順位)

            Dim range As Excel.Range = Nothing
            range = pXlSheet.Range("C4")
            Dim tmp As String

            Dim seisanHeikin As String = "0"
            Dim Heikin As String = GetShukeiKekkahyoValue(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１, seisanHeikin, SHUKEI_CHIIKI)
            Dim heikinName As String = String.Empty

            If Not IsNothing(Heikin) Then
                heikinName = ComConst.地域.リスト(Heikin).名称
                heikinName = heikinName.Replace("平均", String.Empty)
                heikinName = "と" & heikinName & "平均"
            End If

            tmp = 営農法人.sheet6.ZENBUN

            tmp = tmp.Replace("{0}", getNengo(CStr(mChosanenKey(0)), False))
            tmp = tmp.Replace("{1}", heikinName)

            range.Value = tmp
            ReleaseComObject(range)
            range = Nothing

            '農業収入の上位4つの名称を出力
            shunyuList = createLank(営農法人.sheet6.Shunyu)
            Ranking_Title(pXlSheet, shunyuList, 3, 0, 営農法人.sheet6.SHUNYU_TITLE)

            '生産原価の上位6つの名称を出力
            seisanGenkaList = createLank(営農法人.sheet6.seisanGenka)
            Ranking_Title(pXlSheet, seisanGenkaList, 5, 0, 営農法人.sheet6.GENKA_TITLE)

            '管理費および一般管理費の上位2つの名称を出力
            kanrihiList = createLank(営農法人.sheet6.hanbaiKanri)
            Ranking_Title(pXlSheet, kanrihiList, 1, 0, 営農法人.sheet6.KANRI_TITLE)

            '前年度の判定
            Dim zennen As Boolean = False
            If mChosanenKey.Count <> 1 Then
                zennen = True
            End If

            SetKangenShiryoYear(pXlSheet, 営農法人.sheet6.THIS_YEAR, 営農法人.sheet6.PREV_YEAR)
            SetShukeiHeader(pXlSheet, 営農法人.sheet6.THIS_YEAR_SYU, 営農法人.sheet6.PREV_YEAR_SYU, 営農法人.sheet6.HEIKIN, 営農法人.sheet6.MENSEKI)

            '個別結果表/集計結果表の内容を出力
            Dim kobetsu(26, 1) As Object
            Dim shukei(26, 1) As Object

            Dim i As Integer = 0
            Dim j As Integer = 0

            For i = 営農法人.sheet6.SHUNYU_START To 営農法人.sheet6.SHUNYU_END
                If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(0))).Item(shunyuList(j).項番).値) Then
                    kobetsu(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(shunyuList(j).項番))

                    If zennen = True Then
                        If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(1))).Item(shunyuList(j).項番).値) Then
                            kobetsu(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(shunyuList(j).項番))
                        End If
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                        shukei(i, 0) = ShukeiDataHenkan(shunyuList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                        shukei(i, 1) = ShukeiDataHenkan(shunyuList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                    End If
                End If


                j += 1
            Next

            j = 0
            For i = 営農法人.sheet6.GENKA_START To 営農法人.sheet6.GENKA_END
                If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(0))).Item(seisanGenkaList(j).項番).値) Then
                    kobetsu(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(seisanGenkaList(j).項番))

                    If zennen = True Then
                        If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(1))).Item(seisanGenkaList(j).項番).値) Then
                            kobetsu(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(seisanGenkaList(j).項番))
                        End If
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                        shukei(i, 0) = ShukeiDataHenkan(seisanGenkaList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                        shukei(i, 1) = ShukeiDataHenkan(seisanGenkaList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                    End If
                End If
                j += 1
            Next

            j = 0
            For i = 営農法人.sheet6.HANBAIHI_START To 営農法人.sheet6.HANBAIHI_END
                If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(0))).Item(kanrihiList(j).項番).値) Then
                    kobetsu(i, 0) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(kanrihiList(j).項番))

                    If zennen = True Then
                        If Not String.IsNullOrEmpty(mKobetsu(CStr(mChosanenKey(1))).Item(kanrihiList(j).項番).値) Then
                            kobetsu(i, 1) = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(kanrihiList(j).項番))
                        End If
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                        shukei(i, 0) = ShukeiDataHenkan(kanrihiList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                    End If

                    If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                        shukei(i, 1) = ShukeiDataHenkan(kanrihiList(j).項番, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                    End If
                End If
                j += 1
            Next

            ' 固定項目を代入していく
            Dim d1 As DataRow() = mMaster.Select("シート番号 = 6")
            Dim atai As String
            For Each row As DataRow In d1
                atai = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(0))).Item(row.Item("個別結果表項番").ToString))
                kobetsu(CInt(row.Item("ROW")), 0) = atai

                If zennen = True Then
                    atai = ComUtil.KobetsuKekkahyo.GetformattedValue(mKobetsu(CStr(mChosanenKey(1))).Item(row.Item("個別結果表項番").ToString))
                    kobetsu(CInt(row.Item("ROW")), 1) = atai
                End If

                If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１ & "_" & seisanHeikin) Then
                    shukei(CInt(row.Item("ROW")), 0) = ShukeiDataHenkan(row.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_当年_１)
                End If

                If mShukei.ContainsKey(ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年 & "_" & seisanHeikin) Then
                    shukei(CInt(row.Item("ROW")), 1) = ShukeiDataHenkan(row.Item("個別結果表項番").ToString, ComConst.還元資料.集計結果表_明細番号.集計結果表データ_前年)
                End If

            Next

            '合計値を設定
            If changeNum(Convert.ToString(kobetsu(1, 0))) + changeNum(Convert.ToString(kobetsu(8, 0))) + changeNum(Convert.ToString(kobetsu(10, 0))) = 0 Then
                kobetsu(0, 0) = Nothing
            Else
                kobetsu(0, 0) = changeNum(Convert.ToString(kobetsu(1, 0))) + changeNum(Convert.ToString(kobetsu(8, 0))) + changeNum(Convert.ToString(kobetsu(10, 0)))
            End If

            If changeNum(Convert.ToString(kobetsu(1, 1))) + changeNum(Convert.ToString(kobetsu(8, 1))) + changeNum(Convert.ToString(kobetsu(10, 1))) = 0 Then
                kobetsu(0, 1) = Nothing
            Else
                kobetsu(0, 1) = changeNum(Convert.ToString(kobetsu(1, 1))) + changeNum(Convert.ToString(kobetsu(8, 1))) + changeNum(Convert.ToString(kobetsu(10, 1)))
            End If

            If changeNum(Convert.ToString(shukei(1, 0))) + changeNum(Convert.ToString(shukei(8, 0))) + changeNum(Convert.ToString(shukei(10, 0))) = 0 Then
                shukei(0, 0) = Nothing
            Else
                shukei(0, 0) = changeNum(Convert.ToString(shukei(1, 0))) + changeNum(Convert.ToString(shukei(8, 0))) + changeNum(Convert.ToString(shukei(10, 0)))
            End If

            If changeNum(Convert.ToString(shukei(1, 1))) + changeNum(Convert.ToString(shukei(8, 1))) + changeNum(Convert.ToString(shukei(10, 1))) = 0 Then
                shukei(0, 1) = Nothing
            Else
                shukei(0, 1) = changeNum(Convert.ToString(shukei(1, 1))) + changeNum(Convert.ToString(shukei(8, 1))) + changeNum(Convert.ToString(shukei(10, 1)))
            End If

            If changeNum(Convert.ToString(kobetsu(12, 0))) + changeNum(Convert.ToString(kobetsu(25, 0))) + changeNum(Convert.ToString(kobetsu(26, 0))) = 0 Then
                kobetsu(11, 0) = Nothing
            Else
                kobetsu(11, 0) = changeNum(Convert.ToString(kobetsu(12, 0))) + changeNum(Convert.ToString(kobetsu(25, 0))) + changeNum(Convert.ToString(kobetsu(26, 0)))
            End If

            If changeNum(Convert.ToString(kobetsu(12, 1))) + changeNum(Convert.ToString(kobetsu(25, 1))) + changeNum(Convert.ToString(kobetsu(26, 1))) = 0 Then
                kobetsu(11, 1) = Nothing
            Else
                kobetsu(11, 1) = changeNum(Convert.ToString(kobetsu(12, 1))) + changeNum(Convert.ToString(kobetsu(25, 1))) + changeNum(Convert.ToString(kobetsu(26, 1)))
            End If

            If changeNum(Convert.ToString(shukei(12, 0))) + changeNum(Convert.ToString(shukei(25, 0))) + changeNum(Convert.ToString(shukei(26, 0))) = 0 Then
                shukei(11, 0) = Nothing
            Else
                shukei(11, 0) = changeNum(Convert.ToString(shukei(12, 0))) + changeNum(Convert.ToString(shukei(25, 0))) + changeNum(Convert.ToString(shukei(26, 0)))
            End If

            If changeNum(Convert.ToString(shukei(12, 1))) + changeNum(Convert.ToString(shukei(25, 1))) + changeNum(Convert.ToString(shukei(26, 1))) = 0 Then
                shukei(11, 1) = Nothing
            Else
                shukei(11, 1) = changeNum(Convert.ToString(shukei(12, 1))) + changeNum(Convert.ToString(shukei(25, 1))) + changeNum(Convert.ToString(shukei(26, 1)))
            End If

            range = Nothing
            range = pXlSheet.Range(営農法人.sheet6.KOBETSU_DATA)
            range.Value = kobetsu
            ReleaseComObject(range)

            range = Nothing
            range = pXlSheet.Range(営農法人.sheet6.SHUKEI_DATA)
            range.Value = shukei
            ReleaseComObject(range)


        End If

    End Sub

    Private Sub CreateSheet7(xlSheet As Excel.Worksheet)
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then

            '表頭
            SetKangenShiryoYear(xlSheet, 営農個人.sheet7.THIS_YEAR, 営農個人.sheet7.PREV_YEAR)
            SetShukeiHeader(xlSheet, 営農個人.sheet7.THIS_YEAR_SYU, 営農個人.sheet7.PREV_YEAR_SYU, 営農個人.sheet7.HEIKIN, 営農個人.sheet7.MENSEKI)
            '固定項目
            SetKangenShiryoFixed(xlSheet, 7, 1, 14, 営農個人.sheet7.KOBETSU_DATA1, 営農個人.sheet7.SHUKEI_DATA1)
            '任意項目
            SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_営農個人.営個還7_生産概況内訳, 7, 営農個人.sheet7.GROUP1_TITLE, 営農個人.sheet7.KOBETSU_DATA2, 営農個人.sheet7.SHUKEI_DATA2, 2)

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '表頭
            SetKangenShiryoYear(xlSheet, 営農法人.sheet7.THIS_YEAR, 営農法人.sheet7.PREV_YEAR)
            SetShukeiHeader(xlSheet, 営農法人.sheet7.THIS_YEAR_SYU, 営農法人.sheet7.PREV_YEAR_SYU, 営農法人.sheet7.HEIKIN, 営農法人.sheet7.MENSEKI)
            '固定項目
            SetKangenShiryoFixed(xlSheet, 7, 1, 25, 営農法人.sheet7.KOBETSU_DATA1, 営農法人.sheet7.SHUKEI_DATA1)
            '任意項目
            SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_営農法人.営法還7_生産概況内訳, 7, 営農法人.sheet7.GROUP1_TITLE, 営農法人.sheet7.KOBETSU_DATA2, 営農法人.sheet7.SHUKEI_DATA2, 2)
        End If
    End Sub

    Private Sub CreateSheet8(xlSheet As Excel.Worksheet)
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then

            '表頭
            SetKangenShiryoYear(xlSheet, 営農個人.sheet8.THIS_YEAR, 営農個人.sheet8.PREV_YEAR)
            SetShukeiHeader(xlSheet, 営農個人.sheet8.THIS_YEAR_SYU, 営農個人.sheet8.PREV_YEAR_SYU, 営農個人.sheet8.HEIKIN, 営農個人.sheet8.MENSEKI)
            '固定項目
            SetKangenShiryoFixed(xlSheet, 8, 1, 13, 営農個人.sheet8.KOBETSU_DATA1, 営農個人.sheet8.SHUKEI_DATA1)
            '任意項目
            SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_営農個人.営個還8_分析指標, 4, 営農個人.sheet8.GROUP1_TITLE, 営農個人.sheet8.KOBETSU_DATA2, 営農個人.sheet8.SHUKEI_DATA2, 3)

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '表頭
            SetKangenShiryoYear(xlSheet, 営農法人.sheet8.THIS_YEAR, 営農法人.sheet8.PREV_YEAR)
            SetShukeiHeader(xlSheet, 営農法人.sheet8.THIS_YEAR_SYU, 営農法人.sheet8.PREV_YEAR_SYU, 営農法人.sheet8.HEIKIN, 営農法人.sheet8.MENSEKI)
            '固定項目
            SetKangenShiryoFixed(xlSheet, 8, 1, 13, 営農法人.sheet8.KOBETSU_DATA1, 営農法人.sheet8.SHUKEI_DATA1)
            '任意項目
            SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_営農法人.営法還8_分析指標, 4, 営農法人.sheet8.GROUP1_TITLE, 営農法人.sheet8.KOBETSU_DATA2, 営農法人.sheet8.SHUKEI_DATA2, 3)

        End If
    End Sub

    Private Sub CreateSheet9(xlSheet As Excel.Worksheet)
        If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
            '統計表　表頭
            SetKangenShiryoYear(xlSheet, 営農個人.sheet9.THIS_YEAR, 営農個人.sheet9.PREV_YEAR)

            '統計表 任意項目
            SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_営農個人.営個還9_統計表, 18, 営農個人.sheet9.TITLE, 営農個人.sheet9.KOBETSU_DATA)

        ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
            '統計表　表頭
            SetKangenShiryoYear(xlSheet, 営農法人.sheet9.THIS_YEAR, 営農法人.sheet9.PREV_YEAR)

            '統計表 任意項目
            SetKangenShryoOptional(xlSheet, ComConst.還元資料.還元資料項目_営農法人.営法還9_統計表, 29, 営農法人.sheet9.TITLE, 営農法人.sheet9.KOBETSU_DATA)

        End If
    End Sub

    Private Sub CreateSheet10(xlSheet As Excel.Worksheet)
        Dim range As Excel.Range = Nothing

        Try
            If CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_個人 Then
                'この資料等についての連絡先

                Dim d1 As DataRow()
                d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.お問合せ先)

                If d1.Count >= 6 Then
                    range = xlSheet.Range(営農個人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.事務局_拠点))
                    range.Value = d1(0).Item("設定値").ToString + " " + d1(1).Item("設定値").ToString
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農個人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.住所))
                    range.Value = d1(5).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農個人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.電話番号))
                    range.Value = d1(3).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農個人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.FAX))
                    range.Value = d1(4).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農個人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.担当者名))
                    range.Value = d1(2).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing
                End If

            ElseIf CommonInfo.Chosakubun = ComConst.調査区分.営農類型別経営統計_法人 Then
                'この資料等についての連絡先

                Dim d1 As DataRow()
                d1 = mSetteing.Select("項目番号 = " & ComConst.還元資料.お問合せ先)

                If d1.Count >= 6 Then
                    range = xlSheet.Range(営農法人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.事務局_拠点))
                    range.Value = d1(0).Item("設定値").ToString + " " + d1(1).Item("設定値").ToString
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農法人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.住所))
                    range.Value = d1(5).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農法人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.電話番号))
                    range.Value = d1(3).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農法人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.FAX))
                    range.Value = d1(4).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing

                    range = xlSheet.Range(営農法人.sheet10.お問合せ先.Item(ComConst.還元資料.お問合せ先_明細番号.担当者名))
                    range.Value = d1(2).Item("設定値")
                    ReleaseComObject(range)
                    range = Nothing
                End If

            End If
        Finally
            ReleaseComObject(range)
            range = Nothing
        End Try
    End Sub

End Class
