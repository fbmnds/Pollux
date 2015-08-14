namespace Pollux.UnitTests.Excel

module LargeFiles =

    open FsUnit
    open NUnit.Framework

    open Pollux.Excel
    open Pollux.Excel.Utils   
    open Pollux.Excel.Range   

    let ``file6000rows.xlsx``  = __SOURCE_DIRECTORY__ + @"\data\file6000rows.xlsx"
    let ``file6000rows_1.txt`` = __SOURCE_DIRECTORY__ + @"\data\file6000rows_1.txt"

    //let sheetRandom = Sheet (``file6000rows.xlsx``, "Random", false)
(*
    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeFiles : file6000rows.xlsx : UpperLeft``() =
        sheetRandom.UpperLeft |> should equal (Index(0,0))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeFiles : file6000rows.xlsx : LowerRight``() =
        sheetRandom.LowerRight |> should equal (Index(32,7))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeFiles : file6000rows.xlsx : Values``() =
        let i',j' = sheetRandom.UpperLeft.ToTuple
        [ for i in [0 .. sheetRandom.Values2.GetUpperBound(0)] do
              for j in [0 .. sheetRandom.Values2.GetUpperBound(1)] do
                  yield if sheetRandom.Values2.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i') (j+j')) sheetRandom.Values2.[i,j]
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``file6000rows_1.txt``))
*)

    let ``Cost Summary2.xlsx``  = __SOURCE_DIRECTORY__ + @"\data\Cost Summary2.xlsx"
    let ``Cost Summary2_1.txt`` = __SOURCE_DIRECTORY__ + @"\data\Cost Summary2_1.txt"
    let ``Cost Summary2_2.txt`` = __SOURCE_DIRECTORY__ + @"\data\Cost Summary2_2.txt"
    let ``Cost Summary2_3.txt`` = __SOURCE_DIRECTORY__ + @"\data\Cost Summary2_3.txt"

    let sheet  = LargeSheet (``Cost Summary2.xlsx``, "Übersicht", false)
    let sheet2 = LargeSheet (``Cost Summary2.xlsx``, "CheckSums", false)
    let sheet3 = LargeSheet (``Cost Summary2.xlsx``, "CheckSums2", false)

        
    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : UpperLeft : 1``() =
        sheet.UpperLeft |> should equal (Index(0,0))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : UpperLeft : 2``() =
        sheet2.UpperLeft |> should equal (Index(0,1))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : UpperLeft : 3``() =
        sheet3.UpperLeft |> should equal (Index(0,1))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : GetUpperBound(0) : 1``() =
        sheet.Values2.GetUpperBound(0) |> should equal 32

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : GetUpperBound(0) : 2``() =
        sheet2.Values2.GetUpperBound(0) |> should equal 28

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : GetUpperBound(0) : 3``() =
        sheet3.Values2.GetUpperBound(0) |> should equal 28

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : GetUpperBound(1) : 1``() =
        sheet.Values2.GetUpperBound(1) |> should equal 7

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : GetUpperBound(1) : 2``() =
        sheet2.Values2.GetUpperBound(1) |> should equal 7

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : GetUpperBound(1) : 3``() =
        sheet3.Values2.GetUpperBound(1) |> should equal 7

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : LowerRight : 1``() =
        sheet.LowerRight |> should equal (Index(32,7))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : LowerRight : 2``() =
        sheet2.LowerRight |> should equal (Index(28,8))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : LowerRight : 3``() =
        sheet3.LowerRight |> should equal (Index(28,8))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : Values : 1``() =
        let i',j' = sheet.UpperLeft.ToTuple
        let values = sheet.Values ()
        [ for i in [0 .. sheet.Values2.GetUpperBound(0)] do
              for j in [0 .. sheet.Values2.GetUpperBound(1)] do
                  yield if (values i j) <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i') (j+j')) sheet.Values2.[i,j]
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``Cost Summary2_1.txt``))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : Values2 : 1``() =
        let i',j' = sheet.UpperLeft.ToTuple
        [ for i in [0 .. sheet.Values2.GetUpperBound(0)] do
              for j in [0 .. sheet.Values2.GetUpperBound(1)] do
                  yield if sheet.Values2.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i') (j+j')) sheet.Values2.[i,j]
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``Cost Summary2_1.txt``))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : Values : 2``() =
        let i2',j2' = sheet2.UpperLeft.ToTuple
        let values = sheet2.Values ()
        [ for i in [0 .. sheet2.Values2.GetUpperBound(0)] do
              for j in [0 .. sheet2.Values2.GetUpperBound(1)] do
                  yield if sheet2.Values2.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i2') (j+j2')) (values i j);
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``Cost Summary2_2.txt``))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : Values2 : 2``() =
        let i2',j2' = sheet2.UpperLeft.ToTuple
        [ for i in [0 .. sheet2.Values2.GetUpperBound(0)] do
              for j in [0 .. sheet2.Values2.GetUpperBound(1)] do
                  yield if sheet2.Values2.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i2') (j+j2')) sheet2.Values2.[i,j];
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``Cost Summary2_2.txt``))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : Values : 3``() =
        let i3',j3' = sheet3.UpperLeft.ToTuple
        let values = sheet3.Values ()
        [ for i in [0 .. sheet3.Values2.GetUpperBound(0)] do
              for j in [0 .. sheet3.Values2.GetUpperBound(1)] do
                  yield if sheet3.Values2.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i3') (j+j3')) (values i j);
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``Cost Summary2_3.txt``))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : Values2 : 3``() =
        let i3',j3' = sheet3.UpperLeft.ToTuple
        [ for i in [0 .. sheet3.Values2.GetUpperBound(0)] do
              for j in [0 .. sheet3.Values2.GetUpperBound(1)] do
                  yield if sheet3.Values2.[i,j] <> CellContent.Empty 
                        then sprintf "%s %A\r\n" (convertIndex (i+i3') (j+j3')) sheet3.Values2.[i,j];
                        else "" ]
        |> String.concat ""
        |> should equal (System.IO.File.ReadAllText(``Cost Summary2_3.txt``))

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : RangeWithCheckSumsRow : Values : 1``() =
        let sums = 
            [| 13.3672806314049146233M; 13367M;                11141.3622010783753646M; 1597697566.19827133789M; 
                12.488113967731764863M;  13.188791742120878768M; 14.942675417207984297M; 1597722128.54733421646M |]
        let results = [|true; true; true; true; true; true; true; true|]
        let errors : CellIndex [] = [||]
        let range' : Range = 
            {  Name = "Cost Summary2.xlsx : CheckSums"
               DefinedName = None
               UpperLeft  = sheet2.UpperLeft
               LowerRight = sheet2.LowerRight
               Values = sheet2.Values2 }
        RangeWithCheckSumsRow (range')
        |> fun x -> x.CheckSums, x.CheckResults, x.CheckErrors 
        |> fun (x,y,z) -> 
            should (equalWithin 0.000001) sums x,
            should equal results y,
            should equal errors z
        |> ignore



    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : RangeWithCheckSumsRow : Values : 2``() =
        let sums = 
            [|13.2026490463963598333M; 13202M; 10960.4079130749648546M; 1590500864.48966096899M; 0M; 
              12.644524240156997028M; 14.374117698557430477M; 1592168954.47450327896M|] 
        let results = [|true; true; true; true; true; true; true; false|]
        let errors : CellIndex [] = [| Label "I29" |]
        let range' : Range = 
            {  Name = "Cost Summary2.xlsx : CheckSums2"
               DefinedName = None
               UpperLeft  = sheet3.UpperLeft
               LowerRight = sheet3.LowerRight
               Values = sheet3.Values2 }
        RangeWithCheckSumsRow (range')
        |> fun x -> x.CheckSums, x.CheckResults, x.CheckErrors 
        |> fun (x,y,z) -> 
            should (equalWithin 0.000001) sums x,
            should equal results y,
            should equal errors z
        |> ignore

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : RangeWithCheckSumsCol : Values``() =
        let sums = 
            [|1473252699.186540358636202048M; 4977471.99940742972578226M;
              7824012.513421253072052329M; 7197049.16743033179056424M;
              4380179.09636123418458414M; 7994130.72792073246345369M;
              719815.3020090609004359463M; 7668577.43418442105609426M;
              3014968.07541323504888497M; 332669.70036988285268040M;
              7596447.72060111912184819M; 5693153.69312457583938777M;
              7119871.37437327640089722M; 9425186.611950113364771411M;
              2951238.49651129833110113M; 3469163.03635263825066391M;
              1668634.82231510376811296M; 759136.09183924609572114M;
              8177328.23020240588658032M; 472857.95328385717059734M;
              2082448.58280964430615921M; 1305612.394771467962251285M;
              1036179.869346841387631002M; 5835769.90378826703820245M;
              2434518.37491251658884004M; 3249745.30331260332824944M;
              7837939.68817913701354989M; 9245323.19660212314560816M;
              1597722128.547335036840535M|]
        let results = 
            [|true; true; true; true; true; true; true; true; true; true; true; 
              true; true; true; true; true; true; true; true; true; true; true; 
              true; true; true; true; true; true; true|]
        let errors : CellIndex [] = [||]
        let range' : Range = 
            {  Name = "Cost Summary2.xlsx : CheckSums"
               DefinedName = None
               UpperLeft  = sheet2.UpperLeft
               LowerRight = sheet2.LowerRight
               Values = sheet2.Values2 }
        RangeWithCheckSumsCol (range')
        |> fun x -> x.CheckSums, x.CheckResults, x.CheckErrors 
        |> fun (x,y,z) -> 
            should (equalWithin 0.00001) sums x,
            should equal results y,
            should equal errors z
        |> ignore

    [<Test; Category "Pollux.Excel">]
    let ``Excel : LargeSheet : RangeWithCheckSumsCol : Values : conversion``() =
        let sums = 
            [|1473313582.337751886948653268M; 5038355.23136186733680412M; 7884896.291061025649325459M; 
              426188M; 4441062.53729317378630053M; 8055014.22107184628246402M; 780699.2911420647346315033M;
              7729460.50952650150627913M; 3075851.96871464389224959M; 393553.31935607155622756M; 
              7657330.87213877029072848M; 5754037.61264080867521339M; 7180755.26427071320655069M;
              9486070.483799264158219791M; 3012121.52173172267374961M; 3530046.43958584555563794M; 
              1729518.52235097332441411M; 820019.75418832893734622M; 8238211.77612672780749110M;
              533741.17828745101743244M; 2143332.28858016575786625M; 1366496.310243567661732797M; 
              1097063.83722994427445067M; 5896653.42260840181632473M; 2495401.71297560570087879M;
              3310628.69008798425239687M; 7898823.24612222082956598M; 9306206.47861745143270690M; 
              1590525067.118865060073786M|]  
        let results = 
            [|true; true; true; false; true; true; true; true; true; true; true; 
              true; true; true; true; true; true; true; true; true; true; true; 
              true; true; true; true; true; true; true|]
        let errors : CellIndex [] = [| Label "I4" |]
        let range' : Range = 
            {  Name = "Cost Summary2.xlsx : CheckSums2"
               DefinedName = None
               UpperLeft  = sheet3.UpperLeft
               LowerRight = sheet3.LowerRight
               Values = sheet3.Values2 }
        let conversion (i: int) (j: int) x = 
            match x with
            | StringTableIndex _ | InlineString _ | Empty -> 0M
            | Decimal x -> x
            | Date x -> decimal (toJulianDate x)
        RangeWithCheckSumsCol (range', conversion)
        |> fun x -> x.Eps <- 1M; x.CheckSums, x.CheckResults, x.CheckErrors 
        |> fun (x,y,z) -> 
            should (equalWithin 0.000001) sums x,
            should equal results y,
            should equal errors z
        |> ignore
