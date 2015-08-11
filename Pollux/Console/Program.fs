

open System.Xml
open System.Xml.Linq
open System.Xml.XPath

open System.IO.Packaging

open Pollux.Log
open Pollux.Excel
open Pollux.Excel.Utils

[<EntryPoint>]
let main argv = 
    printfn "%A" argv

    // https://onedrive.live.com/redir?resid=48FFA0560F4FC7E2!32731&authkey=!ANg55j9a_t8vWdY&ithint=file%2cxlsx
    let ``file6000rows.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows.xlsx"

    try 
        let sheet = LargeSheet ((new ConsoleLogger() :> ILogger), ``file6000rows.xlsx``, "Random", false)
        Pollux.Log.logInfo "%A" sheet.UpperLeft
    with _ -> Pollux.Log.logError "failed LargeSheet '%s'" ``file6000rows.xlsx``
    let x = System.Console.ReadKey() 
    0
    