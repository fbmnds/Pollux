// Weitere Informationen zu F# unter "http://fsharp.net".
// Weitere Hilfe finden Sie im Projekt "F#-Lernprogramm".

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


    let ``file6000rows.xlsx`` = __SOURCE_DIRECTORY__ + @"..\..\UnitTests\data\file6000rows.xlsx"

    do
        let sheet = Sheet ((new ConsoleLogger()), ``file6000rows.xlsx``, "Random", false)
        Pollux.Log.logInfo "%A" sheet.UpperLeft

    0