// Learn more about F# at http://fsharp.org

open OpenQA.Selenium
open SpreadsheetAutomator
open System

[<EntryPoint>]
let main argv =
    let mutable i = 0
    let username = argv.[i]
    i <- i + 1
    let password = argv.[i]
    i <- i + 1

    logIn browser username password

    let emailNameContains = argv.[i]
    i <- i + 1
    let fName = argv.[i]
    i <- i + 1
    let lName = argv.[i]
    i <- i + 1
    let alpha = argv.[i]
    i <- i + 1
    let year = argv.[i] |> int
    i <- i + 1
    let city = argv.[i]
    i <- i + 1
    let st = argv.[i]
    i <- i + 1
    let zip = argv.[i]
    i <- i + 1

    let formPng = fillOutForm emailNameContains lName fName alpha year city st zip

    let sheet = argv.[i] = "true"
    i <- i + 1

    let mutable sheetPng = ""
    if sheet then
        let sheetId = argv.[i]
        i <- i + 1
        let sheetName = argv.[i]
        i <- i + 1
        let coord = argv.[i]
        i <- i + 1

        let perform =
            if argv.Length > 5 then
                match argv.[i] with
                | "sendkeys" ->
                    i <- i + 1
                    Perform.SendKeys argv.[i]
                | "click" -> Perform.Click
                | _ -> None
            else
                Perform.None
        sheetPng <- signInOnSheet sheetId sheetName coord perform
        ()
        
    sendConfirmationEmail username formPng sheetPng
    0 // return an integer exit code

