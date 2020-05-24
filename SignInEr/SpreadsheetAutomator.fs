module SpreadsheetAutomator

open System
open System.Collections.Generic
open System.Drawing
open System.IO
open System.Net
open System.Net.Mail
open System.Text.RegularExpressions
open System.Threading
open OpenQA.Selenium
open OpenQA.Selenium.Chrome
open OpenQA.Selenium.Interactions

let browser =
    let options = ChromeOptions()
    options.AddArguments("no-sandbox", "headless", "disable-gpu")
    new ChromeDriver(ChromeDriverService.CreateDefaultService(), options)

let executor = browser :> IJavaScriptExecutor

type Perform =
    | Click
    | SendKeys of string
    | None

let logIn (browser: ChromeDriver) username password =
    browser.Manage().Window.Size <- Size(1920, 1080)
    browser.Navigate().GoToUrl("https://drive.google.com")
    browser.Keyboard.SendKeys(username + "@usna.edu\n")
    browser.Keyboard.SendKeys(username + "\t" + password + "\t\n")
    Thread.Sleep(10000)

let rec fillOutForm (containsText: string) lName fName alpha year city st zip =
    try
        let cityStZip = city + ", " + st + " " + zip
        browser.Navigate().GoToUrl("https://mail.google.com/")
        (browser.FindElements(By.TagName("td"))
         |> Seq.find (fun elem ->
             elem.GetAttribute("role") <> null && elem.GetAttribute("role") = "gridcell"
             && elem.Text.ToLower().Contains(containsText.ToLower())))
            .Click()
        Thread.Sleep(5000)
        //File.WriteAllBytes("../../../testForm.png", browser.GetScreenshot().AsByteArray)
        let url =
            (browser.FindElements(By.TagName("a"))
             |> Seq.find (fun elem ->
                 elem.GetAttribute("href") <> null && elem.GetAttribute("href").Contains("forms")))
                .GetAttribute("href")
        browser.Navigate().GoToUrl(url)
        let keysForYear =
            match year - 20 with
            | 0 -> Keys.Space
            | _ ->
                [ for _ in 1 .. year - 20 -> Keys.Down ]
                |> String.concat ""

        let keysForMiguel =
            ["\t\t" + lName; "\t" + fName; "\t" + alpha; "\t" + keysForYear; "\t" + cityStZip; "\t" + Keys.Down; "\t\n"]
        let keysForElliott =
            ["\t\t" + lName; "\t" + fName; "\t" + alpha; "\t" + keysForYear; "\t" + cityStZip; "\t" + Keys.Down; "\t" + Keys.Down; Keys.Down + "\t\n"]
        let keysForAlex =
            ["\t\t" + lName; "\t" + fName; "\t" + alpha; "\t" + keysForYear; "\t" + ([for _ in 1 .. 10 -> Keys.Down] |> String.concat ""); "\t" + cityStZip; "\t" + Keys.Down; "\t" + Keys.Down; "\t\n"]
        let keysForJohn =
            ["\t\t" + lName; "\t" + fName; "\t" + alpha; "\t" + cityStZip; "\t" + Keys.Down; "\t" + Keys.Down; "\t\n"]

        for key in keysForMiguel do
            Thread.Sleep(500)
            browser.Keyboard.SendKeys(key)
        let pngFilename = Path.GetTempFileName() + ".png"
        File.WriteAllBytes(pngFilename, browser.GetScreenshot().AsByteArray)
        printfn "Form screenshot saved at: %s\n" pngFilename
        pngFilename
    with
    | :? WebDriverException -> fillOutForm containsText lName fName alpha year city st zip
    | :? KeyNotFoundException ->
        let errMessage = "Invalid email contains string!!! Try again with better arguments."
        printfn "%s\n" errMessage
        errMessage
    | _ ->
        let errMessage = "Unknown exception occurred. Please try again and check your arguments."
        printfn "%s\n" errMessage
        errMessage

let rec signInOnSheet (sheetId: string) (sheetNameContains: string) (coord: string) action =
    try
        let column = ((Char.ToUpper(Regex.Replace(coord, "\d", "") |> char)) |> int) - 64
        let row = (Regex.Replace(coord, "[a-z]+|[A-Z]+", "") |> int) - 1
        let cellWidth = 101
        let cellHeight = 21
        let cellsOffsetX = 46
        let cellsOffsetY = 24 + 27
        let (pixX, pixY) =
            (column * cellWidth + cellsOffsetX - cellWidth / 2, row * cellHeight + cellsOffsetY - cellHeight / 2)

        let docUrl = "https://docs.google.com/spreadsheets/d/" + sheetId + "/edit#gid=0"
        browser.Navigate().GoToUrl(docUrl)
        Thread.Sleep(10000)
        (browser.FindElements(By.ClassName("docs-sheet-tab-name"))
         |> Seq.find (fun elem -> elem.Text.ToLower().Contains(sheetNameContains.ToLower())))
            .Click()
        Thread.Sleep(10000)

        let browserAction = Actions(browser)
        match action with
        | Click ->
            browserAction.MoveToElement(browser.FindElement(By.Id("waffle-grid-container")), pixX, pixY).Click().Build()
                         .Perform()
        | SendKeys str ->
            browserAction.MoveToElement(browser.FindElement(By.Id("waffle-grid-container")), pixX, pixY).Click()
                         .SendKeys(str + "\n").Build().Perform()
        | None -> ()
        let pngFilename = Path.GetTempFileName() + ".png"
        File.WriteAllBytes(pngFilename, browser.GetScreenshot().AsByteArray)
        printfn "Sheet screenshot saved at: %s\n" pngFilename
        pngFilename
    with
    | :? WebDriverException -> signInOnSheet sheetId sheetNameContains coord action
    | :? KeyNotFoundException ->
        let errMessage = "Invalid sheet contains string!!! Try again with better arguments."
        printfn "%s\n" errMessage
        errMessage
    | _ ->
        let errMessage = "Unknown exception occurred. Please try again and check your arguments."
        printfn "%s\n" errMessage
        errMessage

let sendConfirmationEmail username formPng sheetPng =
    let inline isValidFilename filename =
        not (String.IsNullOrEmpty filename) && File.Exists(filename)
    if not (isValidFilename formPng) then
        printfn "Failed to send confirmation email!!! (Form filling failed)"
    else
        try
            let addrFrom = "titom7373@gmail.com"
            let addrTo = username + "@usna.edu"
            let subject = "Sign-in confirmation"
            let body = ""
            let mailMessage = new MailMessage(addrFrom, addrTo, subject, body)

            mailMessage.Attachments.Add(new Attachment(formPng))
            if isValidFilename sheetPng then mailMessage.Attachments.Add(new Attachment(sheetPng))
            let smtp = new SmtpClient("smtp.gmail.com")
            smtp.EnableSsl <- true
            smtp.Credentials <- NetworkCredential("titom7373@gmail.com", "hY6ML%hvL!%okHfbrnp2p28o@D!z9FFk")
            smtp.EnableSsl <- true
            smtp.Send(mailMessage)
            for attachment in mailMessage.Attachments do
                attachment.Dispose()
            mailMessage.Dispose()
            File.Delete(formPng)
            File.Delete(sheetPng)
            printfn "Sent confirmation email to: %s\n" addrTo
        with _ -> printfn "Probably failed to send confirmation email!!!"
