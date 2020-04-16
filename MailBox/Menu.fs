module Menu

open System

let getUserInput (message: string) =
    Console.WriteLine(message)
    Console.ReadLine().Trim()

let logFunc infoType info = Console.WriteLine(infoType + ": " + info)

let logInfo info = logFunc "Info" info

let logError error = logFunc "Error" error

let logInvalidOption = (fun () -> logFunc "Error" "Invalid option")

let menuOptions = [ "Send Mail"; "Read Mail" ]

let printOptions options = options |> List.iteri (fun i x -> Console.WriteLine("{0}. {1}", (i + 1), x))

let getOption n =
    Console.Write("Enter option:")
    List.iter (fun (i: int) ->
        Console.Write((if i = n then "{0}" else "{0},"), i)) [ 1 .. n ]
    Console.WriteLine()
    Console.ReadLine().Trim()

let mainMenu() =
    Console.WriteLine("Menu:")
    menuOptions |> printOptions
    menuOptions.Length |> getOption

let sendMailQuestions =
    [ "Enter valid recipient mailId's(Comma or semicolon Seperated, if multiple):"
      "Enter valid from MailId:"
      "Enter mail subject:"
      "Enter mail body:" ]

let receiveMailQuestions = [ "Enter Server Name:"; "Enter Username:"; "Enter Password" ]

let retrieveAnswers questions =
    seq {
        for question in questions do
            yield getUserInput question
    }
