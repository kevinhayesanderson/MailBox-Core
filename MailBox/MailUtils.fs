module MailUtils

open System
open System.Text.RegularExpressions
open System.Globalization
open System.Net.Mail
open DnsClient
open System.Linq
open MailKit.Net.Imap
open MailKit.Net.Pop3
open MailKit
open MimeKit
open Logary
open Logary.Message

type ServerName =
    | IMAP
    | POP3
    | SMPT
    | None

type ServerType =
    | Incoming
    | Outgoing

type AuthenticationType =
    | AUTH
    | StartTLS
    | SSL
    | Other

type DefaultPortInfo =
    { Name: ServerName
      Type: ServerType
      Authentication: AuthenticationType
      Port: int }

let defaultServerInfo =
    ([ { Name = SMPT
         Type = Outgoing
         Authentication = AUTH
         Port = 25 }
       { Name = SMPT
         Type = Outgoing
         Authentication = AUTH
         Port = 26 }
       { Name = SMPT
         Type = Outgoing
         Authentication = StartTLS
         Port = 587 }
       { Name = SMPT
         Type = Outgoing
         Authentication = SSL
         Port = 465 }
       { Name = POP3
         Type = Incoming
         Authentication = AUTH
         Port = 110 }
       { Name = POP3
         Type = Incoming
         Authentication = SSL
         Port = 995 }
       { Name = IMAP
         Type = Incoming
         Authentication = AUTH
         Port = 143 }
       { Name = IMAP
         Type = Incoming
         Authentication = SSL
         Port = 993 } ])

let timeout = 9000

let logger = Log.create "MailUtils"

let logInfo info = event Info info |> Logger.logSimple logger

let logError errorInfo = event Error errorInfo |> Logger.logSimple logger

let getPort (name: ServerName) (auth: AuthenticationType) =
    defaultServerInfo.FirstOrDefault(fun si -> si.Name.Equals(name) && (si.Authentication.Equals(auth))).Port

let mutable isMultipleToAddresses = false

let mutable toAddresses: InternetAddress list = []

let mutable invalidmailIds: string list = []

let isValidEmail mailId =
    if String.IsNullOrWhiteSpace(mailId) then
        false
    else
        try
            let domainMapper (mt: Match) =
                let idn = IdnMapping()
                let domainName = idn.GetAscii(mt.Groups.[2].Value)
                mt.Groups.[1].Value + domainName

            let mailId =
                Regex.Replace
                    (mailId.Trim(), @"(@)(.+)$", domainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(200.0))
            Regex.IsMatch
                (mailId,
                 @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))"
                 + @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                 RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250.0))
        with
        | :? RegexMatchTimeoutException -> false
        | :? ArgumentException -> false

let validateMailIds mailIds = List.partition isValidEmail mailIds

let isValidMailAddress (mailAddress: string) =
    match mailAddress.Contains(',') || mailAddress.Contains(';') with
    | true ->
        let (validMailIds, invalidMailIds) = validateMailIds (List.ofArray (mailAddress.Split(',', ';')))
        for mailId in validMailIds do
            toAddresses <- InternetAddress.Parse(mailId) :: toAddresses
        isMultipleToAddresses <- true
        match (invalidMailIds).Any() with
        | true ->
            invalidmailIds <- invalidmailIds @ invalidMailIds
            false
        | _ -> true
    | _ ->
        match isValidEmail mailAddress with
        | true -> true
        | _ ->
            invalidmailIds <- mailAddress :: invalidmailIds
            false

let getDomainName (mailId: string) = (mailId |> MailAddress).Host

let getDomainNamesFor (mailIds: string) =
    match isMultipleToAddresses with
    | true ->
        let mutable domainNames = []
        List.iter (fun mailId -> domainNames <- (getDomainName mailId) :: domainNames)
            (List.ofArray (mailIds.Split(',', ';')))
        (domainNames
         |> Seq.distinct
         |> List.ofSeq).FirstOrDefault()
    | false -> getDomainName mailIds

let getMXRecord domainName = (LookupClient()).Query(domainName, QueryType.MX).Answers.FirstOrDefault()

let getSMPTServerName (record: Protocol.DnsResourceRecord) =
    match record with
    | :? Protocol.MxRecord as mxRecord -> mxRecord.Exchange.Original.TrimEnd('.')
    | _ -> ""

let setUpMailMessage (toAddress: string) fromAddress subject messageBody =
    let message = MimeMessage()
    message.From.Add(MailboxAddress(fromAddress, fromAddress))
    if isMultipleToAddresses
    then message.To.AddRange(toAddresses)
    else message.To.Add(InternetAddress.Parse(toAddress))
    message.Subject <- subject
    let body = TextPart("plain")
    body.Text <- messageBody
    message.Body <- body
    message

let setUpSmptClient =
    let client = new Net.Smtp.SmtpClient()
    client.Timeout <- timeout
    client.Connected
    |> Event.add (fun args ->
        String.Format
            ("SMPT Client Connected: Host-{0} Port-{1} SecureSocketOption-{2}", args.Host, args.Port, args.Options)
        |> logInfo)
    client.Disconnected |> Event.add (fun _ -> logInfo "SMPT Client Disconnected")
    client.MessageSent |> Event.add (fun args -> logInfo args.Response)
    client

let send (client: Net.Smtp.SmtpClient) mailId message =
    let host =
        mailId
        |> getDomainNamesFor
        |> getMXRecord
        |> getSMPTServerName

    let port =
        SMPT
        |> getPort
        <| AUTH
    client.Connect(host, port, Security.SecureSocketOptions.None)
    let options = FormatOptions.Default.Clone()
    if (client.Capabilities.HasFlag(Net.Smtp.SmtpCapabilities.UTF8)) then options.International <- true
    String.Format("Sending an mail to {0} ", mailId) |> logInfo
    client.Send(options, message)
    client.Disconnect(true)

let sendMail toAddress fromAddress subject messageBody =
    try
        match (isValidMailAddress toAddress) && (isValidMailAddress fromAddress) with
        | true ->
            let message = setUpMailMessage toAddress fromAddress subject messageBody
            let smtpClient = setUpSmptClient
            match isMultipleToAddresses with
            | true ->
                for mailid in toAddresses do
                    send smtpClient (mailid.ToString()) message
            | false -> send smtpClient toAddress message
            smtpClient.Dispose()
        | false ->
            "Invalid mailid's" |> logError
            for mailId in invalidmailIds do
                mailId |> logError
    with ex -> ex.Message |> logError

let printMails client =
    match box client with
    | :? ImapClient as imapClient ->
        for summary in imapClient.Inbox.Fetch(1, imapClient.Inbox.Count - 1, MessageSummaryItems.Full) do
            Console.WriteLine("[summary] {0:D2}: {1}", summary.Index, summary.Envelope.Subject)
        Console.WriteLine()
        Console.WriteLine("Printed all {0} mails", imapClient.Inbox.Count)
    | :? Pop3Client as pop3Client ->
        for i = 1 to pop3Client.Count do
            let message = pop3Client.GetMessage(i)
            Console.WriteLine("Subject: {0}", message.Subject)
        Console.WriteLine("Printed all {0} mails", pop3Client.Count)
    | _ -> Console.WriteLine("No Mail Found")

let imapPrintMailsCallback =
    (fun (imapClient: ImapClient) ->
        let inbox = imapClient.Inbox
        inbox.Open(FolderAccess.ReadOnly) |> ignore
        Console.WriteLine("Total messages: {0}", inbox.Count)
        printMails imapClient
        imapClient.Disconnect(true)
        imapClient.Dispose())

let setUpIMAPClient serverName (userName: string) password printMailsCallback =
    let imapClient = new ImapClient()
    imapClient.Timeout <- timeout
    imapClient.Connected
    |> Event.add
        (fun args ->
            Console.WriteLine
                ("IMAP Client Connected: Host-{0} Port-{1} SecureSocketOption-{2}", args.Host, args.Port, args.Options))
    imapClient.ConnectAsync(serverName,
                            IMAP
                            |> getPort
                            <| SSL, true).GetAwaiter().GetResult()
    imapClient.Authenticated
    |> Event.add (fun args ->
        Console.WriteLine(args.Message)
        printMailsCallback imapClient)
    imapClient.AuthenticateAsync(userName, password).GetAwaiter().GetResult()

let pop3PrintMailsCallback =
    (fun (pop3Client: Pop3Client) ->
        Console.WriteLine("Total messages: {0}", pop3Client.Count)
        printMails pop3Client
        pop3Client.Disconnect(true)
        pop3Client.Dispose())

let setUpPOP3Client serverName (userName: string) password printMailsCallback =
    let pop3Client = new Pop3Client()
    pop3Client.Timeout <- timeout
    pop3Client.Connected
    |> Event.add
        (fun args ->
            Console.WriteLine
                ("POP3 Client Connected: Host-{0} Port-{1} SecureSocketOption-{2}", args.Host, args.Port, args.Options))
    pop3Client.ConnectAsync(serverName,
                            POP3
                            |> getPort
                            <| SSL, true).GetAwaiter().GetResult()
    pop3Client.Authenticated
    |> Event.add (fun args ->
        Console.WriteLine(args.Message)
        printMailsCallback pop3Client)
    pop3Client.AuthenticateAsync(userName, password).GetAwaiter().GetResult()

let getServerType (serverName: string) =
    match serverName with
    | serverName when serverName.Contains("IMAP") || serverName.Contains("imap") || serverName.Contains("Imap") -> IMAP
    | serverName when serverName.Contains("POP") || serverName.Contains("pop") -> POP3
    | _ -> None

let receiveMail serverName userName password =
    try
        match getServerType (serverName) with
        | IMAP -> setUpIMAPClient serverName userName password imapPrintMailsCallback
        | POP3 -> setUpPOP3Client serverName userName password pop3PrintMailsCallback
        | _ -> "Invalid Servertype" |> logError
    with ex -> ex.Message |> logError
