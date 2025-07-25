﻿using System;
using System.IO;
using System.Linq;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;

public class GmailAttachmentDownloader
{
    public static void DownloadLatestHoursWeekAttachment(string gmailUser, string appPassword, string downloadFolder, string fileName = null)
    {
        using var client = new ImapClient();
        client.Connect("imap.gmail.com", 993, SecureSocketOptions.SslOnConnect);
        client.Authenticate(gmailUser, appPassword);

        //var inbox = client.Inbox;
        //inbox.Open(MailKit.FolderAccess.ReadOnly);

        var folder = client.GetFolder("Approvals");
        folder.Open(MailKit.FolderAccess.ReadOnly);

        // Search for subjects starting with "FW: Hours Week"
        var query = SearchQuery.SubjectContains("FW: Hours Week");
        var uids = folder.Search(query);

        foreach (var uid in uids.Reverse()) // newest first
        {
            var message = folder.GetMessage(uid);
            if (message.Subject != null && message.Subject.StartsWith("FW: Hours Week"))
            {
                foreach (var attachment in message.Attachments)
                {

                    string writeFileName = string.Empty;
                    if (attachment is MimePart mimePart)
                    {
                        writeFileName = mimePart.FileName;
                    }
                    else if (attachment is MessagePart messagePart)
                    {
                        // MessagePart does not have FileName, use ContentDisposition or ContentType
                        writeFileName = messagePart.ContentDisposition?.FileName ?? messagePart.ContentType.Name;
                    }
                    if (string.IsNullOrEmpty(writeFileName) && !string.IsNullOrEmpty(fileName))
                    {
                        writeFileName = fileName;
                    }
                    if (string.IsNullOrEmpty(writeFileName))
                    {
                        writeFileName = attachment.ContentType.Name ?? "attachment";
                    }
                    if (string.IsNullOrEmpty(writeFileName))
                    {
                        throw new InvalidOperationException("Attachment does not have a valid file name.");
                    }

                    var filePath = Path.Combine(downloadFolder, fileName);
                    using var stream = File.Create(filePath);
                    if (attachment is MessagePart rfc822)
                        rfc822.Message.WriteTo(stream);
                    else if (attachment is MimePart part)
                        part.Content.DecodeTo(stream);

                    Console.WriteLine($"Downloaded: {filePath}");
                }
                break; // Remove this if you want to download from all matching emails
            }
        }

        client.Disconnect(true);
    }
}
