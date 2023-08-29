using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EML2PST
{
    class MailExtractor
    {
        String mAttPath = "attachments\\";
        public ExtractedMail ProcessEMLFile(FileInfo file)
        {
            
            messages msg = new messages();
            ExtractedMail mailExtractor = new ExtractedMail();
            bool IsEML = true;
            try
            {
                MimeKit.HeaderList loaded = new MimeKit.HeaderList();
                MimeKit.MimeMessage mm = new MimeKit.MimeMessage();
                
                DirectoryInfo attachmentsDirectory = new DirectoryInfo(mAttPath);
                String mAttachmentName = "";
                char[] delimiterChars = { ',' };

                if(!attachmentsDirectory.Exists)
                {
                    attachmentsDirectory.Create();
                }

                loaded = MimeKit.HeaderList.Load(file.FullName);
                mm = MimeKit.MimeMessage.Load(file.FullName);
                var subject = loaded[MimeKit.HeaderId.Subject];
                var sender = loaded[MimeKit.HeaderId.Sender];
                var sent = loaded[MimeKit.HeaderId.DateReceived];
                foreach (var attachment in mm.Attachments)
                {

                    var fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
                    Console.WriteLine(fileName);
                    mAttachmentName = mAttPath + fileName;
                    using (var stream = File.Create(mAttPath + fileName.Replace(":", "").Replace("\\", "")))
                    {
                        if (attachment is MessagePart)
                        {
                            var rfc822 = (MessagePart)attachment;

                            rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment;

                            part.Content.DecodeTo(stream);
                        }
                    }
                }
                mailExtractor.headers = loaded;
                mailExtractor.message = mm;
            } catch (Exception ex)
            {
                msg.msgRED("Exception processing EML file: " + ex.Message);
                IsEML = false;
            }
            mailExtractor.IsEML = IsEML;
            return mailExtractor;
        }

        public void cleanAttachmentsFolder()
        {
            DirectoryInfo attachmentsDirectory = new DirectoryInfo(mAttPath);
            foreach(FileInfo file in attachmentsDirectory.GetFiles())
            {
                file.Delete();
            }

        }
    }
}
