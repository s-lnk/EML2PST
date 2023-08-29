using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MimeKit;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EML2PST
{
    public class ExtractedMail
    {
        public MimeKit.HeaderList headers { get; set; }
        public MimeKit.MimeMessage message { get; set; }
        public bool IsEML { get; set; }
    }
    class Program
    {

        

        static void Main(string[] args)
        {
            String sourcePath = "emails\\";

            DirectoryInfo emailsDirectory = new DirectoryInfo(sourcePath);
            messages msg = new messages();
            MailExtractor mailExtractor = new MailExtractor();
            ExtractedMail extractedMail = new ExtractedMail();

            if (!emailsDirectory.Exists)
            {
                emailsDirectory.Create();
            }

            // First process root folder
            foreach (FileInfo file in emailsDirectory.GetFiles())
            {
                msg.msgCOMMON("Process root folder file " + file.Name);
                extractedMail = mailExtractor.ProcessEMLFile(file);
                mail mMail = new mail();
                if (extractedMail.IsEML)
                {
                    if(mMail.uploadMail(extractedMail.headers, extractedMail.message, ""))
                    {
                        file.Delete(); //Remove processed mail
                        mailExtractor.cleanAttachmentsFolder(); //Remove all attachments
                    } else
                    {
                        mailExtractor.cleanAttachmentsFolder(); //Remove attachments
                    }
                } 
            }

            // Process all subfolders
            foreach (DirectoryInfo dir in emailsDirectory.GetDirectories())
            {
                foreach (FileInfo file in dir.GetFiles())
                {
                    msg.msgCOMMON("Process " + dir.Name + " folder file " + file.Name);
                    extractedMail = mailExtractor.ProcessEMLFile(file);
                    mail mMail = new mail();
                    if (extractedMail.IsEML)
                    {
                        if(mMail.uploadMail(extractedMail.headers, extractedMail.message, dir.Name))
                        {
                            file.Delete();
                            mailExtractor.cleanAttachmentsFolder();
                        } else
                        {
                            mailExtractor.cleanAttachmentsFolder(); //Remove attachments
                        }
                    }
                }
            }


            msg.msgGREEN("Done. Pess any key to continue...");
            Console.ReadKey();
        }


    }
}
