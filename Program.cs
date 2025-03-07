using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MsgReader.Outlook;
using System.IO;

//"project" is example name
namespace "project".ExtractAttachmentFromMsgFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Extracting attachments from .msg files... ");

            // change the folder location accordingly
            var folderLocation = @"..\personalLoanFiles"; 

            // change where you want to save the attachments
            var foldertarget = @"..\personalLoanFileAttachments\";

            string[] allFiles = Directory.GetFiles(folderLocation, "*.msg", SearchOption.AllDirectories);

            int totalFilesCreated = 0;

            var failedFiles = new List<string>();

            foreach (var filePath in allFiles)
            {
                using (var msg = new Storage.Message(filePath))
                {

                    var outlook = new Storage.Message(filePath, FileAccess.ReadWrite);
                    if (outlook.Attachments.Count > 0)
                    {
                        foreach (Storage.Attachment i in outlook.Attachments)
                        {
                            var fileName = i.FileName;

                            try
                            {
                                using (var fs = new FileStream(foldertarget + fileName, FileMode.Create, FileAccess.Write))
                                {
                                    fs.Write(i.Data, 0, i.Data.Length);
                                    Console.WriteLine($"created new file {fileName}");
                                    totalFilesCreated++;

                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Exception caught in process: {0}", ex);
                                failedFiles.Add(fileName);
                            }

                        }
                    }
                }
            }

            Console.WriteLine($"\nTotal files created: {totalFilesCreated}");
            Console.WriteLine($"\nTotal files Failed: {failedFiles.Count}");
        
            if (failedFiles.Count > 0 )
            {
                int count = 1;
                foreach (var filename in failedFiles)
                {
                    Console.WriteLine($"\nFailed file {count}: {filename}");
                    count++;
                }
            }

            Console.ReadLine();
        }
    }
}
