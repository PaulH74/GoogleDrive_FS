using System;

namespace GoogleDriveApp.FileShare
{
    class Program
    {
        /// <summary>
        /// Note:
        /// <br></br>
        /// <br></br>
        /// To Upload:
        /// <br></br>
        /// Use command -u "FilePath"
        /// <br></br>
        /// <br></br>
        /// To Delete:
        /// <br></br>
        /// Use command -d
        /// </summary>
        static void Main(string[] args)
        {
            string returnString = "";

            GDFileShare gdfs = new GDFileShare();

            switch (args[0])
            {
                case "-u":
                    // Upload Files
                    returnString = gdfs.UploadFile(args[1]);
                    break;
                case "-d":
                    // Delete Files
                    returnString = gdfs.DeleteOldFiles();
                    break;
                case "-s":
                    // See Files in Folder
                    returnString = gdfs.ShowFilesInFolder();
                    break;
            }

            Console.WriteLine(returnString);
        }
    }
}
