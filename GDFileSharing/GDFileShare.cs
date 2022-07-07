using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Linq;
using System.Text;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace GoogleDriveApp.FileShare
{
    class GDFileShare
    {
        // Private Attributes
        private string[] _Scopes;
        private const string APPLICATION_NAME = "MyAppName";    // Google Drive Application Name (setup on Google Cloud Platform)
        private const string FOLDER_ID = "Folder_ID";           // NOTE: Get string from Google Drive Folder ID, https://drive.google.com/drive/folders/FOLDER-ID?usp=sharing
        private DriveService _DriveService;

        // Custom Constructor
        public GDFileShare()
        {
            _Scopes = new string[] {
                DriveService.Scope.Drive,
                DriveService.Scope.DriveAppdata,
                DriveService.Scope.DriveReadonly,
                DriveService.Scope.DriveFile,
                DriveService.Scope.DriveMetadataReadonly,
                DriveService.Scope.DriveReadonly,
                DriveService.Scope.DriveScripts };

            AuthoriseConnectionToGoogleDrive();
        }

        /// <summary>
        /// Connects to Google Drive according to the OAuth2.0 settings in the Google Drive Account (stored as a local .json file)
        /// <br></br>
        /// NOTE: The .json file is downloaded from APIs and Services => Credentials and contains the ClientID and Secret
        /// </summary>
        private void AuthoriseConnectionToGoogleDrive()
        {
            try
            {
                UserCredential credential;

                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, _Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                    //Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Drive API service.
                _DriveService = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = APPLICATION_NAME,
                });
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Unable to Authorise connection.\n" + e.Message);
            }
        }

        public void GetFile(string fileID)
        {
            FilesResource.GetRequest fileSearchRequest = _DriveService.Files.Get(fileID);
            fileSearchRequest.Fields = "name";

            Google.Apis.Drive.v3.Data.File file = fileSearchRequest.Execute();

            Console.WriteLine("File {0}: {1}", fileID, file.Name);
        }

        /// <summary>
        /// Gets list of all files stored on Google Drive Account (see APPLICATION_NAME)
        /// </summary>
        public void GetFileList(int pageSize)
        {
            // Define parameters of request.
            FilesResource.ListRequest listRequest = _DriveService.Files.List();
            listRequest.PageSize = pageSize;
            listRequest.Fields = "nextPageToken, files(id, name, size, createdTime)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute().Files;
            Console.WriteLine("\nList of Files stored on Google Drive:");
            
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Console.WriteLine("Name: {0}\t ID: ({1}\t Size (kB): {2}\t Uploaded: {3}", file.Name, file.Id, file.Size, file.CreatedTime);
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }
        }

        /// <summary>
        /// Gets a list of folders inside the Google Drive Account (see APPLICATION_NAME)
        /// </summary>
        public void GetFolderList()
        {
            FilesResource.ListRequest folderList = _DriveService.Files.List();
            IList<Google.Apis.Drive.v3.Data.File> folders = folderList.Execute().Files.ToList().Where(x => x.MimeType == "application/vnd.google-apps.folder").ToList();

            Console.WriteLine("\nFound following folders:");
            foreach (var folder in folders)
            {
                Console.WriteLine(folder.Name + "\t[ FolderID: " + folder.Id + " ]");
            }
        }

        /// <summary>
        /// Gets a list of folders and returns the folder name matching the Google Drive Folder ID.
        /// </summary>
        private string GetFolderNameFromID(string id)
        {
            string folderName = "NONE FOUND";
            
            FilesResource.ListRequest folderList = _DriveService.Files.List();
            IList<Google.Apis.Drive.v3.Data.File> folders = folderList.Execute().Files.ToList().Where(x => x.MimeType == "application/vnd.google-apps.folder").ToList();

            foreach (var folder in folders)
            {
                if (folder.Id == id)
                {
                    folderName = folder.Name;
                    break;
                }
            }

            return folderName;
        }

        /// <summary>
        /// Shows the list of files stored in a designated Google Drive folder.
        /// </summary>
        public string ShowFilesInFolder()
        {
            StringBuilder sb = new StringBuilder();

            var request = _DriveService.Files.List();
            request.PageSize = 1000;
            request.Q = string.Format("parents in '{0}'", FOLDER_ID);
            request.Fields = "files(id, name, size, createdTime, mimeType)";
            var results = request.Execute();

            Console.WriteLine("\nFound files in folder: " + GetFolderNameFromID(FOLDER_ID));
            sb.AppendLine("Found files in folder: " + GetFolderNameFromID(FOLDER_ID));
            foreach (var driveFile in results.Files)
            {
                Console.WriteLine("{0} {1} {2} {3}", driveFile.Name, driveFile.MimeType, driveFile.Id, driveFile.CreatedTime);
                sb.AppendLine(string.Format("{0} {1} {2} {3}", driveFile.Name, driveFile.MimeType, driveFile.Id, driveFile.CreatedTime));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Checks the designated Google Drive Folder and automatically deletes any files older than the max days (30 days current setting).
        /// </summary>
        public string DeleteOldFiles()
        {
            const int MAX_DAYS = 30;        
            StringBuilder sb = new StringBuilder();

            var request = _DriveService.Files.List();
            request.PageSize = 1000;
            request.Q = string.Format("parents in '{0}'", FOLDER_ID);
            request.Fields = "files(id, name, size, createdTime, mimeType)";
            var results = request.Execute();

            //Console.WriteLine("\nDelete Old Files:");

            foreach (var driveFile in results.Files)
            {
                sb.AppendLine(string.Format("{0} {1} {2} {3}", driveFile.Name, driveFile.MimeType, driveFile.Id, driveFile.CreatedTime));

                DateTime uploadDate = (DateTime)driveFile.CreatedTime;
                int daysOld = (DateTime.Now.Date - uploadDate).Days;

                if (daysOld >= MAX_DAYS)
                {
                    // Delete File
                    //Console.WriteLine("File: {0} is {1} days old and will be deleted", driveFile.Name, daysOld);
                    sb.AppendLine(string.Format("File: {0} is {1} days old and WILL be deleted", driveFile.Name, daysOld));
                    DeleteFile(driveFile.Id);
                }
                else
                {
                    //Console.WriteLine("File: {0} is {1} days old and will not be deleted", driveFile.Name, daysOld);
                    sb.AppendLine(string.Format("File: {0} is {1} days old and WILL NOT be deleted", driveFile.Name, daysOld));
                }
                sb.AppendLine();
            }

            return sb.ToString(); ;
        }

        public void DeleteFile(string fileId)
        {
            // Attempt to Delete the file
            try
            {
                _DriveService.Files.Delete(fileId).Execute();
                Console.WriteLine("{0} deleted successfully.", fileId);
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: " + e.Message);
            }
        }

        /// <summary>
        /// Uploads a file from the local drive (path is passed in as a string) to a designated Google Drive Folder
        /// </summary>
        public string UploadFile(string uploadFile, string description = "Uploaded with .NET!")
        {
            string mimeType = GetMimeType(uploadFile);
        
            // Check if valid file exists
            if (System.IO.File.Exists(uploadFile))
            {
                // Define Google Drive file attributes
                Google.Apis.Drive.v3.Data.File body = new Google.Apis.Drive.v3.Data.File();
                body.Name = Path.GetFileName(uploadFile);
                body.Description = description;
                body.MimeType = mimeType;
                body.Parents = new List<string> { FOLDER_ID };        
                byte[] byteArray = System.IO.File.ReadAllBytes(uploadFile);
                System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);

                // Attempt to upload file to Google Drive
                try
                {
                    FilesResource.CreateMediaUpload request = _DriveService.Files.Create(body, stream, mimeType);
                    request.SupportsAllDrives = true;
                    request.Fields = "id";
                    request.Fields = "webViewLink";         // Shared link 
                    
                    // Trigger the Upload Request
                    request.Upload();           // Program will wait until Upload is complete

                    Google.Apis.Drive.v3.Data.File file = request.ResponseBody;

                    //Console.WriteLine("{0} [ id: {1} ] uploaded succesfully!", uploadFile, file.Id);
                    //Console.WriteLine("\n Shared Link: {0}", file.WebViewLink);
                    
                    // Trim link string
                    string[] shortenLink = file.WebViewLink.Split('?');

                    return shortenLink[0];
                }
                catch (Exception e)
                {
                    return "ERROR: Unable to upload file.\n" + e.Message;
                }
            }
            else
            {
                return "ERROR: The file does not exist.";
            }
        }

        /// <summary>
        /// Checks the filename against the Windows Registry and determines the MIME Type, returning as a string.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private string GetMimeType(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = System.IO.Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            
            if (regKey != null && regKey.GetValue("Content Type") != null)
            {
                mimeType = regKey.GetValue("Content Type").ToString();
                //Console.WriteLine("Found MIME_TYPE: " + mimeType);
            }

            return mimeType;
        }
    }
}
