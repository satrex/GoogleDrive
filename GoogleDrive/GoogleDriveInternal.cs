using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Drive;
using Google.Apis.Drive.v3;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using System.Threading;
using Google.Apis.Drive.v3.Data;

namespace Satrex.GoogleDrive
{
    public class GoogleDriveInternal
    {
        public static string GOOGLE_MYMETYPE_FOLDER = @"application/vnd.google-apps.folder";
        public static string GOOGLE_MYMETYPE_DOCS= @"application/vnd.google-apps.document";
        static string[] Scopes = { DriveService.Scope.Drive, DriveService.Scope.DriveFile};
        static string ApplicationName = "Google Drive Manipulator";

        private static DriveService _service;

        public static DriveService GoogleDriveService
        {
            get
            {
                if (_service == null)
                {
                    _service = CreateDriveService();
                }
                return _service;
            }
        }

        private static DriveService CreateDriveService()
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            UserCredential credential;
            //認証プロセス。credPathが作成されていないとBrowserが起動して認証ページが開くので認証を行って先に進む
            using (var stream = new FileStream(@"Secrets/client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine
                    (System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal),
                     ".credentials/drive.googleapis.com-satrex.json");
                //CredentialファイルがcredPathに保存される
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync
                    (GoogleClientSecrets.Load(stream).Secrets, 
                    Scopes, 
                    "user", 
                    CancellationToken.None,
                     new FileDataStore(credPath, true)).Result;
            }
            //API serviceを作成、Requestパラメータを設定
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            return service;
        }

        public static IEnumerable<Google.Apis.Drive.v3.Data.File> ListFiles()
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            // Define parameters of request.
            FilesResource.ListRequest listRequest = GoogleDriveService.Files.List();
            // listRequest.PageSize = 10;
            // listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            var fileList= listRequest.Execute();
            IList<Google.Apis.Drive.v3.Data.File> files = fileList.Files;
            return files;
        }

        public static IEnumerable<Google.Apis.Drive.v3.Data.File> ListFiles(string folderId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            // Define parameters of request.
            FilesResource.ListRequest listRequest = GoogleDriveService.Files.List();
            listRequest.Q = string.Format("'{0}' in parents", folderId);
            listRequest.Fields = "files(id, name, mimeType)";

            // List files.
            var fileList= listRequest.Execute();
            IList<Google.Apis.Drive.v3.Data.File> files = fileList.Files;
            return files;
        }

        public static string CreateFolder(string name, string parentFolderId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            Google.Apis.Drive.v3.Data.File newFile = new Google.Apis.Drive.v3.Data.File();
            newFile.Name = name;
            newFile.MimeType = GOOGLE_MYMETYPE_FOLDER;
            if(!string.IsNullOrWhiteSpace(parentFolderId))
            {
                newFile.Parents = new List<string>() { parentFolderId };
            }
            var createRequest = GoogleDriveService.Files.Create(newFile);
            createRequest.Fields = "id";
            var newFolder = createRequest.Execute();

            Console.WriteLine(string.Format("new Folder ID: {0}", newFolder.Id));
            return newFolder.Id;
        }

        public static string CreateSharedDrive(string name)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");

            Google.Apis.Drive.v3.Data.Drive newDrive = new Google.Apis.Drive.v3.Data.Drive();
            newDrive.Name = name;
            var createRequest = GoogleDriveService.Drives.Create(newDrive, name);
            createRequest.Fields = "ID";
            createRequest.Execute();

            Console.WriteLine(string.Format("new Drive ID: {0}", newDrive.Id));
            return newDrive.Id;
        }

        public static void MoveFile(string fileId, string folderId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            // Retrieve the existing parents to remove
            var getRequest = GoogleDriveService.Files.Get(fileId);
            getRequest.Fields = "parents";
            var file = getRequest.Execute();
            var previousParents = String.Join(",", file.Parents);
            // Move the file to the new folder
            var updateRequest = GoogleDriveService.Files.Update(file, null);
            updateRequest.Fields = "id, parents";
            updateRequest.AddParents = folderId;
            updateRequest.RemoveParents = previousParents;
            file = updateRequest.Execute();
        }

        public static string CopyFile(string fileId, string folderId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            // Retrieve the existing parents to remove

            var getRequest = GoogleDriveService.Files.Get(fileId);
            getRequest.Fields = "name, parents, mimeType";
            var file = getRequest.Execute();
            file.Parents = new List<string>(){folderId};

            // フォルダの場合
            if(file.MimeType == GOOGLE_MYMETYPE_FOLDER)
            {
                return CopyFolderSimply(fileId, folderId);
            }
            // それ以外
            var copyRequest = GoogleDriveInternal.GoogleDriveService.Files.Copy(file, fileId);
            copyRequest.Fields = "id";
            var newFile = copyRequest.Execute();

            return newFile.Id;
        }

        public static string CopyFolderSimply(string originFolderId, string destinationFolderId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            var getRequest = GoogleDriveService.Files.Get(originFolderId);
            getRequest.Fields = "id, name";
            var originFolder = getRequest.Execute();
            var name = originFolder.Name;
            var newId = CreateFolder(name, destinationFolderId);
            return newId;
        }

        public static void CopyFolderDeeply(string originFolderId, string destinationFolderId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            var listRequest = GoogleDriveService.Files.List();
            listRequest.Q = string.Format("'{0}' in parents", originFolderId);
            var list = listRequest.Execute();
            var filesOfFoler = list.Files;
            
            foreach(var file in filesOfFoler)
            {
                var newFile = CopyFile(file.Id, destinationFolderId);
                if(file.MimeType == "application/vnd.google-apps.folder"){
                    CopyFolderDeeply(file.Id, newFile);
                }
            }
        }

        /// <summary>
        /// Pin a revision.
        /// </summary>
        /// <param name="service">Drive API service instance.</param>
        /// <param name="fileId">ID of the file to update revision for.</param>
        /// <param name="revisionId">ID of the revision to update.</param>
        /// <returns>The updated revision, null is returned if an API error occurred</returns>
        public static Revision UpdateRevision(String fileId,
            String revisionId)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            try
            {
                Revision revision = new Revision();
                revision.Id = revisionId;
                return GoogleDriveService.Revisions.Update(revision, fileId, revisionId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
                throw e;
            }
            return null;
        }
    }
}
