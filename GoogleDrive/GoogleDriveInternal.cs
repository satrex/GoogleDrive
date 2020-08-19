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

namespace Satrex.GoogleDrive
{
    public class GoogleDriveInternal
    {
        static string[] Scopes = { DriveService.Scope.Drive, DriveService.Scope.DriveFile, DriveService.Scope.DriveAppdata, DriveService.Scope.DriveMetadata, DriveService.Scope.DriveMetadataReadonly, DriveService.Scope.DrivePhotosReadonly, DriveService.Scope.DriveReadonly };
        static string ApplicationName = "Google Drive Manipulator";

        private static DriveService service;

        public static DriveService GoogleDriveService
        {
            get
            {
                if (service == null)
                {
                    service = CreateDriveService();
                }
                return service;
            }
        }

        private static DriveService CreateDriveService()
        {
            UserCredential credential;
            //認証プロセス。credPathが作成されていないとBrowserが起動して認証ページが開くので認証を行って先に進む
            using (var stream = new FileStream(@"Secrets/client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine
                    (System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal),
                     ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
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
            // Define parameters of request.
            FilesResource.ListRequest listRequest = GoogleDriveService.Files.List();
            listRequest.PageSize = 10;
            listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            var fileList= listRequest.Execute();
            IList<Google.Apis.Drive.v3.Data.File> files = fileList.Files;
            return files;
        }

        public static void MoveToFolder(string fileId, string folderId)
        {
            // Retrieve the existing parents to remove
            var getRequest = GoogleDriveService.Files.Get(fileId);
            getRequest.Fields = "parents";
            var file = getRequest.Execute();
            var previousParents = String.Join(",", file.Parents);
            // Move the file to the new folder
            var updateRequest = GoogleDriveService.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
            updateRequest.Fields = "id, parents";
            updateRequest.AddParents = folderId;
            updateRequest.RemoveParents = previousParents;
            file = updateRequest.Execute();
        }
    }
}
