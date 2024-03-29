﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.Net.Http.Headers;
using System.Net;
using System.Reflection;

namespace Satrex.GoogleDrive
{
    public class GoogleDriveInternal
    {
        public static string GOOGLE_MYMETYPE_FOLDER = @"application/vnd.google-apps.folder";
        public static string GOOGLE_MYMETYPE_DOCS= @"application/vnd.google-apps.document";
        public const string GOOGLE_MYMETYPE_PDF = @"application/pdf";
        static string[] Scopes = { DriveService.Scope.Drive, DriveService.Scope.DriveFile};
        static string ApplicationName = "Google Drive Manipulator";

        private static string credentialFilePath = @"Secrets/client_secret.json";
        public static string CredentialFilePath
        {
            get { return credentialFilePath; }
            set { credentialFilePath = value; }
        }
        
        private static DriveService _service;

        public static DriveService GoogleDriveService
        {
            get
            {
                if (_service == null)
                {
                    _service = CreateDriveService(credentialFilePath: CredentialFilePath);
                }
                return _service;
            }
        }

        public static string GetExecutingDirectoryName()
        {
            var location = new Uri(Assembly.GetExecutingAssembly().GetName().CodeBase);
            var locationPath = location.AbsolutePath + location.Fragment;
            var locationDir = new FileInfo(locationPath).Directory;
            return locationDir.FullName;
        }
        private static DriveService CreateDriveService(string credentialFilePath)
        {
            Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
            UserCredential credential;
            //認証プロセス。credPathが作成されていないとBrowserが起動して認証ページが開くので認証を行って先に進む
            var exeDir = GetExecutingDirectoryName();
            var secretFilePath = Path.Combine(exeDir, credentialFilePath);
            using (var stream = new FileStream(secretFilePath, FileMode.Open, FileAccess.Read))
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

        public static IEnumerable<Google.Apis.Drive.v3.Data.File> ListFiles(string folderId)
        {
            // Console.WriteLine(System.Reflection.MethodInfo.GetCurrentMethod().Name + " Start");
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
            System.Diagnostics.Trace.Assert(!string.IsNullOrWhiteSpace(name),
            "name argument not specified.");
            System.Diagnostics.Trace.Assert(!string.IsNullOrWhiteSpace(parentFolderId),"Parent folder Id not specified.");

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

        public static async Task DownloadFileById(string id, string lstrDownloadFile)
        {
            // Create Drive API service.
            _service = GoogleDriveService;
            var file = _service.Files.Get(id).Execute();
            FilesResource.ExportRequest request = _service.Files.Export(file.Id, GOOGLE_MYMETYPE_PDF);
            Console.WriteLine(request.MimeType);
            MemoryStream lobjMS = new MemoryStream();
            await request.DownloadAsync(lobjMS);

            // At this point the MemoryStream has a length of zero?

            lobjMS.Position = 0;
            var lobjFS = new System.IO.FileStream(lstrDownloadFile, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            await lobjMS.CopyToAsync(lobjFS);
        }
        public static async Task DownloadFile(string url, string lstrDownloadFile)
        {
           // Create Drive API service.
            _service = GoogleDriveService;
            // Attempt download
            // Iterate through file-list and find the relevant file
            FilesResource.ListRequest listRequest = _service.Files.List();
            listRequest.Fields = "nextPageToken, files(id, name, mimeType, originalFilename, size)";
            Google.Apis.Drive.v3.Data.File lobjGoogleFile = null;
            var files = listRequest.Execute().Files;
            
            foreach (var item in files)
            {
                if (url.IndexOf(string.Format("id={0}", item.Id)) > -1)
                {
                    Console.WriteLine(string.Format("{0}: {1}", item.OriginalFilename, item.MimeType));
                    lobjGoogleFile = item;
                    break;
                }
            }

            FilesResource.ExportRequest request = _service.Files.Export(lobjGoogleFile.Id, GOOGLE_MYMETYPE_PDF);
            Console.WriteLine(request.MimeType);
            MemoryStream lobjMS = new MemoryStream();
            await request.DownloadAsync(lobjMS);

            // At this point the MemoryStream has a length of zero?

            lobjMS.Position = 0;
            var lobjFS = new System.IO.FileStream(lstrDownloadFile, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            await lobjMS.CopyToAsync(lobjFS);
        }

        public static async Task<byte[]> DownloadFileToByteArray(string url)
        {
           // Create Drive API service.
            _service = GoogleDriveService;
            // Attempt download
            // Iterate through file-list and find the relevant file
            FilesResource.ListRequest listRequest = _service.Files.List();
            listRequest.Fields = "nextPageToken, files(id, name, mimeType, originalFilename, size)";
            Google.Apis.Drive.v3.Data.File lobjGoogleFile = null;
            foreach (var item in listRequest.Execute().Files)
            {
                if (url.IndexOf(string.Format("id={0}", item.Id)) > -1)
                {
                    Console.WriteLine(string.Format("{0}: {1}", item.OriginalFilename, item.MimeType));
                    lobjGoogleFile = item;
                    break;
                }
            }

            FilesResource.ExportRequest request = _service.Files.Export(lobjGoogleFile.Id, GOOGLE_MYMETYPE_PDF);
            Console.WriteLine(request.MimeType);
            using MemoryStream lobjMS = new MemoryStream();
            await request.DownloadAsync(lobjMS);

            // At this point the MemoryStream has a length of zero?
            lobjMS.Position = 0;
            return lobjMS.ToArray();
        }

        public static async Task DownloadFile(string fileId)
        {
            const int KB = 0x400;
            var chunkSize = 256 * KB; // 256KB;

            var fileRequest = GoogleDriveService.Files.Get(fileId);
            fileRequest.Fields = "size";
            var fileResponse = fileRequest.Execute();

            var exportRequest = GoogleDriveService.Files.Export(fileResponse.Id, GOOGLE_MYMETYPE_PDF);
            var client = exportRequest.Service.HttpClient;

            //you would need to know the file size
            var size = fileResponse.Size;

            await using var file = new FileStream(fileResponse.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            file.SetLength((long) size);

            var chunks = (size / chunkSize) + 1;
            for (long index = 0; index < chunks; index++)
            {
                var request = exportRequest.CreateRequest();

                var from = index * chunkSize;
                var to = @from + chunkSize - 1;

                request.Headers.Range = new RangeHeaderValue(@from, to);

                var response = await client.SendAsync(request);

                if (response.StatusCode != HttpStatusCode.PartialContent && !response.IsSuccessStatusCode) continue;

                await using var stream = await response.Content.ReadAsStreamAsync();
                file.Seek(@from, SeekOrigin.Begin);
                await stream.CopyToAsync(file);
            }
        }
        public static Stream ExportFile(string documentFileId, string mymeType)
        {
            return GoogleDriveService.Files.Export(documentFileId, mymeType).ExecuteAsStream();

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
        }
    }
}