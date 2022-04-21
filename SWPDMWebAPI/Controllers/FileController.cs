using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using EPDM.Interop.epdm;
using System.Collections;
using System;
using SWPDMWebAPI.Containers;
using System.Drawing;
using System.IO;
using System.Threading;

namespace SWPDMWebAPI.Controllers
{
    [ApiController]
    [Route("file")]
    public class FileController : ControllerBase
    {
        [HttpGet("/checked_out/{folder}")]
        public ActionResult<List<PDMFile>> GetCheckedOutFrom(string folder)
        {
            List<PDMFile> checkedOutFiles = new();

            try
            {
                var thread = new Thread(new ThreadStart(() =>
                {
                    //Setup file vault interface and log into a vault
                    IEdmVault5 vault = new EdmVault5();

                    vault.Login(Environment.GetEnvironmentVariable("User"), Environment.GetEnvironmentVariable("Pwd"), "FTC");

                    //Get the vault's root folder interface
                    List<IEdmFolder5> foldersToCheck = new() { vault.GetFolderFromPath(vault.RootFolderPath + "\\" + folder) };

                    if (foldersToCheck[0] != null)
                    {
                        //Loop recursively through all folders and files
                        while (foldersToCheck.Count > 0)
                        {
                            IEdmFolder5 currentFolder = foldersToCheck[0];

                            IEdmPos5 nextFolderPos = currentFolder.GetFirstSubFolderPosition();

                            while (!nextFolderPos.IsNull)
                            {
                                foldersToCheck.Add(currentFolder.GetNextSubFolder(nextFolderPos));
                            }

                            IEdmPos5 nextFilePos = currentFolder.GetFirstFilePosition();
                            while (!nextFilePos.IsNull)
                            {
                                IEdmFile5 rawFile = currentFolder.GetNextFile(nextFilePos);

                                if (rawFile.IsLocked)
                                {
                                    PDMFile file = BuildFileFromRaw(rawFile);

                                    checkedOutFiles.Add(file);
                                }
                            }

                            foldersToCheck.RemoveAt(0);
                        }
                    }
                }));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                return checkedOutFiles;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return NotFound();
            }
        }

        [HttpGet("{file_name}")]
        public ActionResult<PDMFile> GetFile(string file_name)
        {
            PDMFile file = null;

            try
            {
                var thread = new Thread(new ThreadStart(() =>
                {
                    //Setup file vault interface and log into a vault
                    IEdmVault5 vault = new EdmVault5();

                    vault.Login(Environment.GetEnvironmentVariable("User"), Environment.GetEnvironmentVariable("Pwd"), "FTC");

                    //Get the vault's root folder interface
                    List<IEdmFolder5> foldersToCheck = new() { vault.RootFolder };

                    bool foundFile = false;
                    //Loop recursively through all folders and files
                    while (foldersToCheck.Count > 0 && !foundFile)
                    {
                        IEdmFolder5 currentFolder = foldersToCheck[0];

                        IEdmPos5 nextFolderPos = currentFolder.GetFirstSubFolderPosition();

                        while (!nextFolderPos.IsNull)
                        {
                            foldersToCheck.Add(currentFolder.GetNextSubFolder(nextFolderPos));
                        }

                        IEdmPos5 nextFilePos = currentFolder.GetFirstFilePosition();
                        while (!nextFilePos.IsNull && !foundFile)
                        {
                            IEdmFile5 rawFile = currentFolder.GetNextFile(nextFilePos);

                            if (rawFile.Name == file_name)
                            {
                                file = BuildFileFromRaw(rawFile);

                                foundFile = true;
                            }
                        }

                        foldersToCheck.RemoveAt(0);
                    }
                }));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();

                return file != null ? file : NotFound();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return NotFound();
            }
        }
        public static PDMFile BuildFileFromRaw(IEdmFile5 input_file)
        {
            PDMFile file = new()
            {
                Name = input_file.Name,
                CheckedOut = input_file.IsLocked,
            };

            if (file.CheckedOut)
            {
                file.CheckedOutBy = input_file.LockedByUser.Name;
            }

            IEdmEnumeratorVersion5 enumVersion = (IEdmEnumeratorVersion5)input_file;
            IEdmVersion5 rawVersion = enumVersion.GetVersion(input_file.CurrentVersion);
            VersionInfo version = new();
            version.VersionNo = rawVersion.VersionNo;
            version.DateModified = (DateTime)rawVersion.VersionDate;
            try { version.User = rawVersion.User.Name; }
            catch { version.User = "N/A"; }
            version.Comment = rawVersion.Comment;

            file.Version = version;

            IEdmFile15 file15 = (IEdmFile15)input_file;
            var obj = file15.GetThumbnail2(input_file.CurrentVersion);

            Image imgPreview = PictureDispToImage(obj as stdole.IPictureDisp);

            using (var ms = new MemoryStream())
            {
                imgPreview.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                file.Thumbnail = ms.ToArray();
            }

            return file;
        }
        public static Image PictureDispToImage(stdole.IPictureDisp pictureDisp)
        {
            Image image = null;
            if (pictureDisp != null)
            {
                IntPtr paletteHandle = new(pictureDisp.hPal);
                IntPtr bitmapHandle = new(pictureDisp.Handle);
                image = Image.FromHbitmap(bitmapHandle, paletteHandle);
            }
            return image;
        }
    }
}
