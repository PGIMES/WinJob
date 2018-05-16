using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;

namespace WinJob
{
    class fn_FTP
    {
        private string ftpServerIP = String.Empty;
        private string ftpUser = String.Empty;
        private string ftpPassword = String.Empty;
        private string ftpRootURL = String.Empty;

        public fn_FTP(string url, string userid, string password)
        {
            this.ftpServerIP = url;
            this.ftpUser = userid;
            this.ftpPassword = password;
            this.ftpRootURL = "ftp://" + url + "/";
        }

        /// 
        /// 上传 
        /// 
        /// 本地文件绝对路径 
        /// 上传到ftp的路径 
        /// 上传到ftp的文件名 
        public string fileUpload(FileInfo localFile, string ftpPath, string ftpFileName)
        {
            //FileInfo localFile = new FileInfo(localFilepath);

            //bool success = false;
            string success = "";

            FtpWebRequest ftpWebRequest = null;
            FileStream localFileStream = null;
            Stream requestStream = null;
            try
            {
                string uri = ftpRootURL + ftpPath + ftpFileName;
                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.UseBinary = true;
                ftpWebRequest.KeepAlive = false;
                ftpWebRequest.Method = WebRequestMethods.Ftp.UploadFile;
                ftpWebRequest.ContentLength = localFile.Length;
                int buffLength = 2048;
                byte[] buff = new byte[buffLength];
                int contentLen;
                localFileStream = localFile.OpenRead();
                requestStream = ftpWebRequest.GetRequestStream();
                contentLen = localFileStream.Read(buff, 0, buffLength);
                while (contentLen != 0)
                {
                    requestStream.Write(buff, 0, contentLen);
                    contentLen = localFileStream.Read(buff, 0, buffLength);
                }
                success = "";
            }
            catch (Exception ex)
            {
                success = ex.Message;
            }
            finally
            {
                if (requestStream != null)
                {
                    requestStream.Close();
                }
                if (localFileStream != null)
                {
                    localFileStream.Close();
                }
            }
            return success;
        }
        /// 
        /// 上传文件 
        /// 
        /// 本地文件地址(没有文件名) 
        /// 本地文件名 
        /// 上传到ftp的路径 
        /// 上传到ftp的文件名 
        public string fileUpload(string localPath, string localFileName, string ftpPath, string ftpFileName)
        {
            // bool success = false;
            string success = "";
            try
            {
                FileInfo localFile = new FileInfo(localPath + localFileName);
                if (localFile.Exists)
                {
                    success = fileUpload(localFile, ftpPath, ftpFileName);
                }
                else
                {
                    success = "已存在";
                }
            }
            catch (Exception ex)
            {
                success = ex.Message;
            }
            return success;
        }
        /// 
        /// 下载文件 
        /// 
        /// 本地文件地址(没有文件名) 
        /// 本地文件名 
        /// 下载的ftp的路径 
        /// 下载的ftp的文件名 
        public string fileDownload(string localPath, string localFileName, string ftpPath, string ftpFileName)
        {
            // bool success = false;
            string success = "";

            FtpWebRequest ftpWebRequest = null;
            FtpWebResponse ftpWebResponse = null;
            Stream ftpResponseStream = null;
            FileStream outputStream = null;
            try
            {
                outputStream = new FileStream(localPath + localFileName, FileMode.Create);
                string uri = ftpRootURL + ftpPath + ftpFileName;
                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.UseBinary = true;
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse();
                ftpResponseStream = ftpWebResponse.GetResponseStream();
                long contentLength = ftpWebResponse.ContentLength;
                int bufferSize = 2048;
                byte[] buffer = new byte[bufferSize];
                int readCount;
                readCount = ftpResponseStream.Read(buffer, 0, bufferSize);
                while (readCount > 0)
                {
                    outputStream.Write(buffer, 0, readCount);
                    readCount = ftpResponseStream.Read(buffer, 0, bufferSize);
                }
                success = "";
            }
            catch (Exception ex)
            {
                success = ex.Message;
            }
            finally
            {
                if (outputStream != null)
                {
                    outputStream.Close();
                }
                if (ftpResponseStream != null)
                {
                    ftpResponseStream.Close();
                }
                if (ftpWebResponse != null)
                {
                    ftpWebResponse.Close();
                }
            }
            return success;
        }
        /// 
        /// 重命名 
        /// 
        /// ftp文件路径 
        /// 
        /// 
        public string fileRename(string ftpPath, string currentFileName, string newFileName)
        {
            // bool success = false;
            string success = "";
            FtpWebRequest ftpWebRequest = null;
            FtpWebResponse ftpWebResponse = null;
            Stream ftpResponseStream = null;
            try
            {
                string uri = ftpRootURL + ftpPath + currentFileName;
                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.UseBinary = true;
                ftpWebRequest.Method = WebRequestMethods.Ftp.Rename;
                ftpWebRequest.RenameTo = newFileName;
                ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse();
                ftpResponseStream = ftpWebResponse.GetResponseStream();
            }
            catch (Exception ex)
            {
                success = ex.Message;
            }
            finally
            {
                if (ftpResponseStream != null)
                {
                    ftpResponseStream.Close();
                }
                if (ftpWebResponse != null)
                {
                    ftpWebResponse.Close();
                }
            }
            return success;
        }
        /// 
        /// 消除文件 
        /// 
        /// 
        public string fileDelete(string ftpPath, string ftpName)
        {
            // bool success = false;
            string success = "";
            FtpWebRequest ftpWebRequest = null;
            FtpWebResponse ftpWebResponse = null;
            Stream ftpResponseStream = null;
            StreamReader streamReader = null;
            try
            {
                string uri = ftpRootURL + ftpPath + ftpName;
                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(uri));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.KeepAlive = false;
                ftpWebRequest.Method = WebRequestMethods.Ftp.DeleteFile;
                ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse();
                long size = ftpWebResponse.ContentLength;
                ftpResponseStream = ftpWebResponse.GetResponseStream();
                streamReader = new StreamReader(ftpResponseStream);
                string result = String.Empty;
                result = streamReader.ReadToEnd();
                success = "";
            }
            catch (Exception ex)
            {
                success = ex.Message;
            }
            finally
            {
                if (streamReader != null)
                {
                    streamReader.Close();
                }
                if (ftpResponseStream != null)
                {
                    ftpResponseStream.Close();
                }
                if (ftpWebResponse != null)
                {
                    ftpWebResponse.Close();
                }
            }
            return success;
        }
        /// 
        /// 文件存现在检查 
        /// 
        /// 
        /// 
        /// 
        public bool fileCheckExist(string ftpPath, string ftpName)
        {
            bool success = false;
            FtpWebRequest ftpWebRequest = null;
            WebResponse webResponse = null;
            StreamReader reader = null;
            try
            {
                string url = ftpRootURL + ftpPath;
                ftpWebRequest = (FtpWebRequest)FtpWebRequest.Create(new Uri(url));
                ftpWebRequest.Credentials = new NetworkCredential(ftpUser, ftpPassword);
                ftpWebRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                ftpWebRequest.KeepAlive = false;
                webResponse = ftpWebRequest.GetResponse();
                reader = new StreamReader(webResponse.GetResponseStream());
                string line = reader.ReadLine();
                while (line != null)
                {
                    if (line == ftpName)
                    {
                        success = true;
                        break;
                    }
                    line = reader.ReadLine();
                }
            }
            catch (Exception)
            {
                success = false;
            }
            finally
            {
                if (reader != null)
                {
                    reader.Close();
                }
                if (webResponse != null)
                {
                    webResponse.Close();
                }
            }
            return success;
        } 
    }
}
