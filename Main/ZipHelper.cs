using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksums;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Main
{
    /// <summary>
    /// Zip壓縮與解壓縮
    /// </summary>
    public class ZipHelper
    {
        /// <summary>
        /// 壓縮多層目錄
        /// </summary>
        /// <param name="strDirectory">要進行壓縮的文件夾</param>
        /// <param name="zipedFile">壓縮后生成的壓縮文件名</param>
        public static void ZipFileDirectory(string strDirectory)
        {
            string zipedFile = string.Format("{0}.docx", strDirectory);

            using (System.IO.FileStream ZipFile = System.IO.File.Create(zipedFile))
            {
                using (ZipOutputStream s = new ZipOutputStream(ZipFile))
                {
                    ZipSetp(strDirectory, s, "");
                }
            }
        }

        /// <summary>
        /// 遞歸遍歷目錄
        /// </summary>
        /// <param name="strDirectory">要進行壓縮的文件夾</param>
        /// <param name="s">The ZipOutputStream Object.</param>
        /// <param name="parentPath">The parent path.</param>
        private static void ZipSetp(string strDirectory, ZipOutputStream s, string parentPath)
        {
            if (strDirectory[strDirectory.Length - 1] != Path.DirectorySeparatorChar)
            {
                strDirectory += Path.DirectorySeparatorChar;
            }
            
            string[] filenames = Directory.GetFileSystemEntries(strDirectory);
            foreach (string file in filenames)// 遍历所有的文件和目录
            {
                if (Directory.Exists(file))// 先当作目录处理如果存在这个目录就递归Copy该目录下面的文件
                {
                    string pPath = parentPath;
                    pPath += file.Substring(file.LastIndexOf("\\") + 1);
                    pPath += "\\";
                    ZipSetp(file, s, pPath);
                }
                else // 否则直接压缩文件
                {
                    //打开压缩文件
                    using (FileStream fs = File.OpenRead(file))
                    {
                        byte[] buffer = new byte[fs.Length];
                        fs.Read(buffer, 0, buffer.Length);
                        string fileName = parentPath + file.Substring(file.LastIndexOf("\\") + 1);
                        ZipEntry entry = new ZipEntry(fileName);
                        entry.DateTime = DateTime.Now;
                        entry.Size = fs.Length;
                        fs.Close();                       
                        s.PutNextEntry(entry);
                        s.Write(buffer, 0, buffer.Length);
                    }
                }
            }
        }        
        
        /// <summary>
        /// 解壓縮文件。
        /// </summary>
        /// <param name="pi_sTargetFile">目標文件路徑。</param>               
        public static string Decompress(string pi_sTargetFile)
        {
            string sReturn = string.Format("{0}\\{1}", System.IO.Path.GetDirectoryName(pi_sTargetFile), DateTime.Now.ToString("yyMMdd-hhmmss"));

            if (!File.Exists(pi_sTargetFile))
            {
                throw new System.IO.FileNotFoundException("指定要解壓縮的文件: " + pi_sTargetFile + " 不存在!");
            }            

            Directory.CreateDirectory(sReturn);

            using (ZipInputStream objZipInputStream = new ZipInputStream(File.OpenRead(pi_sTargetFile)))
            {
                ZipEntry objZipEntry = null;

                while ((objZipEntry = objZipInputStream.GetNextEntry()) != null)
                {   
                    string sEntryDirectory = Path.GetDirectoryName(objZipEntry.Name);
                    string sEntryFileName = Path.GetFileName(objZipEntry.Name);
                    
                    if (sEntryDirectory.Length > 0)
                    {
                        Directory.CreateDirectory(string.Format("{0}\\{1}", sReturn, sEntryDirectory));
                    }

                    if (sEntryFileName != String.Empty)
                    {
                        using (FileStream streamWriter = File.Create(string.Format("{0}\\{1}",sReturn, objZipEntry.Name)))
                        {
                            int size = 2048;
                            byte[] data = new byte[2048];

                            while (true)
                            {
                                size = objZipInputStream.Read(data, 0, data.Length);
                                if (size > 0)
                                {
                                    streamWriter.Write(data, 0, size);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            return sReturn;
        }       
    }
}