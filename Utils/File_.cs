using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace ORT一键报告.Utils
{
    internal static class File_
    {

        /// <summary>
        /// 压缩文件夹并支持文件过滤
        /// </summary>
        /// <param name="sourceDirectoryName">要压缩的源文件夹路径</param>
        /// <param name="destinationArchiveFileName">生成的 ZIP 文件路径</param>
        /// <param name="Filter">文件过滤条件</param>
        /// <param name="isInclude"> true 表示保留，false 表示排除</param>
        public static void CreateFilteredZip(string sourceDirectoryName, string destinationArchiveFileName, string Filter = null, bool isInclude = true)
        {
            // 如果目标文件已存在，先删除（避免抛出异常）
            if (File.Exists(destinationArchiveFileName))
            {
                File.Delete(destinationArchiveFileName);
            }

            using var fileStream = new FileStream(destinationArchiveFileName, FileMode.Create);
            // 使用 UTF8 编码防止中文文件名乱码
            using var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, false, Encoding.UTF8);
            var folders = new Stack<string>();
            folders.Push(sourceDirectoryName);

            Regex regex = null;
            if (!string.IsNullOrEmpty(Filter))
            {
                regex = new Regex(Filter, RegexOptions.IgnoreCase);
            }

            while (folders.Count > 0)
            {
                var currentFolder = folders.Pop();

                // 遍历当前文件夹下的所有文件
                foreach (var filePath in Directory.EnumerateFiles(currentFolder))
                {
                    // 执行过滤逻辑
                    if (regex != null)
                    {
                        string fileName = Path.GetFileName(filePath);
                        if (!regex.IsMatch(fileName) ^ !isInclude)
                        {
                            continue; // 不匹配则跳过
                        }
                    }

                    // 计算文件在压缩包中的相对路径
                    string relativePath = Report.GetRelativePath(sourceDirectoryName, filePath);
                    archive.CreateEntryFromFile(filePath, relativePath, System.IO.Compression.CompressionLevel.Optimal);
                }

                // 将子文件夹压入栈中，实现递归
                foreach (var subFolder in Directory.EnumerateDirectories(currentFolder))
                {
                    folders.Push(subFolder);
                }
            }
        }

        public static void CopyFileInfo(string srcFile, string dstPath)
        {
            var srcWriteTime = File.GetLastWriteTime(srcFile);
            var rand = new Random();
            srcWriteTime.AddMinutes(5 + rand.Next(25) + rand.NextDouble());
            File.SetCreationTime(srcFile, srcWriteTime);
            File.SetLastWriteTime(dstPath, srcWriteTime);
        }
    }
}