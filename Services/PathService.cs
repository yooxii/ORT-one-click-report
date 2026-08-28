using Microsoft.Win32;
using System.IO;

namespace ORT一键报告.Services
{
    public class PathService : IPathService
    {
        public string OpenPathDialog(string title, string filter, string initPath, bool isDir)
        {
            if (isDir)
            {
                return OpenFolderDialog(title, initPath);
            }
            OpenFileDialog dialog = new()
            {
                Title = title,
                Filter = filter,
                InitialDirectory = initPath
            };
            bool? result = dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        /// <summary>
        /// 原生 Vista 风格文件夹选择对话框（Ookii.Dialogs.Wpf，BSD 免费库）；
        /// 系统不支持时回退为旧式文件选择框取目录
        /// </summary>
        private static string OpenFolderDialog(string title, string initPath)
        {
            if (Ookii.Dialogs.Wpf.VistaFolderBrowserDialog.IsVistaFolderDialogSupported)
            {
                Ookii.Dialogs.Wpf.VistaFolderBrowserDialog dialog = new()
                {
                    Description = title,
                    UseDescriptionForTitle = true
                };
                if (!string.IsNullOrWhiteSpace(initPath) && Directory.Exists(initPath))
                {
                    dialog.SelectedPath = initPath;
                }
                return dialog.ShowDialog() == true ? dialog.SelectedPath : null;
            }
            // 回退：旧系统用文件选择框选目录内任意文件后取目录
            OpenFileDialog fallback = new()
            {
                Title = title + "（请选择目标文件夹内的任意文件）",
                Filter = "所有文件|*.*",
                InitialDirectory = initPath,
                CheckFileExists = false
            };
            bool? result = fallback.ShowDialog();
            return result == true ? Path.GetDirectoryName(fallback.FileName) : null;
        }

        public string SavePathDialog(string title, string saveName, string filter, string initPath)
        {
            SaveFileDialog dialog = new()
            {
                Title = title,
                FileName = saveName,
                Filter = filter,
                InitialDirectory = initPath
            };
            bool? result = dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }
    }
}
