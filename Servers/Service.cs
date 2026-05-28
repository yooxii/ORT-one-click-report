using Microsoft.Win32;
using System.IO;
using CommunityToolkit.Mvvm.ComponentModel;

namespace ORT一键报告
{
    public class Service : IService
    {
        public string OpenPathDialog(string title, string filter, string initPath, bool isDir)
        {
            OpenFileDialog dialog = new()
            {
                Title = title,
                Filter = filter,
                InitialDirectory = initPath
            };
            bool? result = dialog.ShowDialog();
            return result == true ? isDir ? Path.GetDirectoryName(dialog.FileName) : dialog.FileName : null;
        }
    }
}