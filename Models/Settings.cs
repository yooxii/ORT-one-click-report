using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace ORT一键报告.Models
{
    public class SettingItem
    {
        public string Key { get; set; }
        public object Value { get; set; }
        public string Type { get; set; } // "string", "bool", "combobox"
        public List<string> Options { get; set; } = new List<string>(); // 用于combobox
        public ObservableCollection<SettingItem> Children { get; set; } = new ObservableCollection<SettingItem>();
        public bool IsGroup { get; set; } = false;
    }

    public class SettingSection
    {
        public string Name { get; set; }
        public ObservableCollection<SettingItem> Items { get; set; } = new ObservableCollection<SettingItem>();
    }
}
