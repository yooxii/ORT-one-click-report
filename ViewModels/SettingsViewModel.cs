using CommunityToolkit.Mvvm.ComponentModel;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace ORT一键报告.ViewModels
{
    // 暴露给其他控件使用的数据结构
    public class SettingItemViewModel : ObservableObject
    {
        private object _value;
        public string Key { get; set; }
        public string Label { get; set; }
        public string Type { get; set; } // "string", "bool", "combo"
        public List<string> Options { get; set; } // Combo专用选项

        public object Value
        {
            get => _value;
            set
            {
                if (_value != value)
                {
                    _value = value;
                    OnPropertyChanged();
                    ValueChanged?.Invoke(this, EventArgs.Empty); // 触发值改变事件
                }
            }
        }

        public ObservableCollection<SettingItemViewModel> Children { get; set; } = [];

        public event EventHandler ValueChanged;

        public static Dictionary<string, object> GetDictionary(ObservableCollection<SettingItemViewModel> items)
        {
            Dictionary<string, object> res = [];
            foreach (SettingItemViewModel item in items)
            {
                if (item.Type == "group" && item.Children.Count != 0)
                    res[item.Label] = GetDictionary(item.Children);
                else
                    res[item.Label] = item.Value;
            }
            return res;
        }
    }

    public class SettingsViewModel(string savePath = "") : ObservableObject
    {
        private readonly string _savePath = savePath;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        // 暴露给外部读取的设置数据字典
        public Dictionary<string, object> SettingsData { get; } = [];

        // 绑定到 UI 的根节点集合
        public ObservableCollection<SettingItemViewModel> RootItems { get; } = [];

        public static Dictionary<string, object> ParseJson(string json)
        {

            JObject jObj = JObject.Parse(json);
            Dictionary<string, object> result = [];

            foreach (var prop in jObj)
            {
                result[prop.Key] = prop.Value.Type == JTokenType.Object ?
                    prop.Value.ToObject<Dictionary<string, object>>() :
                    prop.Value.ToString();
            }
            return result;
        }

        /// <summary>
        /// 核心方法：传入 JSON 字符串进行解析并生成界面树
        /// </summary>
        public void LoadFromJson(string json)
        {
            RootItems.Clear();
            SettingsData.Clear();

            try
            {
                var jObject = JObject.Parse(json);
                ParseNode(jObject, RootItems, "");
            }
            catch (Exception ex)
            {
                _logger.Error($"JSON 解析失败: {ex.Message}");
            }
        }

        public void LoadFromFile(string filePath)
        {
            var content = File.ReadAllText(filePath);
            LoadFromJson(content);
        }

        /// <summary>
        /// 递归解析 JSON 节点
        /// </summary>
        private void ParseNode(JToken token, ObservableCollection<SettingItemViewModel> parentCollection, string parentPath)
        {
            foreach (var property in ((JObject)token).Properties())
            {
                string currentKey = property.Name;
                string fullPath = string.IsNullOrEmpty(parentPath) ? currentKey : $"{parentPath}.{currentKey}";
                JToken value = property.Value;

                var item = new SettingItemViewModel { Key = fullPath };

                // 判断是否为嵌套对象（分组/分类）
                if (value.Type == JTokenType.Object && !IsSettingSchema(value))
                {
                    //item.Label = FormatLabel(currentKey);
                    item.Label = currentKey;
                    item.Type = "group";
                    ParseNode(value, item.Children, fullPath);

                    // 订阅子项的值改变事件，以便触发整体保存
                    foreach (var child in item.Children)
                        SubscribeToValueChange(child);
                }
                else
                {
                    // 叶子节点（具体的设置项）
                    ParseLeafNode(value, item, fullPath);
                    SubscribeToValueChange(item);
                }

                parentCollection.Add(item);
            }
        }

        /// <summary>
        /// 解析叶子节点（string, bool, combo）
        /// </summary>
        private void ParseLeafNode(JToken value, SettingItemViewModel item, string fullPath)
        {
            if (value.Type == JTokenType.Object)
            {
                // 格式预期: { "type": "combo", "label": "主题", "default": "dark", "options": ["light", "dark"] }
                var schema = (JObject)value;
                item.Type = schema["type"]?.ToString() ?? "string";
                var keyParts = item.Key.Split('.');
                //item.Label = schema["label"]?.ToString() ?? FormatLabel(keyParts[keyParts.Length - 1]);
                item.Label = schema["label"]?.ToString() ?? keyParts[keyParts.Length - 1];

                var defaultVal = schema["default"];
                item.Value = GetTypedValue(defaultVal, item.Type);

                if (item.Type == "combo" && schema["options"] is JArray optionsArr)
                {
                    item.Options = optionsArr.ToObject<List<string>>();
                }
            }
            else
            {
                // 简单格式预期: "editor.fontSize": 14 或 "editor.wordWrap": true
                var keyParts = item.Key.Split('.');
                //item.Label = FormatLabel(keyParts[keyParts.Length - 1]);
                item.Label = keyParts[keyParts.Length - 1];
                item.Type = value.Type == JTokenType.Boolean ? "bool" : "string";
                item.Value = GetTypedValue(value, item.Type);
            }

            // 同步到全局数据字典
            SettingsData[fullPath] = item.Value;
        }

        /// <summary>
        /// 判断一个 Object 是否是设置项的描述（包含 type 字段），还是普通的嵌套分组
        /// </summary>
        private bool IsSettingSchema(JToken token)
        {
            return token["type"] != null;
        }

        private object GetTypedValue(JToken token, string type)
        {
            if (token == null || token.Type == JTokenType.Null) return null;
            return type switch
            {
                "bool" => token.Value<bool>(),
                "combo" => token.ToString(),
                _ => token.ToString(),
            };
        }

        private string FormatLabel(string key)
        {
            // 简单的驼峰转空格处理，例如 fontSize -> Font Size
            return System.Text.RegularExpressions.Regex.Replace(key, "(\\B[A-Z])", " $1");
        }

        /// <summary>
        /// 订阅值改变事件，实现实时保存
        /// </summary>
        private void SubscribeToValueChange(SettingItemViewModel item)
        {
            item.ValueChanged += (s, e) =>
            {
                SettingsData[item.Key] = item.Value;
                SaveSettings();
            };

            // 如果有子节点，继续向下订阅
            foreach (var child in item.Children)
                SubscribeToValueChange(child);
        }

        /// <summary>
        /// 实时保存到文件
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(SettingsData, Newtonsoft.Json.Formatting.Indented);
                if (!string.IsNullOrEmpty(_savePath))
                    File.WriteAllText(_savePath, json);
            }
            catch (Exception ex)
            {
                _logger.Error($"保存设置失败: {ex.Message}");
                return;
            }
        }
    }
}