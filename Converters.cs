using System.Windows;
using System.Windows.Controls;
using ORT一键报告.ViewModels;

namespace ORT一键报告.Converters
{
    public class SettingTemplateSelector : DataTemplateSelector
    {
        public DataTemplate StringTemplate { get; set; }
        public DataTemplate BoolTemplate { get; set; }
        public DataTemplate ComboTemplate { get; set; }
        public DataTemplate GroupTemplate { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            if (item is SettingItemViewModel vm)
            {
                return vm.Type switch
                {
                    "bool" => BoolTemplate,
                    "combo" => ComboTemplate,
                    "group" => GroupTemplate,
                    _ => StringTemplate
                };
            }
            return base.SelectTemplate(item, container);
        }
    }
}