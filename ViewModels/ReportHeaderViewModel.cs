using CommunityToolkit.Mvvm.ComponentModel;
using ORT一键报告.Models;
using System;

namespace ORT一键报告.ViewModels
{
    public class ReportHeaderViewModel : ObservableObject
    {
        private DataCell _testedBy;
        public DataCell TESTED_BY { get => _testedBy; set => SetProperty(ref _testedBy, value); }

        private DataCell _approvedBy;
        public DataCell APPROVED_BY { get => _approvedBy; set => SetProperty(ref _approvedBy, value); }

        private DataCell _projectName;
        public DataCell PROJECT_NAME { get => _projectName; set => SetProperty(ref _projectName, value); }

        private DataCell _testStage;
        public DataCell TEST_STAGE { get => _testStage; set => SetProperty(ref _testStage, value); }

        private DataCell _testDescription;
        public DataCell TestDescription { get => _testDescription; set => SetProperty(ref _testDescription, value); }
        public DataCell Test_Description_Pic { get; set; }
        public DataCell Issue_Photos_Pics { get; set; }
        public DataCell Test_Setup_Pics { get; set; }
        public DataCell Test_ATE_Data { get; set; }
        public DateTime TestStart { get; set; }
        public DateTime TestEnd { get; set; }
        public bool TestPass { get; set; }

        public string[] HeaderInfoList
        {
            get
            {
                string[] res =
                [
                    TESTED_BY.Data,
                    APPROVED_BY.Data,
                    PROJECT_NAME.Data,
                    TEST_STAGE.Data,
                    TestStart.ToString("d"),
                    TestEnd.ToString("d"),
                    TestPass ? "Pass" : "Fail",
                    TestDescription.Data,
                ];
                return res;
            }
        }

        protected void SetPropertyAndNotifyDependents<T>(ref T field, T value, params string[] dependentProperties)
        {
            // 先更新字段并触发自身通知
            if (SetProperty(ref field, value))
            {
                // 再触发所有依赖属性的通知
                foreach (string propName in dependentProperties)
                {
                    OnPropertyChanged(propName);
                }
            }
        }
    }
}