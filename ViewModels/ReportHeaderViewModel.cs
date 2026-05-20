using CommunityToolkit.Mvvm.ComponentModel;
using System;
using static ORT一键报告.ReportUtils;

namespace ORT一键报告.ViewModels
{
    public class ReportHeaderViewModel : ObservableObject
    {
        private DataCell _testedBy;
        public DataCell TESTED_BY { get => _testedBy; set => SetPropertyAndNotifyDependents(ref _testedBy, value, nameof(IsAllFilled)); }

        private DataCell _approvedBy;
        public DataCell APPROVED_BY { get => _approvedBy; set => SetPropertyAndNotifyDependents(ref _approvedBy, value, nameof(IsAllFilled)); }

        private DataCell _projectName;
        public DataCell PROJECT_NAME { get => _projectName; set => SetPropertyAndNotifyDependents(ref _projectName, value, nameof(IsAllFilled)); }

        private DataCell _testStage;
        public DataCell TEST_STAGE { get => _testStage; set => SetPropertyAndNotifyDependents(ref _testStage, value, nameof(IsAllFilled)); }

        private DataCell _testDescription;
        public DataCell TestDescription { get => _testDescription; set => SetPropertyAndNotifyDependents(ref _testDescription, value, nameof(IsAllFilled)); }
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
                string[] res = new string[8];

                res[0] = TESTED_BY.Data;
                res[1] = APPROVED_BY.Data;
                res[2] = PROJECT_NAME.Data;
                res[3] = TEST_STAGE.Data;
                res[4] = TestStart.ToString("d");
                res[5] = TestEnd.ToString("d");
                res[6] = TestPass ? "Pass" : "Fail";
                res[7] = TestDescription.Data;

                return res;
            }
        }

        public bool IsAllFilled => DataCell.IsNotNull(TESTED_BY) &&
                    DataCell.IsNotNull(APPROVED_BY) &&
                    DataCell.IsNotNull(PROJECT_NAME) &&
                    DataCell.IsNotNull(TEST_STAGE) &&
                    DataCell.IsNotNull(TestDescription);

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