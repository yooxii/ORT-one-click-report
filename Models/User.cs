using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ORT一键报告.Models
{
    public class User
    {
        public string Username { get; set; }
        public string Password { get; set; }

        public bool Validate()
        {
            // 实际应用中，这里应该有实际的验证逻辑
            return !string.IsNullOrEmpty(Username) && Password.Length >= 6;
        }
    }
}
