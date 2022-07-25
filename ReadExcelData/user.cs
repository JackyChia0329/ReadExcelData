using System;
using System.Collections.Generic;
using System.Text;

namespace ReadExcelData
{
    class user
    {
        public string username { get; set; }
        public List<userAcc> userAccList { get; set; }
        public class userAcc
        {
            public string ID { get; set; }
            public string Username { get; set; }
          
        }

    }
}
