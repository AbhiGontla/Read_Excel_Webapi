﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Read_Excel_Webapi.Models
{
    public class EmployeeModel
    {
        public int EmployeeId { get; set; }


        public string Name { get; set; }

        public string Address { get; set; }
        public string EmailId { get; set; }
    }
}