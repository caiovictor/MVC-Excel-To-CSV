﻿using System.Web;
using System.Web.Mvc;

namespace Excel.Para.CSV
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
