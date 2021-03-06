﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Actigraph.Forms.Reports
{
    static class Program
    {

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
            (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            log.Info("Reports Started!!!");
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

           

            Application.Run(new LoadData());
            log.Info("Reports Ended!!!");
        }
      
    }
}
