﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C_Wimform
{
    static class Program
    {
        private static System.Threading.Mutex mutex;
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            mutex = new System.Threading.Mutex(true, "OnlyRun");
            if (mutex.WaitOne(0,false))
            {
                Application.Run(new Form1());
            }
            else
            {
                MessageBox.Show("程式已在執行","錯誤",MessageBoxButtons.OK,MessageBoxIcon.Error);
                Application.Exit();
            }
            
        }
    }
}
