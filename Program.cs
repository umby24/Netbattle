using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Netbattle.Common;
using Netbattle.Forms;
using TaskScheduler = Netbattle.Common.TaskScheduler;

namespace Netbattle {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main() {
            new Logger().Setup(); 
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(true);
            Application.Run(new Container());
        }

        static void TaskMain() {
            while (true) {
                TaskScheduler.RunMainTasks();
                Thread.Sleep(1);
            }
        }
    }
}
