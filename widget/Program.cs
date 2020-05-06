using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Reflection;
using System.Windows.Forms;

namespace widget
{
    static class Program
    {
        /// <summary>
        /// Hlavní vstupní bod aplikace.
        /// </summary>
        /// 
        static SplashScreen mySplashForm;
        static Form1 myMainForm;
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());






            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Show Splash Form
            mySplashForm = new SplashScreen();
            if (mySplashForm != null)
            {
                Thread splashThread = new Thread(new ThreadStart(
                    () => { Application.Run(mySplashForm); }));
                splashThread.SetApartmentState(ApartmentState.STA);
                splashThread.Start();
            }
            //Create and Show Main Form
            myMainForm = new Form1();
            myMainForm.LoadCompleted += MainForm_LoadCompleted;
            Application.Run(myMainForm);
            if (!(mySplashForm == null || mySplashForm.Disposing || mySplashForm.IsDisposed))
                mySplashForm.Invoke(new Action(() => {
                    mySplashForm.TopMost = true;
                    mySplashForm.Activate();
                    mySplashForm.TopMost = false;
                }));
        }





        private static void MainForm_LoadCompleted(object sender, EventArgs e)
        {
            if (mySplashForm == null || mySplashForm.Disposing || mySplashForm.IsDisposed)
                return;
            mySplashForm.Invoke(new Action(() => { mySplashForm.Close(); }));
            mySplashForm.Dispose();
            mySplashForm = null;
            myMainForm.TopMost = true;
            myMainForm.Activate();
            myMainForm.TopMost = false;
        }







    }
}
