using Interop.IdeaLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace UtiltyCasewareIdea
{
  public static class UtilityCasewareIdea
   {
      public static string GetProjectDirectory()
      {

         IdeaClient client = null;
         string dir;
         try
         {
            client = new IdeaClient();
            try
            {
               string temp = client.WorkingDirectory;
               temp = temp.Remove(temp.Length - 1);
               string temp2 = Path.GetFileName(temp);

                dir = client.ManagedProject;

               if (dir==temp2)
               {
                  return client.WorkingDirectory; ;
               }
               else
               {
                  return Directory.GetParent(temp).FullName+"\\";
               }
              
              
            }
            catch (Exception)
            {
               try
               {
                  dir = client.ExternalProject;
                  return dir;
               }
               catch (Exception ex)
               {
                  MessageBox.Show("could not get current projet" + Environment.NewLine + ex.Message);
                  Environment.Exit(1);
                  throw ex;
               }
            }
         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.Message);
            Environment.Exit(1);
            throw ex;
         }
         finally
         {
            DisposeCom(client);
         }
      }

      public static void DisposeCom(object obj)
      {
         if (obj == null)
         {
            return;
         }
         try
         {
            Marshal.ReleaseComObject(obj);
         }
         catch (Exception)
         {
            obj = null;
         }
      }

      public static string externalProject()
      {
         IdeaClient client = null;
         ProjectManagement pm = null;
         try
         {
            client = new IdeaClient();
            pm = client.ProjectManagement();
            return client.ExternalProject;
                   
         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.Message);
            return "Exception occured";
         }
         finally
         {
            UtilityCasewareIdea.DisposeCom(client);
            UtilityCasewareIdea.DisposeCom(pm);

         }
      }
      public static string GetCurrentDB()
      {
         IdeaClient client = null;

         try
         {
            client = new IdeaClient();
            string temp = client.CurrentDatabase().Name;
            return temp;

         }
         catch (Exception ex)
         {
           // MessageBox.Show(ex.Message);
            return "";
         }
         finally
         {
            UtilityCasewareIdea.DisposeCom(client);          

         }

      }
         
      public static void StartIdea()
      {
         IdeaClient client = null;
         try
         {

            Process[] process = Process.GetProcessesByName("idea");
            if (process.Length == 0)
            {
               Process p = Process.Start(@"C:\Program Files (x86)\CaseWare IDEA\IDEA\idea.exe");
               p.WaitForInputIdle(-1);

            }
            else
            {
               client = new IdeaClient();
               client.ShowWindow();
               // MessageBox.Show("running");
            }


            //client.ShowWindow();
            //if (first)
            //{
            //   first = false;
            //   Execute();
            //   return;
            //}

            //IntPtr handle ;
            //Process[] process = Process.GetProcessesByName("idea");
            //handle= process[0].MainWindowHandle;
            //foreach (Process p in Process.GetProcesses())
            //{
            //   if (p.MainWindowTitle.ToLower().Contains("idea"))
            //   {
            //      handle = p.MainWindowHandle;
            //      break;
            //   }
            //}
            //client.ShowWindow();
            //SetForegroundWindow(handle);
         }
         catch (Exception ex)
         {

            MessageBox.Show(ex.Message);
         }
         finally
         {
            DisposeCom(client);
         }


      }

      public static void ShowWindow()
      {
         IdeaClient client = null;
         IntPtr handle = IntPtr.Zero;
         try
         {
            Process[] process = Process.GetProcessesByName("idea");
            if (process.Length == 0)
            {
               Process p = Process.Start(@"C:\Program Files (x86)\CaseWare IDEA\IDEA\idea.exe");
               p.WaitForInputIdle(-1);
               handle = p.MainWindowHandle;
            }
            else
            {
               handle = process[0].MainWindowHandle;
            }
            client = new IdeaClient();
            client.ShowWindow();
            SetForegroundWindow(handle);
            client.ShowWindow();
            SetForegroundWindow(handle);

            #region code history
            //foreach (Process p in Process.GetProcesses())
            //{
            //   if (p.MainWindowTitle.ToLower().Contains("caseware"))
            //   {
            //      handle = p.MainWindowHandle;
            //      break;
            //   }
            //}

            //if (first ==true)
            //{
            //   first = false;
            //   Thread.Sleep(500);
            //   DisposeCom(client);
            //   ShowWindow();

            //}
            //else
            //{
            //   first = true;
            //}
            #endregion
         }
         catch (Exception ex)
         {

            MessageBox.Show(ex.Message);
         }
         finally
         {
            DisposeCom(client);
         }

      }

      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      static extern bool SetForegroundWindow(IntPtr hWnd);
   }
}
