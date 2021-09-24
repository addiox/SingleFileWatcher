using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FileWatcherV1
{
    [RunInstaller(true)]
    public partial class Service1 : ServiceBase
    {
        readonly string WatchPath1 = ConfigurationManager.AppSettings["WatchPath1"];
        private readonly string ConnectionStrings = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
        public Service1()
        {
            InitializeComponent();
            fileSystemWatcher1.Created += FileWatcherWatchDDriveFolder_Created;

        }

        protected override void OnStart(string[] args)
        {
            try
            {
                fileSystemWatcher1.Path = WatchPath1;
                fileSystemWatcher1.IncludeSubdirectories = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected override void OnStop()
        {
            try
            {
                Create_ServiceStoptextfile();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        /// <summary>  
        /// This event monitor folder wheater file created or not.  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        void FileWatcherWatchDDriveFolder_Created(object sender, System.IO.FileSystemEventArgs e)
        {
            try
            {
                Thread.Sleep(15000);
                //Then we need to check file is exist or not which is created.  
                if (CheckFileExistance(WatchPath1, e.Name))
                {
                    //Then write code for log detail of file in text file.  
                    CreateTextFile(WatchPath1, e.Name);
                    //StoredProcedureExecute(e.Name);
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private bool CheckFileExistance(string FullPath, string FileName)
        {
            // Get the subdirectories for the specified directory.'  
            bool IsFileExist = false;
            DirectoryInfo dir = new DirectoryInfo(FullPath);
            if (!dir.Exists)
                IsFileExist = false;
            else
            {
                string FileFullPath = Path.Combine(FullPath, FileName);
                if (File.Exists(FileFullPath))
                    IsFileExist = true;
            }
            return IsFileExist;


        }

        public static void Create_ServiceStoptextfile()
        {
            string Destination = "C:\\Files\\FileWatcherWinService";
            StreamWriter SW;
            if (Directory.Exists(Destination))
            {
                Destination = System.IO.Path.Combine(Destination, "txtServiceStop_" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
                if (!File.Exists(Destination))
                {
                    SW = File.CreateText(Destination);
                    SW.Close();
                }
            }
            using (SW = File.AppendText(Destination))
            {
                SW.Write("\r\n\n");
                SW.WriteLine("Service Stopped at: " + DateTime.Now.ToString("dd-MM-yyyy H:mm:ss"));
                SW.Close();
            }
        }

        private void CreateTextFile(string FullPath, string FileName)
        {
            StreamWriter SW;
            if (!File.Exists(Path.Combine("C:\\Debug", "txtStatus_" + DateTime.Now.ToString("yyyyMMdd") + ".txt")))
            {
                SW = File.CreateText(Path.Combine("C:\\Debug", "txtStatus_" + DateTime.Now.ToString("yyyyMMdd") + ".txt"));
                SW.Close();
            }
            using (SW = File.AppendText(Path.Combine("C:\\Debug", "txtStatus_" + DateTime.Now.ToString("yyyyMMdd") + ".txt")))
            {
                SW.WriteLine("File Created with Name: " + FileName + " at this location: " + FullPath);
                SW.Close();
            }

            MoveAnotherLocation(FullPath, FileName);
        }

        private void MoveAnotherLocation(string FullPath, string FileName)
        {
            try
            {
                string source = Path.Combine(FullPath, FileName);
                string destination = Path.Combine(@"C:\ShiftData\InBound", FileName);

                File.Move(source, destination);

            }
            catch (Exception)
            {
            }
        }

        private void StoredProcedureExecute(string FileName)
        {
            string connectionString = ConnectionStrings;
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand
                {
                    CommandText = "ShiftBulkInsert",
                    CommandType = CommandType.StoredProcedure,
                    Connection = con
                };
                cmd.Parameters.AddWithValue("@FileName", FileName);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        CreateTextFile(WatchPath1, "StoredProcedure Executed");
                    }
                }
            }
        }
    }
}
