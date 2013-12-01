using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Tasks
{

    public class MyTaskItem {
        public string Id { get; set;}
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime DueDate { get; set; }
        public bool IsFinished { get; set; }
    }

    
    public class OutlookSync
    {
        private Outlook._Application outlookObj = new Outlook.Application();

        public IList<MyTaskItem> AllTasks
        {
            get
            {
                Outlook.MAPIFolder outlookTasks = (Outlook.MAPIFolder)outlookObj.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
                int numberOfTasks = outlookTasks.Items.Count;
                IList<MyTaskItem> allTasks = new List<MyTaskItem>(numberOfTasks);
                foreach (Microsoft.Office.Interop.Outlook.TaskItem taskItem in outlookTasks.Items)
                {
                    allTasks.Add(new MyTaskItem() { Id = taskItem.EntryID, Subject = taskItem.Subject, Body = taskItem.Body, DueDate = taskItem.DueDate, IsFinished = taskItem.Status == Outlook.OlTaskStatus.olTaskComplete});
                }

                return allTasks;
            }
        }


        //public void AddTask(Outlook._TaskItem task)
        //{
        //    Outlook.MAPIFolder outlookTasks = (Outlook.MAPIFolder)outlookObj.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks);
        //    outlookTasks.Items.Add(task);
        //}

        //public void DeleteTask(Outlook._TaskItem task)
        //{
        //    task.Delete();
        //}
        
        static int Main(string[] args)
        {
            //Operations.RatePerHour = 5;
            //double setting = Operations.RatePerHour;
            //OutlookSync n1 = new OutlookSync();
            //IList<Outlook._TaskItem> allTasks = n1.AllTasks;
            //Console.WriteLine(allTasks[0].Status);
            //Operations.CalculateOverdue(allTasks[0]);
            //Console.Write(allTasks.Count);
            //Console.Read();

            ////OutlookSync myOutlook = new OutlookSync();
            ////myOutlook.c_tasks();
            ////myOutlook.iGetAllTaskItems();

            
            return 0;
        }
    }
}
