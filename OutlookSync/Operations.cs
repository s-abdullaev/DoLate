using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace Tasks
{
    class Operations
    {
        public static TimeSpan CalculateOverdue(_TaskItem task)
        {
            TimeSpan overdue = DateTime.Now.Subtract(task.DueDate);
            return overdue;
        }

        public static double CalculateTotalFines(double rate, TimeSpan overdue)
        {
            double totalAmount = rate * overdue.Hours;
            return totalAmount;
        }

        public static double RatePerHour
        {
            get
            {
                if (ConfigurationManager.AppSettings["RatePerHour"] != null)
                {
                    string strPerHour = ConfigurationManager.AppSettings["RatePerHour"];
                    return Double.Parse(strPerHour);
                }
                else
                {
                    return -1;
                }
            }

            set
            {
                Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings.Add("RatePerHour", value.ToString());
                config.Save(ConfigurationSaveMode.Modified);
            }
        }
    }
}
