using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookVSTOAddIn
{
    public partial class CalendarAddIn
    {
        private void CalendarAddIn_Startup(object sender, System.EventArgs e)
        {
            var caseId = DateTime.Now.ToString("hhmmssyyyyMMdd");
            var subject = $"[Case {caseId}] Trial";
            var body = "This appointment is programmatically created.";
            var startTime = DateTime.Today;
            var endTime = startTime;
            AddAppointment(subject, body, startTime, endTime);
        }

        private void CalendarAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void AddAppointment(string subject, string body, DateTime startTime, DateTime endTime)
        {
            // Source: https://docs.microsoft.com/en-us/visualstudio/vsto/how-to-programmatically-create-appointments
            try
            {
                Outlook.AppointmentItem newAppointment =
                    (Outlook.AppointmentItem)
                Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                newAppointment.Start = startTime;
                newAppointment.End = endTime;
                newAppointment.Subject = subject;
                newAppointment.Body = body;
                newAppointment.AllDayEvent = true;
                //newAppointment.Recipients.Add("Roger Harui");
                //Outlook.Recipients sentTo = newAppointment.Recipients;
                //Outlook.Recipient sentInvite = null;
                //sentInvite = sentTo.Add("Holly Holt");
                //sentInvite.Type = (int)Outlook.OlMeetingRecipientType
                //    .olRequired;
                //sentInvite = sentTo.Add("David Junca ");
                //sentInvite.Type = (int)Outlook.OlMeetingRecipientType
                //    .olOptional;
                //sentTo.ResolveAll();
                newAppointment.Save();
                //newAppointment.Display(false);
                LogInfo($"Check calendar for \"{subject}\"");
            }
            catch (Exception ex)
            {
                LogInfo("The following error occurred: " + ex.Message);
            }
        }

        private void LogInfo(string msg)
        {
            MessageBox.Show(msg);
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(CalendarAddIn_Startup);
            this.Shutdown += new System.EventHandler(CalendarAddIn_Shutdown);
        }

        #endregion
    }
}
