using Microsoft.AspNetCore.Mvc;
using Microsoft.Office.Interop.Outlook;

namespace OutlookCalendarAPI.Controllers;

[ApiController]
[Route("[controller]")]
public class OutlookCalendarController : ControllerBase
{
    private readonly ILogger<OutlookCalendarController> _logger;

    public OutlookCalendarController(ILogger<OutlookCalendarController> logger)
    {
        _logger = logger;
    }

    [HttpPost(Name = "PostOutlookAppointment")]
    public IEnumerable<OutlookCalendar> Post(OutlookAppointment aptm)
    {
        //-->
        var oc = new OutlookCalendar();
        try
            {
                var app = new Application();
                AppointmentItem newAppointment = (AppointmentItem)app.CreateItem(OlItemType.olAppointmentItem);
                newAppointment.Start = aptm.Start;//DateTime.Now.AddHours(2);
                newAppointment.End = aptm.End;//DateTime.Now.AddHours(3);
                newAppointment.Location = aptm.Location;//"ConferenceRoom #2345";
                newAppointment.Body = aptm.Body;//"We will discuss progress on the group project.";
                newAppointment.AllDayEvent = aptm.AllDayEvent;//false;
                newAppointment.Subject = aptm.Subject;//"Group Project";

                Recipients sentTo = newAppointment.Recipients;
                Recipient sentInvite = null;
                
                int i=0;
                foreach(var item in aptm.OlRecipients)
                {    
                    if(i==0)
                    {
                        if(item.Recipient != null)
                        {
                            newAppointment.Recipients.Add(item.Recipient);
                            i++;
                        }
                    }
                    else
                    {
                        if(item.Recipient != null)
                        {
                            sentInvite = sentTo.Add(item.Recipient);
                            sentInvite.Type = item.Required?(int)OlMeetingRecipientType.olRequired:(int)OlMeetingRecipientType.olOptional;
                        }
                    }                 
                }
                sentTo.ResolveAll();
                
                newAppointment.Save();
                app = null;
                oc.Message= "Appointment added successfully";
                oc.Status = 200;

            }
            catch (System.Exception ex)
            {
                oc.Message= "The following error occurred: " + ex.Message;
                oc.Status = ex.HResult;

            }
        //<--
        

        
        return new List<OutlookCalendar>{ oc };
        
    }
    
    [HttpPut(Name = "PutOutlookAppointment")]
    public IEnumerable<OutlookCalendar> Put(OutlookAppointment aptm)
    {
        var app = new Application();
        var oc = new OutlookCalendar();
        try
        {
            MAPIFolder calendar = app.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            Items calendarItems = calendar.Items;

            AppointmentItem item = (AppointmentItem)calendarItems[aptm.Subject];

            if (item != null)
            {
                item.Start = aptm.Start==null?item.Start:aptm.Start;
                item.End = aptm.End==null?item.End:aptm.End;
                item.Location = aptm.Location==""?item.Location:aptm.Location;
                item.Body = aptm.Body==""?item.Body:aptm.Body;
                item.AllDayEvent = aptm.AllDayEvent;
                item.Subject = aptm.NewSubject==""?item.Subject:aptm.NewSubject;
                
                int i=item.Recipients.Count;

                while(i>0)
                {
                    item.Recipients.Remove(i);
                    i--;
                }

                Recipients sentTo = item.Recipients;
                Recipient sentInvite = null;
                
                
                foreach(var rec in aptm.OlRecipients)
                {    
                        if(rec.Recipient != null)
                        {
                            sentInvite = sentTo.Add(rec.Recipient);
                            sentInvite.Type = rec.Required?(int)OlMeetingRecipientType.olRequired:(int)OlMeetingRecipientType.olOptional;
                        }
                }
                sentTo.ResolveAll();
                
                item.Save();
                app = null;

                oc.Message= "The event was successfully updated ";
                oc.Status = 200;
            }
            else
            {
                oc.Message= "The event you are trying to update does not exist";
                oc.Status = 500;
            }
        }
        catch(System.Exception ex)
        {
            oc.Message= "The following error occurred: " + ex.Message;
            oc.Status = ex.HResult;
        }
        return new List<OutlookCalendar>{ oc };
    }

    [HttpDelete(Name = "DeleteOutlookAppointment")]
    public IEnumerable<OutlookCalendar> Delete(string subject)
    {
        var app = new Application();
        var oc = new OutlookCalendar();
        try
        {
            MAPIFolder calendar = app.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            Items calendarItems = calendar.Items;

            AppointmentItem item = (AppointmentItem)calendarItems[subject];

            if (item != null)
            {
                item.Delete();
                oc.Message= "The event was deleted successfully";
                oc.Status = 200;
            }
            else
            {
                oc.Message= "The event you are trying to delete does not exist";
                oc.Status = 500;
            }
        }
        catch(System.Exception ex)
        {
            oc.Message= "The following error occurred: " + ex.Message;
            oc.Status = ex.HResult;
        }
        return new List<OutlookCalendar>{ oc };
    }



    [HttpGet(Name = "GetOutlookAppointment")]
    public IEnumerable<OutlookCalendar> Get(DateTime StartDate, DateTime EndDate)
    {
        try{
            var app = new Application();
            var loc = new List<OutlookCalendar>();

            MAPIFolder calFolder = app.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            DateTime start = StartDate;
            DateTime end = EndDate;
            Items rangeAppts = GetAppointmentsInRange(calFolder, start, end);
            
            if (rangeAppts != null)
            {
                foreach (AppointmentItem appt in rangeAppts)
                {
                    loc.Add(new OutlookCalendar{ 
                        Message="Subject: '" + appt.Subject + "' Start: " + appt.Start.ToString("g"),
                        Status=200
                    });
                }
            }
            
            return loc;
        }
        catch(System.Exception ex)
        {
            return new List<OutlookCalendar>{ new OutlookCalendar{ Message=ex.Message, Status=ex.HResult} };
        }
    }

        /// <summary>
        /// Get recurring appointments in date range.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="startTime"></param>
        /// <param name="endTime"></param>
        /// <returns>Outlook.Items</returns>
        private static Items GetAppointmentsInRange(MAPIFolder folder, DateTime startTime, DateTime endTime)
        {
            string filter = "[Start] >= '" + startTime.ToString("g") + "' AND [END] <= '"+ endTime.ToString("g") + "'";
            //Console.WriteLine(filter);
            try
            {
                Items calItems = folder.Items;
                
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]");
                Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }

}

