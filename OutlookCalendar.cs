namespace OutlookCalendarAPI;

public class OutlookCalendar
{
    public string? Message { get; set; }

    public int Status { get; set; }
}

public class OutlookRecipient
{
    public string? Recipient { get; set; }
    
    public Boolean Required  { get; set; } = true;
    
}
public class OutlookAppointment
{
    public DateTime Start { get; set; }

    public DateTime End  { get; set; }
    
    public string? Location  { get; set; }

    public string? Body  { get; set; }

    public Boolean AllDayEvent  { get; set ; }=false;
                
    public string? Subject { get; set; }

    ///For Editing Only
    public string? NewSubject { get; set; }

    public List<OutlookRecipient>? OlRecipients { get; set; }
    
}