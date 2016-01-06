using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;
using System.Net;
using System.Web.Http;
using System.Linq;
using System;

namespace Slalom.Huddle.OutlookApi.Controllers
{
    public class RoomsController : ApiController
    {
        // SET THESE AS A TEST UNTIL ENCODING IS IN PLACE
        private const string email = "notexist@slalom.com";
        private const string password = "notapassword";
        private const int defaultMeetingDuration = 30;

        // GET api/values
        public IEnumerable<string> Get()
        {
            // Set up the exchange service for UTC
            ExchangeService service = new ExchangeService(TimeZoneInfo.Utc);

            // Create network credentials and specify the public office 365 URL
            service.Credentials = new NetworkCredential(email, password);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

            // For test output, here's a list of strings!
            List<string> strings = new List<string>();

            // For each room in the Seattle conference address...
            foreach (var room in service.GetRooms("Slalom--SeattleConferenceRooms@slalom.com"))
            {
                // .. Create an attendee and a time window, which has to be at least 1 day.
                AttendeeInfo attendee = new AttendeeInfo(room.Address);
                TimeWindow timeWindow = new TimeWindow(DateTime.Now.ToUniversalTime().Date, DateTime.Now.ToUniversalTime().Date.AddDays(1));

                // We want to know if the room is free or busy.
                AvailabilityData availabilityData = AvailabilityData.FreeBusy;
                AvailabilityOptions availabilityOptions = new AvailabilityOptions();
                attendee.ExcludeConflicts = false;
                attendee.AttendeeType = MeetingAttendeeType.Organizer;

                // Get the availability of the room.
                var result = service.GetUserAvailability(new[] { attendee }, timeWindow, availabilityData, availabilityOptions);

                // See if there are any rooms that have this block of time cleared.
                DateTime utcTime = DateTime.Now.ToUniversalTime();
                bool available = !(from n in result.AttendeesAvailability[0].CalendarEvents
                                 where (n.StartTime < utcTime && n.EndTime > utcTime) ||
                                 (n.StartTime > utcTime && n.StartTime < utcTime.AddMinutes(defaultMeetingDuration))
                                   select n).Any();

                strings.Add(string.Format("{0} - {1}: {2}", room.Name, room.Address, available ? "Available" : "NOPE"));
            }

            return strings.ToArray();
        }        
    }
}
