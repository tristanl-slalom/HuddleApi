using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Web.Http;
using System.Linq;
using System;

using Slalom.Huddle.OutlookApi.Models;
using Slalom.Huddle.OutlookApi.Services;

namespace Slalom.Huddle.OutlookApi.Controllers
{
    public class RoomsController : ApiController
    {
        public const int DefaultMinimumPeople = 2;
        public const int DefaultMeetingDuration = 30;
        public const int DefaultPreferredFloor = 15;
        private ExchangeService service;
        private string emailAccount;

        public ServiceGenerator ServiceGenerator { get; private set; }

        public RoomLoader RoomLoader { get; private set; }

        public RoomsController()
        {
            // Values from the config file
            this.emailAccount = ConfigurationManager.AppSettings["emailAccount"];
            string password = ConfigurationManager.AppSettings["accountPassword"];

            // Create a service generator and give us an authenticated service.
            ServiceGenerator = new ServiceGenerator();
            this.service = ServiceGenerator.GenerateService(emailAccount, password);

            // Initailize a room loader with the service and service account.
            RoomLoader = new RoomLoader(service, emailAccount);
        }

        // GET api/values
        public IHttpActionResult Get(
            int minimumPeople = DefaultMinimumPeople,
            int duration = DefaultMeetingDuration, 
            int preferredFloor = DefaultPreferredFloor,
            string command = "")
        {
            try
            {
                // Use the service to load all of the rooms
                List<Room> rooms = RoomLoader.LoadRooms(preferredFloor);

                // Use the service to load the schedule of all the rooms
                RoomLoader.LoadRoomSchedule(rooms, duration);

                // Find the first available room that supports the number of people.
                Room selectedRoom = rooms.FirstOrDefault(n => n.Available && n.RoomInfo.MaxPeople >= minimumPeople);
                if (selectedRoom == null)
                {
                    return Ok(new { Text = "I'm sorry, there are no rooms available for you right now. Try again another time!" });
                }

                // Acquire the meeting room for the duration.
                Appointment meeting = RoomLoader.AcquireMeetingRoom(selectedRoom, duration, preferredFloor, command);

                // Verify that the meeting was created by matching the subject.
                //Item item = Item.Bind(service, meeting.Id, new PropertySet(ItemSchema.Subject));
                //if (item.Subject != meeting.Subject)
                //{
                //    return StatusCode(HttpStatusCode.ServiceUnavailable);
                //}

                // Return a 200
                return Ok(meeting.Body);
            }
            catch (Exception exception)
            {
                return Ok(new { Text = "I'm sorry, there appears to be a problem with the service. Make sure that authorization credentials have been loaded into the service configuration." + exception.Message });
            }
        }
    }
}
