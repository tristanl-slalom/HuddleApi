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
        public const int DefaultMinimumPeople = 3;
        public const int DefaultMeetingDuration = 30;
        public const int DefaultPreferredFloor = 15;

        public const string SuccessAudioUrl = "https://s3.amazonaws.com/uploads.hipchat.com/41664/1050651/vXTXicIkfZ8M0BI/success_short.mp3";
        private ExchangeService service;
        private string emailAccount;

        public ServiceGenerator ServiceGenerator { get; private set; }

        public RoomLoader RoomLoader { get; private set; }

        public DurationAdjuster DurationAdjuster { get; private set; }

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
            DurationAdjuster = new DurationAdjuster();
        }

        // GET api/values
        public IHttpActionResult Get(
            int minimumPeople = DefaultMinimumPeople,
            int duration = DefaultMeetingDuration, 
            int preferredFloor = DefaultPreferredFloor,
            String startTime = null,
            string command = "")
        {
            try
            {
                // AMAZON.TIME – converts words that indicate time (“four in the morning”, “two p m”) into a time value (“04:00”, “14:00”).
                DateTime startDate = DateTime.Now.ToLocalTime();
                if (!String.IsNullOrEmpty(startTime))
                {
                    startDate = DateTime.ParseExact(startTime, "H:mm", null, System.Globalization.DateTimeStyles.None);
                } 
                
                DateTime endDate = DurationAdjuster.ExtendDurationToNearestBlock(startDate, duration);
                throw new Exception("Atest");
                // Use the service to load all of the rooms
                List<Room> rooms = RoomLoader.LoadRooms(preferredFloor);

                // Use the service to load the schedule of all the rooms
                RoomLoader.LoadRoomSchedule(rooms, endDate);

                // Find the first available room that supports the number of people.
                Room selectedRoom = rooms.FirstOrDefault(n => n.Available && n.RoomInfo.MaxPeople >= minimumPeople);
                if (selectedRoom == null)
                {
                    return Ok(new { Text = RoomLoader.Wrap("I'm sorry, there are no rooms available for you right now. Try again another time!") });
                }

                // Acquire the meeting room for the duration.
                Appointment meeting = RoomLoader.AcquireMeetingRoom(selectedRoom, startDate, endDate, preferredFloor, command);

                // Verify that the meeting was created by matching the subject.

                // Return a 200
                return Ok(new { Text = RoomLoader.Wrap(meeting.Body, SuccessAudioUrl) });
            }
            catch (Exception exception)
            {
                return Ok(new { Text = RoomLoader.Wrap("I'm sorry, there appears to be a problem with the service. Make sure that authorization credentials have been loaded into the service configuration.") });
            }
        }
    }
}
