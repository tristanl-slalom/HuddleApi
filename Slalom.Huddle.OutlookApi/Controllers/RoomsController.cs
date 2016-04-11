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
                DateTime currentDateTime = DateTime.Now.ToLocalTime();
                DateTime startDate = currentDateTime;

                if (!String.IsNullOrEmpty(startTime))
                {
                    // startTime will be in the form of = 04:00 or 14:00 (military time)
                    var startTimeChunk = startTime.Split(":".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                    // set the hour and minute part while keeping the date to current date
                    TimeSpan hourMinutes = new TimeSpan(int.Parse(startTimeChunk[0]), int.Parse(startTimeChunk[1]), 0);
                    startDate = startDate.Date + hourMinutes;

                    // verify that startDate is NOT in the past. Huddle API is only allowing to request room for current day, 
                    // users are allowed to request rooms sometime in the future as long as it is still today.
                    if (startDate < currentDateTime)
                    {
                        return Ok(new { Text = RoomLoader.Wrap("I'm sorry, there are no rooms available for you right now. Try again another time!") });
                    }                
                }

                DateTime endDate = DurationAdjuster.ExtendDurationToNearestBlock(startDate, duration);

                // Use the service to load all of the rooms
                List<Room> rooms = RoomLoader.LoadRooms(preferredFloor);

                // Use the service to load the schedule of all the rooms
                RoomLoader.LoadRoomSchedule(rooms, startDate, endDate);

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
