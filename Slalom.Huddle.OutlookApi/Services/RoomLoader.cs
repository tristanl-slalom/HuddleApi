using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Exchange.WebServices.Data;
using Slalom.Huddle.OutlookApi.Models;
using Slalom.Huddle.OutlookApi.Controllers;

namespace Slalom.Huddle.OutlookApi.Services
{
    public class RoomLoader
    {
        private ExchangeService service;
        private string serviceAccount;

        public RoomLoader(ExchangeService service, string serviceAccount)
        {
            this.service = service;
            this.serviceAccount = serviceAccount;
        }

        public List<Room> LoadRooms(int preferredFloor)
        {
            List<Room> rooms = new List<Room>();

            // For each room in the Seattle conference address...
            foreach (var room in service.GetRooms("Slalom--SeattleConferenceRooms@slalom.com"))
            {
                // Skip any room that doesn't have the parsable data.
                if (!room.Name.Contains('['))
                    continue;

                RoomInfo roomInfo = RoomInfo.ParseRoomInfo(room.Name);

                // .. Create an attendee and a time window, which has to be at least 1 day.
                AttendeeInfo attendee = GetAttendeeFromRoom(room);
                rooms.Add(new Room { AttendeeInfo = attendee, RoomInfo = roomInfo });
            }

            // Sort by max people first, then floor.
            rooms = (from n in rooms
                     orderby Math.Abs(n.RoomInfo.Floor - preferredFloor) ascending,
                             n.RoomInfo.MaxPeople ascending
                     select n).ToList();
            return rooms;
        }

        public GetUserAvailabilityResults LoadRoomSchedule(List<Room> rooms, DateTime startDate, DateTime endDate)
        {
            //TimeWindow timeWindow = new TimeWindow(DateTime.Now.ToUniversalTime().Date, DateTime.Now.ToUniversalTime().Date.AddDays(1));
            TimeWindow timeWindow = new TimeWindow(startDate.ToUniversalTime().Date, startDate.ToUniversalTime().Date.AddDays(1));

            // We want to know if the room is free or busy.
            AvailabilityData availabilityData = AvailabilityData.FreeBusy;
            AvailabilityOptions availabilityOptions = new AvailabilityOptions();
            availabilityOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;
            availabilityOptions.MaximumSuggestionsPerDay = 0;

            // Get the availability of the room.
            var result = service.GetUserAvailability(from n in rooms
                                                     select n.AttendeeInfo,
                                                     timeWindow,
                                                     availabilityData,
                                                     availabilityOptions);

            // Use the schedule to determine if a room is available for the next
            DetermineRoomAvailability(rooms, result, startDate, endDate);
            return result;
        }

        private static AttendeeInfo GetAttendeeFromRoom(EmailAddress room)
        {
            var attendee = new AttendeeInfo(room.Address);
            attendee.ExcludeConflicts = false;
            attendee.AttendeeType = MeetingAttendeeType.Organizer;

            return attendee;
        }

        private static void DetermineRoomAvailability(List<Room> rooms,
            GetUserAvailabilityResults result, DateTime startDate,
            DateTime endDate
            )
        {
            //DateTime utcTime = DateTime.Now.ToUniversalTime();
            DateTime utcTime = startDate.ToUniversalTime();

            if (rooms.Count != result.AttendeesAvailability.Count)
            {
                throw new Exception($"The number of known rooms ({rooms.Count}) did not match the number of availabilities ({result.AttendeesAvailability.Count}).");
            }

            for (int i = 0; i < rooms.Count; i++)
            {
                var availability = result.AttendeesAvailability[i];
                var room = rooms[i];

                room.Available = !(from n in result.AttendeesAvailability[i].CalendarEvents
                                   where (n.StartTime < utcTime && n.EndTime > utcTime) ||
                                   (n.StartTime > utcTime && n.StartTime < endDate.ToUniversalTime())
                                   select n).Any();
            }
        }

        public Appointment AcquireMeetingRoom(Room selectedRoom, DateTime startDate, DateTime endDate, int requestedFloor, string command)
        {
            Appointment meeting = new Appointment(service);
            meeting.Subject = "Group Huddle";
            if (selectedRoom.RoomInfo.Floor == requestedFloor || requestedFloor == RoomsController.DefaultPreferredFloor)
            {
                CommandResponder responder = new CommandResponder();
                string message = responder.CreateResponseForCommand(command, selectedRoom, endDate, requestedFloor);
                meeting.Body = message;
            }
            else
            {
                string timeAsString = endDate.ToShortTimeString();
                meeting.Body = $"I'm sorry, I couldn't find a room on floor {requestedFloor}, but I have scheduled '{selectedRoom.RoomInfo.Name}' for you on floor {selectedRoom.RoomInfo.Floor} until {timeAsString}";
            }
            
            meeting.Start = startDate;
            meeting.End = endDate;
            meeting.Location = $"{selectedRoom.RoomInfo.Name} on Floor {selectedRoom.RoomInfo.Floor}";
            meeting.RequiredAttendees.Add(selectedRoom.AttendeeInfo.SmtpAddress);
            meeting.RequiredAttendees.Add(serviceAccount);

            // Save the meeting to the Calendar folder and send the meeting request.
            meeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);
            return meeting;
        }

        public static string Wrap(string message, string audioUrl = "")
        {
            string formattedUrl = audioUrl == "" ? "" : $"<audio src='{audioUrl}'/>";
            return $"<speak>{formattedUrl}{message}</speak>";
        }
    }
}