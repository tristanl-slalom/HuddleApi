using System;
using System.Linq;
using Slalom.Huddle.OutlookApi.Models;

namespace Slalom.Huddle.OutlookApi.Services
{
    internal class CommandResponder
    {
        private string command;

        public string CreateResponseForCommand(string command, Room selectedRoom, DateTime endDate, int requestedFloor)
        {
            var endDateAsTime = endDate.ToShortTimeString();
            this.command = command;
            if (Contains("set us up the"))
            {
                return $"All your '{selectedRoom.RoomInfo.Name}' are belong to you until {endDateAsTime}";
            }
            else if (ContainsSwearWords())
            {
                return $"I have scheduled '{selectedRoom.RoomInfo.Name}' for you on floor {selectedRoom.RoomInfo.Floor} until {endDateAsTime} despite your potty mouth.";
            }
            else
            {
                return $"I have scheduled '{selectedRoom.RoomInfo.Name}' for you on floor {selectedRoom.RoomInfo.Floor} until {endDateAsTime}";
            }
        }

        private bool ContainsSwearWords()
        {
            string[] swearWords = new[] { "fuck", "shit", "ass", "bitch", "dick", "damn" };
            return swearWords.Any(n => Contains(n));
        }

        private bool Contains(string value)
        {
            return command.IndexOf(value, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }
    }
}