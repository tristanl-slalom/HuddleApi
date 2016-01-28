using System;
using System.Linq;
using Slalom.Huddle.OutlookApi.Models;

namespace Slalom.Huddle.OutlookApi.Services
{
    internal class CommandResponder
    {
        private string command;

        public string CreateResponseForCommand(string command, Room selectedRoom, int duration, int requestedFloor)
        {
            this.command = command;
            if (Contains("set us up the"))
            {
                return $"All your '{selectedRoom.RoomInfo.Name}' are belong to you for the next {duration} minutes";
            }
            else if (ContainsSwearWords())
            {
                return $"I have scheduled '{selectedRoom.RoomInfo.Name}' for you on floor {selectedRoom.RoomInfo.Floor} for the next {duration} minutes despite your potty mouth.";
            }
            else
            {
                return $"I have scheduled '{selectedRoom.RoomInfo.Name}' for you on floor {selectedRoom.RoomInfo.Floor} for the next {duration} minutes";
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