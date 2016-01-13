using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Xml;
using System.Xml.Serialization;

namespace Slalom.Huddle.OutlookApi.Models
{
    public class RoomInfo
    {
        public string Office { get; set; }

        public int Floor { get; set; }

        public string Name { get; set; }

        public int MaxPeople { get; set; }

        public bool HasAvEquipment { get; set; }

        public string Location { get; set; }

        public string Descriptor { get; set; }

        public static RoomInfo ParseRoomInfo(string roomDescriptor)
        {
            int indexOfBracket = roomDescriptor.IndexOf('[');
            string firstPart = roomDescriptor.Substring(0, indexOfBracket).Trim();
            string secondPart = roomDescriptor.Substring(indexOfBracket);
            string[] elements = firstPart.Split(' ');
            string office = elements[0];
            int floor = int.Parse(elements[1]);
            string name = string.Join(" ", elements.Skip(2));
            string[] subElements = secondPart.Substring(1, secondPart.Length-2).Split('/');
            int maxPeople = GetPeopleLimit(subElements);
            bool hasAvEquipment = GetAvEquipment(subElements);
            string location = subElements[subElements.Length-1];
            if (location == "AV" || location == maxPeople + "p")
            {
                location = "";
            }

            return new RoomInfo() {
                Office = office,
                Floor = floor,
                Name = name,
                MaxPeople = maxPeople,
                HasAvEquipment = hasAvEquipment,
                Location = location,
                Descriptor = roomDescriptor
            };
        }

        public override string ToString()
        {
            using (MemoryStream writer = new MemoryStream())
            {
                DataContractJsonSerializer contractSerializer = new DataContractJsonSerializer(typeof(RoomInfo));
                contractSerializer.WriteObject(writer, this);
                return System.Text.Encoding.Default.GetString(writer.ToArray());
            }
        }

        private static int GetPeopleLimit(string[] subElements)
        {
            try
            {
                int dontCare;
                var matchingSubElement = (from n in subElements
                                          where !n.Contains(' ') && n.EndsWith("p") && int.TryParse(n.Substring(0, n.Length - 1), out dontCare)
                                          select n).First();

                return int.Parse(matchingSubElement.Substring(0, matchingSubElement.Length - 1));
            }
            catch (Exception) {
                return 0;
            }
        }

        private static bool GetAvEquipment(string[] subElements)
        {
            return subElements.Any(n=>n == "AV");
        }
    }
}