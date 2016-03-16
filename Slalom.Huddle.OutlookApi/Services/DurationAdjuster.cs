using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Slalom.Huddle.OutlookApi.Services
{
    public class DurationAdjuster
    {
        public DateTime ExtendDurationToNearestBlock(int duration)
        {
            // Figure out what current end date would be with existing duration.
            DateTime endDate = DateTime.Now.AddMinutes(duration);

            // If end date ends at a 15 minute increment of time (00, 15, 30, 45) then call duration good and return.
            if (endDate.Minute % 15 == 0)
            {
                endDate = endDate.AddSeconds(endDate.Second);
                endDate = endDate.AddMilliseconds(endDate.Millisecond);
                return endDate;
            }

            // Otherwise, find the next 15 minute increment higher than the current one, return a duration representing that difference.
            endDate = endDate.AddMinutes(15 - (endDate.Minute % 15));
            endDate = endDate.AddSeconds(-endDate.Second);
            endDate = endDate.AddMilliseconds(-endDate.Millisecond);

            return endDate;
        }
    }
}