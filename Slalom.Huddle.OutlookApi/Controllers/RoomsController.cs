using System.Collections.Generic;
using System.Web.Http;

namespace Slalom.Huddle.OutlookApi.Controllers
{
    public class RoomsController : ApiController
    {
        // GET api/values
        public IEnumerable<string> Get()
        {
            return new string[] { "Room1", "Room2" };
        }        
    }
}
