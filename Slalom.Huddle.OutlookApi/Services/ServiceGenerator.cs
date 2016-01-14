using Microsoft.Exchange.WebServices.Data;
using System;
using System.Net;

namespace Slalom.Huddle.OutlookApi.Services
{
    public class ServiceGenerator
    {
        public ExchangeService GenerateService(string email, string password)
        {
            // Set up the exchange service for UTC
            ExchangeService service = new ExchangeService(TimeZoneInfo.Utc);

            // Create network credentials and specify the public office 365 URL
            service.Credentials = new NetworkCredential(email, password);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            return service;
        }
    }
}