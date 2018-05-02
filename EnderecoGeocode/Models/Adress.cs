using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class Address
    {
       // public List<AddressComponent> AddressComponents { get; set; }
        public string FormattedAddress { get; set; }
       // public Geometry Geometry { get; set; }
        public string PlaceId { get; set; }
        public List<string> Types { get; set; }
    }
}