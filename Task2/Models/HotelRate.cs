using System;
using System.Collections.Generic;
using System.Text;

namespace Task2
{
    public class HotelRate
    {
        public int adults { get; set; }
        public int los { get; set; }
        public Price price { get; set; }
        public string rateDescription { get; set; }
        public string rateID { get; set; }
        public string rateName { get; set; }
        public List<RateTag> rateTags { get; set; }
        public DateTime targetDay { get; set; }
    }
}
