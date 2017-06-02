using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartPhotoLibrary.Azure.Functions.Models
{
    public class VisionResponseModel
    {
        public IEnumerable<Category> categories {get;set;}

        public Adult adult { get; set; }

        public Description description { get; set; }

        public IEnumerable<Tag> tags { get; set; }

        public Color color { get; set; }
    }

    public class Description
    {
        public string[] tags { get; set; }
        public IEnumerable<Caption> captions { get; set; }
    }

    public class Color
    {
        public string dominantColorForeground { get; set; }
        public string dominantColorBackground { get; set; }
        public string accentColor { get; set; }
        public bool isBWImg { get; set; }

    }

    public class Caption
    {
        public string text { get; set; }
        public decimal confidence { get; set; }
    }

    public class Tag
    {
        public string name { get; set; }
        public decimal confidence { get; set; }
    }

    public class Category
    {
        public string name { get; set; }
        public decimal score { get; set; }
    }

    public class Adult
    {
        public bool isAdultContent { get; set; }
        public bool isRacyContent { get; set; }
        public decimal adultScore { get; set; }
        public decimal racyScore { get; set; }
    }

}
