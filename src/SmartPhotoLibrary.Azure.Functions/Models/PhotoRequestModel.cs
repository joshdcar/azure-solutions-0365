using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartPhotoLibrary.Azure.Functions.Models
{
    public class PhotoRequestModel
    {
        public Guid ListId{get;set;}
        public int Id { get; set; }
    }
}
