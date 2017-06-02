using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;

namespace SmartPhotoLibrary.Azure.Functions.Models
{

    public class WebhookHistoryModel : TableEntity
    {
        public System.Guid Id { get; set; }
        public System.Guid ListId { get; set; }
        public string LastChangeToken { get; set; }
    }
}
