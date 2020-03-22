using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace FinResearch.Models
{
    [Table("RDs")]
    public class RDs
    {
        [Key]
        public long RDId { get; set; }
        public int CompanyId { get; set; }
        public int CategoryId { get; set; }
        public string Payload { get; set; }
        public string FileName { get; set; }
        public string UserId { get; set; }
        public bool? IsAdmin { get; set; }
        public int? FileVersion { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
