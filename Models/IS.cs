using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace FinResearch.Models
{
    [Table("ISs")]
    public class ISs
    {
        [Key]
        public long ISId { get; set; }
        public int CompanyId { get; set; }
        public int CategoryId { get; set; }
        public string Payload { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
     
    }
}
