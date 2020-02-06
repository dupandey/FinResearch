using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace FinResearch.Models
{
    [Table("LineItem")]
    public class LineItem
    {
        [Key]
        public long LineItemId { get; set; }
        public long ? ParentLineItemId { get; set; }
        public string LineItemText { get; set; }
        public Category Category { get; set; }
        public int CategoryId { get; set; }
        public bool IsActive { get; set; }
        public bool IsBold { get; set; }
        public bool IsDollar { get; set; }
        public bool IsUnderline { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
