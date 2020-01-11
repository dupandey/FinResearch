using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace FinResearch.Models
{
    [Table("CashFlow")]

    public class CashFlow
    {
        [Key]
        public long CashFlowId { get; set; }
        public FinanceStatement FinanceStatement { get; set; }
        public long StatementId { get; set; }
        public LineItem LineItem { get; set; }
        public long LineItemId { get; set; }
        public long? ItemValue { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
