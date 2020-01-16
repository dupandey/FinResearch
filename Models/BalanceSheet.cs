using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace FinResearch.Models
{
    [Table("BalanceSheet")]
    public class BalanceSheet
    {
        [Key]
        public long BalanceSheetId { get; set; }
        //public FinanceStatement FinanceStatement { get; set; }
        public long StatementId { get; set; }
        //public LineItem LineItem { get; set; }
        public long LineItemId { get; set; }
        public string ItemValue { get; set; }        
        public bool IsActive { get; set; }
        public DateTime ? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
