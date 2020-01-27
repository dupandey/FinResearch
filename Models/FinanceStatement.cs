using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace FinResearch.Models
{
	[Table("FinanceStatement")]
	public class FinanceStatement
	{
        //public int Id { get; set; }

        //public int Year { get; set; }
        [Key]
        public long StatementId { get; set; }
        public int CompanyId { get; set; }
        public DateTime TransactionDate { get; set; }
        public string Quarter { get; set; }
        public bool? IsHistorical { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public int? OrderNo { get; set; }

    }

    [Table("Company")]
    public class Company
    {
        [Key]
        public int CompanyId { get; set; }
        public string CompanyCode { get; set; }
        public string Address { get; set; }
        public string CompanyName { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
