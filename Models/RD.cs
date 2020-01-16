﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace FinResearch.Models
{
    [Table("RD")]
    public class RD
    {
        [Key]
        public long RDId { get; set; }
        public long StatementId { get; set; }
        public long LineItemId { get; set; }
        public string ItemValue { get; set; }
        public bool IsActive { get; set; }
        public DateTime? CreatedDate { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
