using System;
using System.Collections.Generic;
using FinResearch.Models;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace FinResearch.Models
{
    public class FinResearchContext : DbContext
    {
        public FinResearchContext (DbContextOptions<FinResearchContext> options)
            : base(options)
        {
        }

        public DbSet<Company> Companies { get; set; }
        public DbSet<Category> Categories { get; set; }
        public DbSet<LineItem> LineItems { get; set; }
        public DbSet<FinanceStatement> FinanceStatement { get; set; }
		public DbSet<FinQuery> FinQuery { get; set; }
        public DbSet<BalanceSheet> BalanceSheets { get; set; }
        public DbSet<CashFlow> CashFlows { get; set; }
        public DbSet<IS> ISs { get; set; }
        public DbSet<ISNonGAAP> ISNonGAAPs { get; set; }
        public DbSet<RD> RDs { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
		{
			base.OnModelCreating(modelBuilder);
		}

		}
	}
