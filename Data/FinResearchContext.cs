using System;
using System.Collections.Generic;

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

        public DbSet<FinResearch.Models.FinanceStatement> FinanceStatement { get; set; }
		public DbSet<FinResearch.Models.FinQuery> FinQuery { get; set; }
		protected override void OnModelCreating(ModelBuilder modelBuilder)
		{
			base.OnModelCreating(modelBuilder);
		}

		}
	}
