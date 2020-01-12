using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using FinResearch.Models;

namespace FinResearch.Controllers
{
    public class FinanceStatementsController : Controller
    {
        private readonly FinResearchContext _context;

        public FinanceStatementsController(FinResearchContext context)
        {
            _context = context;
        }

        // GET: FinanceStatements
        public async Task<IActionResult> Index()
        {
			var p= await _context.FinQuery.FromSqlRaw("exec financedata").ToListAsync();
			return View(await _context.FinanceStatement.ToListAsync());
		
			
        }

		public ActionResult GetStatements()
		{
            //var p =  _context.FinQuery.FromSqlRaw("exec financedata").ToList();
            //return Json(p.FirstOrDefault().Records);
            return null;
		}

		// GET: FinanceStatements/Details/5
		public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var financeStatement = await _context.FinanceStatement
                .FirstOrDefaultAsync(m => m.StatementId == id);
            if (financeStatement == null)
            {
                return NotFound();
            }

            return View(financeStatement);
        }

        // GET: FinanceStatements/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: FinanceStatements/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("StatementId,Year")] FinanceStatement financeStatement)
        {
            if (ModelState.IsValid)
            {
                _context.Add(financeStatement);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(financeStatement);
        }

        // GET: FinanceStatements/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var financeStatement = await _context.FinanceStatement.FindAsync(id);
            if (financeStatement == null)
            {
                return NotFound();
            }
            return View(financeStatement);
        }

        // POST: FinanceStatements/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(long id, [Bind("StatementId,Year")] FinanceStatement financeStatement)
        {
            if (id != financeStatement.StatementId)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(financeStatement);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!FinanceStatementExists(financeStatement.StatementId))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(financeStatement);
        }

        // GET: FinanceStatements/Delete/5
        public async Task<IActionResult> Delete(long? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var financeStatement = await _context.FinanceStatement
                .FirstOrDefaultAsync(m => m.StatementId == id);
            if (financeStatement == null)
            {
                return NotFound();
            }

            return View(financeStatement);
        }

        // POST: FinanceStatements/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var financeStatement = await _context.FinanceStatement.FindAsync(id);
            _context.FinanceStatement.Remove(financeStatement);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool FinanceStatementExists(long id)
        {
            return _context.FinanceStatement.Any(e => e.StatementId == id);
        }
    }
}
