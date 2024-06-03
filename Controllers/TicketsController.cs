using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using FORM_TICKETING_SYSTEM.Data;
using FORM_TICKETING_SYSTEM.Models;
using Microsoft.SqlServer.Server;
using OfficeOpenXml;

namespace FORM_TICKETING_SYSTEM.Controllers
{
    public class TicketsController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly string _excelFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "TicketDataForm.xlsx");

        public TicketsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: Tickets
        public async Task<IActionResult> Index()
        {
            return View(await _context.Ticket.ToListAsync());
        }

        // GET: Tickets/Details/5
        public async Task<IActionResult> Details(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var ticket = await _context.Ticket
                .FirstOrDefaultAsync(m => m.Id == id);
            if (ticket == null)
            {
                return NotFound();
            }

            return View(ticket);
        }

        // GET: Tickets/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Tickets/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,FirstName,MiddleName,LastName,Email,Description")] Ticket ticket)
        {
            if (ModelState.IsValid)
            {
                ticket.Id = Guid.NewGuid();
                _context.Add(ticket);
                await _context.SaveChangesAsync();
                return RedirectToAction("ExportToExcel", ticket);
            }
            return View(ticket);
        }
        [HttpPost]
        public IActionResult SubmitForm(Ticket ticket)
        {
            if (ModelState.IsValid)
            {
                _context.Ticket.Add(ticket);
                _context.SaveChanges();
                return RedirectToAction("ExportToExcel", new { id = ticket.Id });
            }
            return View(ticket);
        }

        public IActionResult ExportToExcel(Guid id)
        {
            var ticket = _context.Ticket.Find(id);
            if (ticket == null)
            {
                return NotFound();
            }

            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo existingFile = new FileInfo(_excelFilePath);
            ExcelPackage package;
            ExcelWorksheet worksheet;

            if (existingFile.Exists)
            {
                // Open the existing Excel file
                package = new ExcelPackage(existingFile);
                worksheet = package.Workbook.Worksheets["Ticket"];
            }
            else
            {
                // Create a new Excel file
                package = new ExcelPackage();
                worksheet = package.Workbook.Worksheets.Add("Ticket");
                worksheet.Cells["A1"].Value = "ID";
                worksheet.Cells["B1"].Value = "FIRST NAME";
                worksheet.Cells["C1"].Value = "MIDDLE NAME";
                worksheet.Cells["D1"].Value = "LAST NAME";
                worksheet.Cells["E1"].Value = "EMAIL";
                worksheet.Cells["F1"].Value = "DESCRIPTION";
            }

            var lastRow = worksheet.Dimension?.End.Row ?? 0;
            var newRow = lastRow + 1;

            worksheet.Cells[newRow, 1].Value = ticket.Id;
            worksheet.Cells[newRow, 2].Value = ticket.FirstName;
            worksheet.Cells[newRow, 3].Value = ticket.MiddleName;
            worksheet.Cells[newRow, 4].Value = ticket.LastName;
            worksheet.Cells[newRow, 5].Value = ticket.Email;
            worksheet.Cells[newRow, 6].Value = ticket.Description;

            var stream = new MemoryStream();
            package.SaveAs(stream);

            // Save the updated Excel file
            package.SaveAs(new FileInfo(_excelFilePath));

            var content = stream.ToArray();
            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Tickets.xlsx");
        }

        // GET: Tickets/Edit/5
        public async Task<IActionResult> Edit(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var ticket = await _context.Ticket.FindAsync(id);
            if (ticket == null)
            {
                return NotFound();
            }
            return View(ticket);
        }

        // POST: Tickets/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(Guid id, [Bind("Id,FirstName,MiddleName,LastName,Email,Description")] Ticket ticket)
        {
            if (id != ticket.Id)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(ticket);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!TicketExists(ticket.Id))
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
            return View(ticket);
        }

        // GET: Tickets/Delete/5
        public async Task<IActionResult> Delete(Guid? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var ticket = await _context.Ticket
                .FirstOrDefaultAsync(m => m.Id == id);
            if (ticket == null)
            {
                return NotFound();
            }

            return View(ticket);
        }

        // POST: Tickets/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(Guid id)
        {
            var ticket = await _context.Ticket.FindAsync(id);
            if (ticket != null)
            {
                _context.Ticket.Remove(ticket);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool TicketExists(Guid id)
        {
            return _context.Ticket.Any(e => e.Id == id);
        }
    }
 }
