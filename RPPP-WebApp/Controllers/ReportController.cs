using RPPP_WebApp.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PdfRpt.ColumnsItemsTemplates;
using PdfRpt.Core.Contracts;
using PdfRpt.Core.Helper;
using PdfRpt.FluentInterface;
using System;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using RPPP_WebApp.Extensions;
using RPPP_WebApp.Controllers;
using RPPP_WebApp;
using SkiaSharp;
using RPPP_WebApp.ViewModels;

namespace MVC.Controllers
{
  public class ReportController : Controller
  {
        private readonly ProjektDbContext ctx;
        private readonly IWebHostEnvironment environment;
        private const string ExcelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public ReportController(ProjektDbContext  ctx, IWebHostEnvironment environment)
    {
      this.ctx = ctx;
      this.environment = environment;
    }

    public IActionResult Index()
    {
      return View();
    }

    [HttpGet]
    public async Task<IActionResult> ProjektPDF()
    {
      string naslov = "Popis projekata";
      var proj = await ctx.Projekt
                            .Include(d => d.IdVrsteNavigation)
                            .AsNoTracking()
                            .OrderBy(d => d.ImeProjekta)
                            .ToListAsync();
      PdfReport report = CreateReport(naslov);
      #region Podnožje i zaglavlje
      report.PagesFooter(footer =>
      {
        footer.DefaultFooter(DateTime.Now.ToString("dd.MM.yyyy."));
      })
      .PagesHeader(header =>
      {
        header.CacheHeader(cache: true); // It's a default setting to improve the performance.
              header.DefaultHeader(defaultHeader =>
        {
          defaultHeader.RunDirection(PdfRunDirection.LeftToRight);
          defaultHeader.Message(naslov);
        });
      });
      #endregion
      #region Postavljanje izvora podataka i stupaca
      report.MainTableDataSource(dataSource => dataSource.StronglyTypedList(proj));

      report.MainTableColumns(columns =>
      {
        columns.AddColumn(column =>
        {
          column.IsRowNumber(true);
          column.CellsHorizontalAlignment(HorizontalAlignment.Right);
          column.IsVisible(true);
          column.Order(0);
          column.Width(1);
          column.HeaderCell("#", horizontalAlignment: HorizontalAlignment.Right);
        });

        columns.AddColumn(column =>
        {
          column.PropertyName(nameof(Projekt.IdProjekta));
          column.CellsHorizontalAlignment(HorizontalAlignment.Center);
          column.IsVisible(true);
          column.Order(1);
          column.Width(1);
          column.HeaderCell("IdProjekta");
        });

        columns.AddColumn(column =>
        {
          column.PropertyName<Projekt>(x => x.ImeProjekta);
          column.CellsHorizontalAlignment(HorizontalAlignment.Center);
          column.IsVisible(true);
          column.Order(2);
          column.Width(2);
          column.HeaderCell("ImeProjekta", horizontalAlignment: HorizontalAlignment.Center);
        });

        columns.AddColumn(column =>
        {
          column.PropertyName<Projekt>(x => x.Kratica);
          column.CellsHorizontalAlignment(HorizontalAlignment.Center);
          column.IsVisible(true);
          column.Order(3);
          column.Width(2);
          column.HeaderCell("Kratica", horizontalAlignment: HorizontalAlignment.Center);
        });

        columns.AddColumn(column =>
        {
          column.PropertyName<Projekt>(x => x.Sazetak);
          column.CellsHorizontalAlignment(HorizontalAlignment.Center);
          column.IsVisible(true);
          column.Order(4);
          column.Width(5);
          column.HeaderCell("Sazetak", horizontalAlignment: HorizontalAlignment.Center);
        });

        columns.AddColumn(column =>
        {
            column.PropertyName<Projekt>(x => x.DatumPoc);
            column.CellsHorizontalAlignment(HorizontalAlignment.Center);
            column.IsVisible(true);
            column.Order(5);
            column.Width(1);
            column.HeaderCell("DatumPoc", horizontalAlignment: HorizontalAlignment.Center);
        });

        columns.AddColumn(column =>
        {
            column.PropertyName<Projekt>(x => x.DatumZav);
            column.CellsHorizontalAlignment(HorizontalAlignment.Center);
            column.IsVisible(true);
            column.Order(6);
            column.Width(1);
            column.HeaderCell("DatumZav", horizontalAlignment: HorizontalAlignment.Center);
        });
          columns.AddColumn(column =>
          {
              column.PropertyName<Projekt>(x => x.BrKartice);
              column.CellsHorizontalAlignment(HorizontalAlignment.Center);
              column.IsVisible(true);
              column.Order(7);
              column.Width(2);
              column.HeaderCell("BrKartice", horizontalAlignment: HorizontalAlignment.Center);
          });
          columns.AddColumn(column =>
          {
              column.PropertyName<Projekt>(x => x.IdVrsteNavigation.ImeVrste);
              column.CellsHorizontalAlignment(HorizontalAlignment.Center);
              column.IsVisible(true);
              column.Order(8);
              column.Width(2);
              column.HeaderCell("ImeVrste", horizontalAlignment: HorizontalAlignment.Center);
          });
      });



      #endregion
      byte[] pdf = report.GenerateAsByteArray();

      if (pdf != null)
      {
        Response.Headers.Add("content-disposition", "inline; filename=drzave.pdf");
        //return File(pdf, "application/pdf");
        return File(pdf, "application/pdf", "drzave.pdf"); 
      }
      else
      {
        return NotFound();
      }
    }

    [HttpGet]
    public async Task<IActionResult> ProjektPDF2(int id)
    {
      string naslov = "Projekt i njegove dokumentacije";
            var entitet = await ctx.Projekt
                                          .Where(d => d.IdProjekta == id)
                                          .Include(d => d.IdVrsteNavigation)
                                          .AsNoTracking()
                                          .OrderBy(d => d.ImeProjekta)
                                          .SingleOrDefaultAsync();

            var entitet2 = await ctx.Dokumentacija
                                .Where(d => d.IdProjekta == id)
                                .Include(d => d.IdProjektaNavigation)
                                .Include(d => d.IdVrsteNavigation)
                                .AsNoTracking()
                                .OrderBy(d => d.ImeDok)
                                .ToListAsync();



            PdfReport report = CreateReport(naslov);

            #region Podnožje i zaglavlje
            report.PagesFooter(footer =>
            {
                footer.DefaultFooter(DateTime.Now.ToString("dd.MM.yyyy."));
            })
            .PagesHeader(header =>
            {
                header.HtmlHeader(rptHeader =>
                {
                    rptHeader.PageHeaderProperties(new HeaderBasicProperties
                    {
                        RunDirection = PdfRunDirection.LeftToRight,
                        ShowBorder = true,
                        PdfFont = header.PdfFont
                    });
                    rptHeader.AddPageHeader(pageHeader =>
                    {
                        var message = $"MD za {entitet.ImeProjekta}";
                        return string.Format(@"<table style='width: 100%;font-size:9pt;font-family:tahoma;'>
													<tr>
														<td align='center'>{0}</td>
													</tr>
												</table>", message);
                    });

                    rptHeader.GroupHeaderProperties(new HeaderBasicProperties
                    {
                        RunDirection = PdfRunDirection.LeftToRight,
                        ShowBorder = true,
                        SpacingBeforeTable = 10f,
                        PdfFont = header.PdfFont
                    });
                    rptHeader.AddGroupHeader(groupHeader =>
                    {
                        var idProjekta = entitet?.IdProjekta ?? 0;
                        var imeProjekta = entitet?.ImeProjekta ?? "NEMA DATOTEKA!";
                        var kratica = entitet.Kratica;
                        var sazetak = entitet.Sazetak;
                        var datumPoc = entitet.DatumPoc;
                        var datumkraj = entitet.DatumZav;
                        var vrsta = entitet.IdVrsteNavigation?.ImeVrste != null ? entitet.IdVrsteNavigation.ImeVrste : "Ne pripada nijednoj vrsti";
                        return string.Format(@"<table style='width: 100%; font-size:9pt;font-family:tahoma;'>
															<tr>
																<td style='width:25%;border-bottom-width:0.2; border-bottom-color:red;border-bottom-style:solid'>Id projekta:</td>
																<td style='width:75%'>{0}</td>
															</tr>
															<tr>
																<td style='width:25%'>Ime projekta:</td>
																<td style='width:75%'>{1}</td>
															</tr>
																<td style='width:25%'>Kratica:</td>
																<td style='width:75%'>{2}</td>
															</tr>
																<td style='width:25%'>Sazetak:</td>
																<td style='width:75%'>{3}</td>
															</tr>
                                                            </tr>
																<td style='width:25%'>Datum pocetka:</td>
																<td style='width:75%'>{4}</td>
															</tr>
                                                            </tr>
																<td style='width:25%'>Datum zavrsetka:</td>
																<td style='width:75%'>{5}</td>
															</tr>
                                                            </tr>
																<td style='width:25%'>Ime vrste:</td>
																<td style='width:75%'>{6}</td>
															</tr>
												</table>", idProjekta, imeProjekta, kratica, sazetak, datumPoc, datumkraj, vrsta);
                    });
                });
            });
            #endregion

            #region Postavljanje izvora podataka i stupaca
            report.MainTableDataSource(dataSource => dataSource.StronglyTypedList(entitet2));

            report.MainTableColumns(columns =>
            {
                #region Stupci po kojima se grupira
                columns.AddColumn(column =>
                {
                    column.PropertyName<Projekt>(s => s.IdProjekta);
                    column.Group(
                        (val1, val2) =>
                        {
                            return (int)val1 == (int)val2;
                        });
                });
                #endregion
                columns.AddColumn(column =>
                {
                    column.IsRowNumber(true);
                    column.CellsHorizontalAlignment(HorizontalAlignment.Right);
                    column.IsVisible(true);
                    column.Order(0);
                    column.Width(1);
                    column.HeaderCell("#", horizontalAlignment: HorizontalAlignment.Right);
                });
                columns.AddColumn(column =>
                {
                    column.PropertyName<Dokumentacija>(x => x.IdDok);
                    column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                    column.IsVisible(true);
                    column.Order(1);
                    column.Width(1);
                    column.HeaderCell("IdDok", horizontalAlignment: HorizontalAlignment.Center);
                });

                columns.AddColumn(column =>
                {
                    column.PropertyName<Dokumentacija>(x => x.ImeDok);
                    column.CellsHorizontalAlignment(HorizontalAlignment.Left);
                    column.IsVisible(true);
                    column.Order(2);
                    column.Width(4);
                    column.HeaderCell("ImeDok", horizontalAlignment: HorizontalAlignment.Center);
                });

                columns.AddColumn(column =>
                {
                    column.PropertyName<Dokumentacija>(x => x.IdVrsteNavigation.ImeVrste);
                    column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                    column.IsVisible(true);
                    column.Order(3);
                    column.Width(3);
                    column.HeaderCell("ImeVrste", horizontalAlignment: HorizontalAlignment.Center);
                });

                columns.AddColumn(column =>
                {
                    column.PropertyName<Dokumentacija>(x => x.IdProjektaNavigation.ImeProjekta);
                    column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                    column.IsVisible(true);
                    column.Order(4);
                    column.Width(4);
                    column.HeaderCell("ImeProjekta", horizontalAlignment: HorizontalAlignment.Center);
                });
            });
            #endregion

            byte[] pdf = report.GenerateAsByteArray();

            if (pdf != null)
            {
                Response.Headers.Add("content-disposition", "inline; filename=master.pdf");
                return File(pdf, "application/pdf");
            }
            else
            {
                return NotFound();
            }

        }

        [HttpGet]
    public async Task<IActionResult> VrstaProjektaPDF()
    {
        string naslov = "Popis vrsta projekta";
        var vrstaProj = await ctx.VrstaProjekta
                                .AsNoTracking()
                                .OrderBy(d => d.ImeVrste)
                                .ToListAsync();
        PdfReport report = CreateReport(naslov);
        #region Podnožje i zaglavlje
        report.PagesFooter(footer =>
        {
            footer.DefaultFooter(DateTime.Now.ToString("dd.MM.yyyy."));
        })
        .PagesHeader(header =>
        {
            header.CacheHeader(cache: true); // It's a default setting to improve the performance.
            header.DefaultHeader(defaultHeader =>
            {
                defaultHeader.RunDirection(PdfRunDirection.LeftToRight);
                defaultHeader.Message(naslov);
            });
        });
        #endregion
        #region Postavljanje izvora podataka i stupaca
        report.MainTableDataSource(dataSource => dataSource.StronglyTypedList(vrstaProj));

        report.MainTableColumns(columns =>
        {
            columns.AddColumn(column =>
            {
                column.IsRowNumber(true);
                column.CellsHorizontalAlignment(HorizontalAlignment.Right);
                column.IsVisible(true);
                column.Order(0);
                column.Width(1);
                column.HeaderCell("#", horizontalAlignment: HorizontalAlignment.Right);
            });

            columns.AddColumn(column =>
            {
                column.PropertyName(nameof(VrstaProjekta.IdVrste));
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(1);
                column.Width(1);
                column.HeaderCell("IdVrsteProjekta");
            });

            columns.AddColumn(column =>
            {
                column.PropertyName<VrstaProjekta>(x => x.ImeVrste);
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(2);
                column.Width(2);
                column.HeaderCell("Ime vrste", horizontalAlignment: HorizontalAlignment.Center);
            });
        });



        #endregion
        byte[] pdf = report.GenerateAsByteArray();

        if (pdf != null)
        {
            Response.Headers.Add("content-disposition", "inline; filename=drzave.pdf");
            //return File(pdf, "application/pdf");
            return File(pdf, "application/pdf", "drzave.pdf");
        }
        else
        {
            return NotFound();
        }
    }

    [HttpGet]
    public async Task<IActionResult> DokumentacijaPDF()
    {
        string naslov = "Popis dokumentacija";
        var doku = await ctx.Dokumentacija
                                .Include(d=>d.IdProjektaNavigation)
                                .Include(d => d.IdVrsteNavigation)
                                .AsNoTracking()
                                .OrderBy(d => d.ImeDok)
                                .ToListAsync();
        PdfReport report = CreateReport(naslov);
        #region Podnožje i zaglavlje
        report.PagesFooter(footer =>
        {
            footer.DefaultFooter(DateTime.Now.ToString("dd.MM.yyyy."));
        })
        .PagesHeader(header =>
        {
            header.CacheHeader(cache: true); // It's a default setting to improve the performance.
            header.DefaultHeader(defaultHeader =>
            {
                defaultHeader.RunDirection(PdfRunDirection.LeftToRight);
                defaultHeader.Message(naslov);
            });
        });
        #endregion
        #region Postavljanje izvora podataka i stupaca
        report.MainTableDataSource(dataSource => dataSource.StronglyTypedList(doku));

        report.MainTableColumns(columns =>
        {
            columns.AddColumn(column =>
            {
                column.IsRowNumber(true);
                column.CellsHorizontalAlignment(HorizontalAlignment.Right);
                column.IsVisible(true);
                column.Order(0);
                column.Width(1);
                column.HeaderCell("#", horizontalAlignment: HorizontalAlignment.Right);
            });

            columns.AddColumn(column =>
            {
                column.PropertyName(nameof(Dokumentacija.IdDok));
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(1);
                column.Width(1);
                column.HeaderCell("IdDok");
            });

            columns.AddColumn(column =>
            {
                column.PropertyName<Dokumentacija>(x => x.ImeDok);
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(2);
                column.Width(2);
                column.HeaderCell("ImeDok", horizontalAlignment: HorizontalAlignment.Center);
            });

            columns.AddColumn(column =>
            {
                column.PropertyName<Dokumentacija>(x => x.IdVrsteNavigation.ImeVrste);
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(3);
                column.Width(2);
                column.HeaderCell("ImeVrste", horizontalAlignment: HorizontalAlignment.Center);
            });

            columns.AddColumn(column =>
            {
                column.PropertyName<Dokumentacija>(x => x.IdProjektaNavigation.ImeProjekta);
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(4);
                column.Width(2);
                column.HeaderCell("ImeProjekta", horizontalAlignment: HorizontalAlignment.Center);
            });

        });



        #endregion
        byte[] pdf = report.GenerateAsByteArray();

        if (pdf != null)
        {
            Response.Headers.Add("content-disposition", "inline; filename=drzave.pdf");
            //return File(pdf, "application/pdf");
            return File(pdf, "application/pdf", "drzave.pdf");
        }
        else
        {
            return NotFound();
        }
    }

    [HttpGet]
    public async Task<IActionResult> VrstaDokumentacijePDF()
    {
        string naslov = "Popis vrsta dokumentacije";
        var vrstaDok = await ctx.VrstaDok
                                .AsNoTracking()
                                .OrderBy(d => d.ImeVrste)
                                .ToListAsync();
        PdfReport report = CreateReport(naslov);
        #region Podnožje i zaglavlje
        report.PagesFooter(footer =>
        {
            footer.DefaultFooter(DateTime.Now.ToString("dd.MM.yyyy."));
        })
        .PagesHeader(header =>
        {
            header.CacheHeader(cache: true); // It's a default setting to improve the performance.
            header.DefaultHeader(defaultHeader =>
            {
                defaultHeader.RunDirection(PdfRunDirection.LeftToRight);
                defaultHeader.Message(naslov);
            });
        });
        #endregion
        #region Postavljanje izvora podataka i stupaca
        report.MainTableDataSource(dataSource => dataSource.StronglyTypedList(vrstaDok));

        report.MainTableColumns(columns =>
        {
            columns.AddColumn(column =>
            {
                column.IsRowNumber(true);
                column.CellsHorizontalAlignment(HorizontalAlignment.Right);
                column.IsVisible(true);
                column.Order(0);
                column.Width(1);
                column.HeaderCell("#", horizontalAlignment: HorizontalAlignment.Right);
            });

            columns.AddColumn(column =>
            {
                column.PropertyName(nameof(VrstaUloge.IdVrste));
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(1);
                column.Width(1);
                column.HeaderCell("Id vrste dokumentacije");
            });

            columns.AddColumn(column =>
            {
                column.PropertyName<VrstaUloge>(x => x.ImeVrste);
                column.CellsHorizontalAlignment(HorizontalAlignment.Center);
                column.IsVisible(true);
                column.Order(2);
                column.Width(2);
                column.HeaderCell("Ime vrste dokumentacije", horizontalAlignment: HorizontalAlignment.Center);
            });


        });



        #endregion
        byte[] pdf = report.GenerateAsByteArray();

        if (pdf != null)
        {
            Response.Headers.Add("content-disposition", "inline; filename=drzave.pdf");
            //return File(pdf, "application/pdf");
            return File(pdf, "application/pdf", "drzave.pdf");
        }
        else
        {
            return NotFound();
        }
    }

    [HttpGet]
    public async Task<IActionResult> ProjektExcel()
    {
        var entitet = await ctx.Projekt
                                .Include(d => d.IdVrsteNavigation)
                                .AsNoTracking()
                                .OrderBy(d => d.ImeProjekta)
                                .ToListAsync();
        byte[] content;
        using (ExcelPackage excel = new ExcelPackage())
        {
            excel.Workbook.Properties.Title = "Popis projekta";
            excel.Workbook.Properties.Author = "rppp-02";
            var worksheet = excel.Workbook.Worksheets.Add("Projekt");

            worksheet.Cells[1, 1].Value = "Ime projekta";
            worksheet.Cells[1, 2].Value = "Kratica";
            worksheet.Cells[1, 3].Value = "Id projekta";
            worksheet.Cells[1, 4].Value = "Sažetak";
            worksheet.Cells[1, 5].Value = "DatumPoc";
            worksheet.Cells[1, 6].Value = "DatumZav";
            worksheet.Cells[1, 7].Value = "Broj kartice";
            worksheet.Cells[1, 8].Value = "Ime vrste";

            for (int i = 0; i < entitet.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = entitet[i].ImeProjekta;
                worksheet.Cells[i + 2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[i + 2, 2].Value = entitet[i].Kratica;
                worksheet.Cells[i + 2, 3].Value = entitet[i].IdProjekta;
                worksheet.Cells[i + 2, 4].Value = entitet[i].Sazetak;
                worksheet.Cells[i + 2, 5].Value = entitet[i].DatumPoc.ToString();
                worksheet.Cells[i + 2, 6].Value = entitet[i].DatumZav.ToString();
                worksheet.Cells[i + 2, 7].Value = entitet[i].BrKartice.ToString();
                worksheet.Cells[i + 2, 8].Value = entitet[i].IdVrsteNavigation.ImeVrste;
                }

            worksheet.Cells[1, 1, entitet.Count + 1, 6].AutoFitColumns();

            content = excel.GetAsByteArray();
        }
        return File(content, ExcelContentType, "projekt.xlsx");
    }

        [HttpGet]
        public async Task<IActionResult> ProjektExcel2(int id)
        {
            var entitet = await ctx.Projekt
                                    .Where(d => d.IdProjekta == id)
                                    .Include(d => d.IdVrsteNavigation)
                                    .AsNoTracking()
                                    .OrderBy(d => d.ImeProjekta)
                                    .SingleOrDefaultAsync();

            var entitet2 = await ctx.Dokumentacija
                                .Where(d=>d.IdProjekta==id)
                                .Include(d => d.IdProjektaNavigation)
                                .Include(d => d.IdVrsteNavigation)
                                .AsNoTracking()
                                .OrderBy(d => d.ImeDok)
                                .ToListAsync();
            byte[] content;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Properties.Title = "Popis projekta i njegovih dokumentacija";
                excel.Workbook.Properties.Author = "rppp-02";
                var worksheet = excel.Workbook.Worksheets.Add("ProjektDokMD");

                worksheet.Cells[1, 1].Value = "Ime projekta";
                worksheet.Cells[1, 2].Value = "Kratica";
                worksheet.Cells[1, 3].Value = "Id projekta";
                worksheet.Cells[1, 4].Value = "Sažetak";
                worksheet.Cells[1, 5].Value = "DatumPoc";
                worksheet.Cells[1, 6].Value = "DatumZav";
                worksheet.Cells[1, 7].Value = "Broj kartice";
                worksheet.Cells[1, 8].Value = "Ime vrste";

                worksheet.Cells[2, 1].Value = entitet.ImeProjekta;
                worksheet.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[2, 2].Value = entitet.Kratica;
                worksheet.Cells[2, 3].Value = entitet.IdProjekta;
                worksheet.Cells[2, 4].Value = entitet.Sazetak;
                worksheet.Cells[2, 5].Value = entitet.DatumPoc.ToString();
                worksheet.Cells[2, 6].Value = entitet.DatumZav.ToString();
                worksheet.Cells[2, 7].Value = entitet.BrKartice.ToString();
                worksheet.Cells[2, 8].Value = entitet.IdVrsteNavigation.ImeVrste;

                worksheet.Cells[4, 1].Value = "Ime dokumentacije";
                worksheet.Cells[4, 2].Value = "Id dokumentacije";
                worksheet.Cells[4, 3].Value = "Ime projekta";
                worksheet.Cells[4, 4].Value = "Ime vrste dokumenta";

                for (int i = 0; i < entitet2.Count; i++)
                {
                    int row = i + 5;
                    worksheet.Cells[row, 1].Value = entitet2[i].ImeDok;
                    worksheet.Cells[row, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells[row, 2].Value = entitet2[i].IdDok;
                    worksheet.Cells[row, 3].Value = entitet2[i].IdProjektaNavigation?.ImeProjekta != null ? entitet2[i].IdProjektaNavigation.ImeProjekta : "Ne pripada nijednom projektu";
                    worksheet.Cells[row, 4].Value = entitet2[i].IdVrsteNavigation?.ImeVrste != null ? entitet2[i].IdVrsteNavigation?.ImeVrste : "Nema vrstu";
                }

                worksheet.Cells[1, 1, entitet2.Count + 5, 6].AutoFitColumns();

                content = excel.GetAsByteArray();
            }
            return File(content, ExcelContentType, "projektMD.xlsx");
        }

        [HttpGet]
    public async Task<IActionResult> VrstaProjektaExcel()
    {
        var entitet = await ctx.VrstaProjekta
                                .AsNoTracking()
                                .OrderBy(d => d.ImeVrste)
                                .ToListAsync();
        byte[] content;
        using (ExcelPackage excel = new ExcelPackage())
        {
            excel.Workbook.Properties.Title = "Popis vrsta projekta";
            excel.Workbook.Properties.Author = "rppp-02";
            var worksheet = excel.Workbook.Worksheets.Add("Vrsta Projekta");

            worksheet.Cells[1, 1].Value = "Ime vrste";

                for (int i = 0; i < entitet.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = entitet[i].ImeVrste;
                worksheet.Cells[i + 2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }

            worksheet.Cells[1, 1, entitet.Count + 1, 6].AutoFitColumns();

            content = excel.GetAsByteArray();
        }
        return File(content, ExcelContentType, "vrsteProjekta.xlsx");
    }
    
    [HttpGet]
    public async Task<IActionResult> DokumentacijaExcel()
    {
        var entitet = await ctx.Dokumentacija
                                .Include(d => d.IdProjektaNavigation)
                                .Include(d => d.IdVrsteNavigation)
                                .AsNoTracking()
                                .OrderBy(d => d.ImeDok)
                                .ToListAsync();

        byte[] content;
        using (ExcelPackage excel = new ExcelPackage())
        {
            excel.Workbook.Properties.Title = "Popis dokumentacija";
            excel.Workbook.Properties.Author = "rppp-02";
            var worksheet = excel.Workbook.Worksheets.Add("Dokumentacije");

            worksheet.Cells[1, 1].Value = "Ime dokumentacije";
            worksheet.Cells[1, 2].Value = "Id dokumentacije";
            worksheet.Cells[1, 3].Value = "Ime projekta";
            worksheet.Cells[1, 4].Value = "Ime vrste dokumenta";

            for (int i = 0; i < entitet.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = entitet[i].ImeDok;
                worksheet.Cells[i + 2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[i + 2, 2].Value = entitet[i].IdDok;
                worksheet.Cells[i + 2, 3].Value = entitet[i].IdProjektaNavigation?.ImeProjekta != null ? entitet[i].IdProjektaNavigation.ImeProjekta : "Ne pripada nijednom projektu";
                worksheet.Cells[i + 2, 4].Value = entitet[i].IdVrsteNavigation?.ImeVrste != null ? entitet[i].IdVrsteNavigation?.ImeVrste : "Nema vrstu";
                }

            worksheet.Cells[1, 1, entitet.Count + 1, 3].AutoFitColumns();

            content = excel.GetAsByteArray();
        }
        return File(content, ExcelContentType, "dokumentacija.xlsx");
    }

    [HttpGet]
    public async Task<IActionResult> VrstaDokumentacijeExcel()
    {
        var entitet = await ctx.VrstaDok
                                .AsNoTracking()
                                .OrderBy(d => d.ImeVrste)
                                .ToListAsync();
        byte[] content;
        using (ExcelPackage excel = new ExcelPackage())
        {
            excel.Workbook.Properties.Title = "Popis vrsta dokumentacija";
            excel.Workbook.Properties.Author = "rppp-02";
            var worksheet = excel.Workbook.Worksheets.Add("Vrste dokumentacija");

            worksheet.Cells[1, 1].Value = "Ime vrste dokumentacije";
            worksheet.Cells[1, 2].Value = "Id vrste dokumentacije";

            for (int i = 0; i < entitet.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = entitet[i].ImeVrste;
                worksheet.Cells[i + 2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                worksheet.Cells[i + 2, 2].Value = entitet[i].IdVrste;
            }

            worksheet.Cells[1, 1, entitet.Count + 1, 2].AutoFitColumns();

            content = excel.GetAsByteArray();
        }
        return File(content, ExcelContentType, "vrstedokumentacije.xlsx");
    }
    
        
    
    private PdfReport CreateReport(string naslov)
    {
      var pdf = new PdfReport();

      pdf.DocumentPreferences(doc =>
      {
        doc.Orientation(PageOrientation.Portrait);
        doc.PageSize(PdfPageSize.A4);
        doc.DocumentMetadata(new DocumentMetadata
        {
          Author = "rppp02",
          Application = "RPPP_WebApp Core",
          Title = naslov
        });
        doc.Compression(new CompressionSettings
        {
          EnableCompression = true,
          EnableFullCompression = true
        });
      })
      //fix za linux https://github.com/VahidN/PdfReport.Core/issues/40
      .DefaultFonts(fonts => {
        fonts.Path(Path.Combine(environment.WebRootPath, "fonts", "verdana.ttf"),
                         Path.Combine(environment.WebRootPath, "fonts", "tahoma.ttf"));
        fonts.Size(9);
        fonts.Color(System.Drawing.Color.Black);
      })
      //
      .MainTableTemplate(template =>
      {
        template.BasicTemplate(BasicTemplate.ProfessionalTemplate);
      })
      .MainTablePreferences(table =>
      {
        table.ColumnsWidthsType(TableColumnWidthType.Relative);
              //table.NumberOfDataRowsPerPage(20);
              table.GroupsPreferences(new GroupsPreferences
        {
          GroupType = GroupType.HideGroupingColumns,
          RepeatHeaderRowPerGroup = true,
          ShowOneGroupPerPage = true,
          SpacingBeforeAllGroupsSummary = 5f,
          NewGroupAvailableSpacingThreshold = 150,
          SpacingAfterAllGroupsSummary = 5f
        });
        table.SpacingAfter(4f);
      });

      return pdf;
    }
  }
}
