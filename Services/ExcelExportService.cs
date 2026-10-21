
using ClosedXML.Excel;
using CRJSwaggerExcelExport_NET10.Data;
using Microsoft.EntityFrameworkCore;

namespace CRJSwaggerExcelExport_NET10.Services;

public class ExcelExportService(
	CrjDbContext db,
    SqlServerCrjRepository sqlServer,
    IConfiguration configuration,
	IWebHostEnvironment environment)
{
	private static readonly string[] TargetSheets =
	[
		"WELFARE",
		"HOUSING",
		"TRAINING",
		"LEGAL"
	];

	public async Task<byte[]> BuildWorkbookAsync(
		CancellationToken cancellationToken = default)
	{
		// Retrieve records from PostgreSQL.
		var records = string.Equals(configuration["Database:Provider"], "SqlServer", StringComparison.OrdinalIgnoreCase)
            ? (await sqlServer.GetRecordsAsync(cancellationToken)).SelectMany(h => h.Details).ToList()
            : await db.CrjDetails
			.AsNoTracking()
			.Include(x => x.Header)
			.OrderBy(x => x.RecordId)
			.ThenBy(x => x.Id)
			.ToListAsync(cancellationToken);

		var templatePath = Path.Combine(
			environment.ContentRootPath,
			"Templates",
			"CRJTemplate.xlsx");

		if (!File.Exists(templatePath))
		{
			templatePath = Path.Combine(
				AppContext.BaseDirectory,
				"Templates",
				"CRJTemplate.xlsx");
		}

		if (!File.Exists(templatePath))
		{
			throw new FileNotFoundException(
				"CRJ Excel template was not found.",
				templatePath);
		}

		using var workbook = new XLWorkbook(templatePath);

		// Validate that all required template sheets exist.
		foreach (var sheetName in TargetSheets)
		{
			if (!workbook.Worksheets
				.Any(x => string.Equals(
					x.Name,
					sheetName,
					StringComparison.OrdinalIgnoreCase)))
			{
				throw new InvalidOperationException(
					$"Required worksheet '{sheetName}' was not found in the template.");
			}
		}

		foreach (var sheetName in TargetSheets)
		{
			var worksheet = workbook.Worksheet(sheetName);

			var sheetRecords = records
				.Where(x => string.Equals(
					NormalizeType(x.ReportType ?? x.Header?.RecordCode),
					sheetName,
					StringComparison.OrdinalIgnoreCase))
				.ToList();

			var row = 10;

			foreach (var item in sheetRecords)
			{
				// A:AG - CRJ journal layout.
				Set(worksheet, row, 1, item.RcvdDateType);
				Set(worksheet, row, 2, item.DepositDate);
				Set(worksheet, row, 3, item.FileId);
				Set(worksheet, row, 4, item.Days);
				Set(worksheet, row, 5, item.Hours);
				Set(worksheet, row, 6, item.DueAmount);
				Set(worksheet, row, 7, item.AmountRecd);
				Set(worksheet, row, 8, item.AdjHrs);
				Set(worksheet, row, 9, item.AdjDue);
				Set(worksheet, row, 10, item.AdjPaid);
				Set(worksheet, row, 11, item.Interest);
				Set(worksheet, row, 12, item.LiqDamgs);
				Set(worksheet, row, 13, item.LegalFees);
				Set(worksheet, row, 14, item.Principal);

				// The model and database have only one Interest field.
				// Column O is left blank because there is no second Interest value.
				Set(worksheet, row, 15, null);

				Set(worksheet, row, 16, item.Claims);
				Set(worksheet, row, 17, item.Cost);
				Set(worksheet, row, 18, item.Legal);
				Set(worksheet, row, 19, item.TotalDeposit);
				Set(worksheet, row, 20, item.SystemDifference);
				Set(worksheet, row, 21, item.EmployerNo);
				Set(worksheet, row, 22, item.EmployerName);
				Set(worksheet, row, 23, item.ReportType ?? item.Header?.RecordCode);
				Set(worksheet, row, 24, item.WorkMonth);
				Set(worksheet, row, 25, item.Comments);
				Set(worksheet, row, 26, item.UnderPaid);
				Set(worksheet, row, 27, item.SideNotes);
				Set(worksheet, row, 28, item.BatchedBy);
				Set(worksheet, row, 29, item.BatchedByDate);
				Set(worksheet, row, 30, item.KeyedBy);
				Set(worksheet, row, 31, item.KeyedByDate);
				Set(worksheet, row, 32, item.CheckedBy);
				Set(worksheet, row, 33, item.CheckedByDate);

				row++;
			}
		}

		// Remove every sheet except the four requested sheets.
		var targetSheetNames = new HashSet<string>(
			TargetSheets,
			StringComparer.OrdinalIgnoreCase);

		foreach (var worksheet in workbook.Worksheets.ToList())
		{
			if (!targetSheetNames.Contains(worksheet.Name))
			{
				worksheet.Delete();
			}
		}

		using var stream = new MemoryStream();

		workbook.SaveAs(stream);

		return stream.ToArray();
	}

	private static string NormalizeType(string? value)
	{
		var type = (value ?? string.Empty)
			.Trim()
			.ToUpperInvariant();

		foreach (var target in TargetSheets)
		{
			if (type.Contains(
				target,
				StringComparison.OrdinalIgnoreCase))
			{
				return target;
			}
		}

		return type;
	}

	private static void Set(
		IXLWorksheet worksheet,
		int row,
		int column,
		object? value)
	{
		var cell = worksheet.Cell(row, column);

		if (value is null)
		{
			cell.Clear(XLClearOptions.Contents);
			return;
		}

		switch (value)
		{
			case DateTime date:
				cell.Value = date;
				break;

			case decimal number:
				cell.Value = (double)number;
				break;

			case int integer:
				cell.Value = integer;
				break;

			case double dbl:
				cell.Value = dbl;
				break;

			default:
				cell.Value = Convert.ToString(value) ?? string.Empty;
				break;
		}
	}
}
