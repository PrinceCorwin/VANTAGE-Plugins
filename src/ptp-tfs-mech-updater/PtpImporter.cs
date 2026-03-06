using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;
using VANTAGE;
using VANTAGE.Models;
using VANTAGE.Services.Plugins;
using Microsoft.Data.Sqlite;

namespace PtpTfsMechUpdater
{
    // Aggregated data per CWP from the PTP report
    internal class CwpData
    {
        public string Cwp { get; set; } = "";
        public double TotalQuantity { get; set; }
        public double TotalShippedQty { get; set; }
        public bool AllDeleted { get; set; }
        public DateTime? MinShipDate { get; set; }
        public DateTime? MaxShipDate { get; set; }
    }

    // Existing activity data needed for comparison
    internal class ExistingRecord
    {
        public string UniqueID { get; set; } = "";
        public string AssignedTo { get; set; } = "";
        public string Notes { get; set; } = "";
        public double Quantity { get; set; }
        public double PercentEntry { get; set; }
        public string ActStart { get; set; } = "";
        public string ActFin { get; set; } = "";
    }

    internal class PtpImporter
    {
        private readonly IPluginHost _host;

        // Required column headers (matched case-insensitive, whitespace-normalized)
        private static readonly string[] RequiredHeaders = { "CWP", "Quantity", "Status", "Shipped QTY" };

        // Constants for activity creation
        private const string ROCStep = "7.SHP";
        private const string DescriptionPrefix = "FABRICATION - 7.SHP ";
        private const string Area = "TFS";
        private const string CompType = "P";
        private const string PhaseCategory = "PSF";
        private const string PhaseCode = "xx.xxx.xxx.";
        private const string ProjectID = "25.005.";
        private const string ShopField = "Shop";
        private const string UDF1 = "1";
        private const string UDF3 = "NEARSITE";
        private const string WorkPackage = "x";
        private const string RespParty = "SUMMIT - PM";
        private const string SchedActNO = "x";
        private const double BudgetMHs = 0.001;

        public PtpImporter(IPluginHost host)
        {
            _host = host;
        }

        public async Task RunAsync(string filePath)
        {
            // Parse the Excel file
            var cwpDataList = ParseReport(filePath);
            if (cwpDataList == null) return;

            _host.LogInfo($"Parsed {cwpDataList.Count} unique CWPs from PTP report", "PtpImporter.RunAsync");

            // Look up existing activities by description
            var existingRecords = await FindExistingRecordsAsync();

            // Ownership check — reject if any existing records belong to another user
            var currentUser = _host.CurrentUsername;
            var foreignOwners = existingRecords.Values
                .Where(r => !string.Equals(r.AssignedTo, currentUser, StringComparison.OrdinalIgnoreCase))
                .Select(r => r.AssignedTo)
                .Distinct()
                .ToList();

            if (foreignOwners.Count > 0)
            {
                _host.ShowError(
                    $"Cannot update records. The following user(s) own existing PTP activities:\n\n" +
                    $"{string.Join(", ", foreignOwners)}\n\n" +
                    $"Only the original importer can update these records.",
                    "Ownership Conflict");
                return;
            }

            // Process: create new or update existing
            int created = 0;
            int updated = 0;
            int unchanged = 0;
            var deletedCwps = new List<string>();
            var today = DateTime.Now;

            await Task.Run(() =>
            {
                using var connection = DatabaseSetup.GetConnection();
                connection.Open();
                using var transaction = connection.BeginTransaction();

                try
                {
                    var timestamp = DateTime.Now.ToString("yyMMddHHmmss");
                    var userSuffix = currentUser.Length >= 3
                        ? currentUser.Substring(currentUser.Length - 3).ToLower()
                        : "usr";
                    int sequence = 1;

                    foreach (var cwp in cwpDataList)
                    {
                        var description = $"{DescriptionPrefix}{cwp.Cwp}";

                        // Calculate percent entry
                        double percentEntry;
                        if (cwp.AllDeleted)
                        {
                            percentEntry = 100;
                            deletedCwps.Add(cwp.Cwp);
                        }
                        else
                        {
                            percentEntry = cwp.TotalQuantity > 0
                                ? Math.Round((cwp.TotalShippedQty / cwp.TotalQuantity) * 100, 3)
                                : 0;
                        }

                        // Notes for all-deleted case
                        string notes = "";
                        if (cwp.AllDeleted)
                            notes = "DELETED";

                        if (existingRecords.TryGetValue(description, out var existing))
                        {
                            // For updates, use existing dates as fallback instead of today
                            string actStart = "";
                            if (percentEntry > 0)
                            {
                                if (cwp.MinShipDate.HasValue)
                                    actStart = cwp.MinShipDate.Value.ToString("yyyy-MM-dd HH:mm:ss");
                                else
                                    actStart = !string.IsNullOrEmpty(existing.ActStart) ? existing.ActStart : today.ToString("yyyy-MM-dd HH:mm:ss");
                            }

                            string actFin = "";
                            if (percentEntry >= 100)
                            {
                                if (cwp.MaxShipDate.HasValue)
                                    actFin = cwp.MaxShipDate.Value.ToString("yyyy-MM-dd HH:mm:ss");
                                else
                                    actFin = !string.IsNullOrEmpty(existing.ActFin) ? existing.ActFin : today.ToString("yyyy-MM-dd HH:mm:ss");
                            }

                            // Check if anything actually changed
                            var existingNotes = existing.Notes;
                            if (cwp.AllDeleted && !existingNotes.Contains("DELETED"))
                                notes = string.IsNullOrEmpty(existingNotes) ? "DELETED" : $"{existingNotes} DELETED";
                            else
                                notes = existingNotes;

                            var newQty = cwp.AllDeleted ? existing.Quantity : cwp.TotalQuantity;

                            bool changed = Math.Abs(existing.PercentEntry - percentEntry) > 0.0001
                                || Math.Abs(existing.Quantity - newQty) > 0.0001
                                || existing.ActStart != actStart
                                || existing.ActFin != actFin
                                || notes != existingNotes;

                            if (changed)
                            {
                                UpdateActivity(connection, transaction, existing, newQty, percentEntry, actStart, actFin, notes);
                                updated++;
                            }
                            else
                            {
                                unchanged++;
                            }
                        }
                        else
                        {
                            // For new records, use today as fallback
                            string actStart = "";
                            if (percentEntry > 0)
                                actStart = (cwp.MinShipDate ?? today).ToString("yyyy-MM-dd HH:mm:ss");

                            string actFin = "";
                            if (percentEntry >= 100)
                                actFin = (cwp.MaxShipDate ?? today).ToString("yyyy-MM-dd HH:mm:ss");

                            var uniqueId = $"i{timestamp}{sequence}{userSuffix}";
                            sequence++;
                            InsertActivity(connection, transaction, cwp, description, percentEntry, uniqueId, currentUser, actStart, actFin, notes);
                            created++;
                        }
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            });

            // Show results
            var message = $"PTP Import Complete\n\nCreated: {created}\nUpdated: {updated}\nUnchanged: {unchanged}";

            if (deletedCwps.Count > 0)
            {
                message += $"\n\nThe following CWPs had ALL items marked as Deleted.\n" +
                           $"Records set to 100% with DELETED in Notes:\n\n" +
                           string.Join("\n", deletedCwps);
            }

            _host.ShowInfo(message, "PTP TFS MECH Updater");
            _host.LogInfo($"PTP import: {created} created, {updated} updated, {unchanged} unchanged, {deletedCwps.Count} all-deleted", "PtpImporter.RunAsync");
            // Refresh Progress view to show new/updated records
            await _host.RefreshProgressViewAsync();
        }

        // Collapse newlines and extra whitespace into single spaces
        private static string NormalizeWhitespace(string value)
        {
            return Regex.Replace(value, @"\s+", " ").Trim();
        }

        // Try to parse a date from a cell (handles DateTime and string formats)
        private static DateTime? GetDateValue(IXLCell cell)
        {
            if (cell.IsEmpty()) return null;
            if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
            var text = cell.GetString().Trim();
            if (DateTime.TryParse(text, out var date)) return date;
            return null;
        }

        // Parse the PTP Excel report and aggregate quantities per CWP
        private List<CwpData>? ParseReport(string filePath)
        {
            try
            {
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheets.First();

                // Validate headers on row 1 (vendor headers may have embedded newlines)
                var headerRow = worksheet.Row(1);
                var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 1; col <= (worksheet.LastColumnUsed()?.ColumnNumber() ?? 0); col++)
                {
                    var value = NormalizeWhitespace(headerRow.Cell(col).GetString());
                    if (!string.IsNullOrEmpty(value))
                        headers[value] = col;
                }

                foreach (var required in RequiredHeaders)
                {
                    if (!headers.ContainsKey(required))
                    {
                        _host.ShowError(
                            $"File is not formatted properly. Column headers must be on the first row.\n\n" +
                            $"Missing required column: \"{required}\"",
                            "Invalid File Format");
                        return null;
                    }
                }

                int cwpCol = headers["CWP"];
                int qtyCol = headers["Quantity"];
                int statusCol = headers["Status"];
                int shippedCol = headers["Shipped QTY"];

                // Actual Ship Date is optional — used for ActStart/ActFin if available
                int shipDateCol = headers.TryGetValue("Actual Ship Date", out var col2) ? col2 : -1;

                // Aggregate per CWP
                var cwpMap = new Dictionary<string, (double totalQty, double shippedQty, int totalRows, int deletedRows, DateTime? minDate, DateTime? maxDate)>();

                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                for (int row = 2; row <= lastRow; row++)
                {
                    var cwp = worksheet.Cell(row, cwpCol).GetString().Trim();
                    if (string.IsNullOrEmpty(cwp)) continue;

                    var status = worksheet.Cell(row, statusCol).GetString().Trim();
                    bool isDeleted = status.Equals("Deleted", StringComparison.OrdinalIgnoreCase);

                    double qty = 0;
                    double shipped = 0;
                    DateTime? shipDate = null;

                    if (!isDeleted)
                    {
                        qty = GetNumericValue(worksheet.Cell(row, qtyCol));
                        shipped = GetNumericValue(worksheet.Cell(row, shippedCol));
                        if (shipDateCol > 0)
                            shipDate = GetDateValue(worksheet.Cell(row, shipDateCol));
                    }

                    if (cwpMap.TryGetValue(cwp, out var existing))
                    {
                        // Track min/max ship dates across rows
                        var minDate = existing.minDate;
                        var maxDate = existing.maxDate;
                        if (shipDate.HasValue)
                        {
                            minDate = minDate.HasValue ? (shipDate < minDate ? shipDate : minDate) : shipDate;
                            maxDate = maxDate.HasValue ? (shipDate > maxDate ? shipDate : maxDate) : shipDate;
                        }

                        cwpMap[cwp] = (
                            existing.totalQty + qty,
                            existing.shippedQty + shipped,
                            existing.totalRows + 1,
                            existing.deletedRows + (isDeleted ? 1 : 0),
                            minDate,
                            maxDate
                        );
                    }
                    else
                    {
                        cwpMap[cwp] = (qty, shipped, 1, isDeleted ? 1 : 0, shipDate, shipDate);
                    }
                }

                return cwpMap.Select(kvp => new CwpData
                {
                    Cwp = kvp.Key,
                    TotalQuantity = kvp.Value.totalQty,
                    TotalShippedQty = kvp.Value.shippedQty,
                    AllDeleted = kvp.Value.totalRows > 0 && kvp.Value.deletedRows == kvp.Value.totalRows,
                    MinShipDate = kvp.Value.minDate,
                    MaxShipDate = kvp.Value.maxDate
                }).ToList();
            }
            catch (Exception ex)
            {
                _host.LogError(ex, "PtpImporter.ParseReport");
                _host.ShowError($"Failed to read PTP report:\n\n{ex.Message}", "File Error");
                return null;
            }
        }

        private static double GetNumericValue(IXLCell cell)
        {
            if (cell.IsEmpty()) return 0;
            if (cell.DataType == XLDataType.Number) return cell.GetDouble();
            if (double.TryParse(cell.GetString().Trim(), out var result)) return result;
            return 0;
        }

        // Find existing activities matching the description pattern
        private async Task<Dictionary<string, ExistingRecord>> FindExistingRecordsAsync()
        {
            var result = new Dictionary<string, ExistingRecord>(StringComparer.OrdinalIgnoreCase);

            await Task.Run(() =>
            {
                using var connection = DatabaseSetup.GetConnection();
                connection.Open();

                var cmd = connection.CreateCommand();
                cmd.CommandText = "SELECT UniqueID, Description, AssignedTo, Notes, Quantity, PercentEntry, ActStart, ActFin FROM Activities WHERE Description LIKE 'FABRICATION - 7.SHP %'";

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var desc = reader["Description"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(desc))
                    {
                        result[desc] = new ExistingRecord
                        {
                            UniqueID = reader["UniqueID"]?.ToString() ?? "",
                            AssignedTo = reader["AssignedTo"]?.ToString() ?? "",
                            Notes = reader["Notes"]?.ToString() ?? "",
                            Quantity = reader["Quantity"] != DBNull.Value ? Convert.ToDouble(reader["Quantity"]) : 0,
                            PercentEntry = reader["PercentEntry"] != DBNull.Value ? Convert.ToDouble(reader["PercentEntry"]) : 0,
                            ActStart = reader["ActStart"]?.ToString() ?? "",
                            ActFin = reader["ActFin"]?.ToString() ?? ""
                        };
                    }
                }
            });

            return result;
        }

        // Update an existing activity
        private void UpdateActivity(SqliteConnection connection, SqliteTransaction transaction,
            ExistingRecord existing, double quantity, double percentEntry, string actStart, string actFin, string notes)
        {
            var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = @"
                UPDATE Activities SET
                    Quantity = @Quantity,
                    PercentEntry = @PercentEntry,
                    Notes = @Notes,
                    ActStart = @ActStart,
                    ActFin = @ActFin,
                    LocalDirty = 1,
                    UpdatedBy = @UpdatedBy,
                    UpdatedUtcDate = @UpdatedUtcDate
                WHERE UniqueID = @UniqueID";

            cmd.Parameters.AddWithValue("@Quantity", quantity);
            cmd.Parameters.AddWithValue("@PercentEntry", percentEntry);
            cmd.Parameters.AddWithValue("@Notes", notes);
            cmd.Parameters.AddWithValue("@ActStart", actStart);
            cmd.Parameters.AddWithValue("@ActFin", actFin);
            cmd.Parameters.AddWithValue("@UpdatedBy", _host.CurrentUsername);
            cmd.Parameters.AddWithValue("@UpdatedUtcDate", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@UniqueID", existing.UniqueID);

            cmd.ExecuteNonQuery();
        }

        // Insert a new activity
        private void InsertActivity(SqliteConnection connection, SqliteTransaction transaction,
            CwpData cwp, string description, double percentEntry, string uniqueId, string currentUser,
            string actStart, string actFin, string notes)
        {
            var budgetQty = cwp.AllDeleted ? 0.001 : cwp.TotalQuantity;
            var now = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss");

            var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;
            cmd.CommandText = @"
                INSERT INTO Activities (
                    UniqueID, ActivityID, Area, AssignedTo, Aux1, Aux2, Aux3,
                    BaseUnit, BudgetHoursGroup, BudgetHoursROC, BudgetMHs, ChgOrdNO, ClientBudget,
                    ClientCustom3, ClientEquivQty, CompType, CreatedBy, DateTrigger, Description,
                    DwgNO, EarnQtyEntry, EarnedMHsRoc, EqmtNO, EquivQTY, EquivUOM, Estimator,
                    HexNO, HtTrace, InsulType, LineNumber, LocalDirty, MtrlSpec, Notes, PaintCode,
                    PercentEntry, PhaseCategory, PhaseCode, PipeGrade, PipeSize1, PipeSize2,
                    PrevEarnMHs, PrevEarnQTY, ProjectID, Quantity, RevNO, RFINO,
                    ROCBudgetQTY, ROCID, ROCPercent, ROCStep, SchedActNO, ActFin, ActStart,
                    SecondActno, SecondDwgNO, Service, ShopField, ShtNO, SubArea, PjtSystem, PjtSystemNo,
                    TagNO, UDF1, UDF2, UDF3, UDF4, UDF5, UDF6, UDF7, UDF8, UDF9,
                    UDF10, UDF11, UDF12, UDF13, UDF14, UDF15, UDF16, UDF17, RespParty, UDF20,
                    UpdatedBy, UpdatedUtcDate, UOM, WorkPackage, XRay, SyncVersion, AzureUploadUtcDate,
                    PlanStart, PlanFin
                ) VALUES (
                    @UniqueID, 0, @Area, @AssignedTo, '', '', '',
                    0, 0, 0, @BudgetMHs, '', 0.001,
                    0, 0, @CompType, @CreatedBy, 0, @Description,
                    '', 0, 0, '', 0, '', '',
                    0, '', '', '', 1, '', @Notes, '',
                    @PercentEntry, @PhaseCategory, @PhaseCode, '', 0, 0,
                    0, 0, @ProjectID, @Quantity, '', '',
                    0, 0, 0, @ROCStep, @SchedActNO, @ActFin, @ActStart,
                    '', '', '', @ShopField, '', '', '', '',
                    '', @UDF1, @UDF2, @UDF3, '', '', '', 0, '', '',
                    '', '', '', '', '', '', '', '', @RespParty, '',
                    @UpdatedBy, @UpdatedUtcDate, '', @WorkPackage, 0, 0, '',
                    '', ''
                )";

            cmd.Parameters.AddWithValue("@UniqueID", uniqueId);
            cmd.Parameters.AddWithValue("@Area", Area);
            cmd.Parameters.AddWithValue("@AssignedTo", currentUser);
            cmd.Parameters.AddWithValue("@BudgetMHs", BudgetMHs);
            cmd.Parameters.AddWithValue("@CompType", CompType);
            cmd.Parameters.AddWithValue("@CreatedBy", currentUser);
            cmd.Parameters.AddWithValue("@Description", description);
            cmd.Parameters.AddWithValue("@Notes", notes);
            cmd.Parameters.AddWithValue("@PercentEntry", percentEntry);
            cmd.Parameters.AddWithValue("@PhaseCategory", PhaseCategory);
            cmd.Parameters.AddWithValue("@PhaseCode", PhaseCode);
            cmd.Parameters.AddWithValue("@ProjectID", ProjectID);
            cmd.Parameters.AddWithValue("@Quantity", budgetQty);
            cmd.Parameters.AddWithValue("@ROCStep", ROCStep);
            cmd.Parameters.AddWithValue("@SchedActNO", SchedActNO);
            cmd.Parameters.AddWithValue("@ShopField", ShopField);
            cmd.Parameters.AddWithValue("@UDF1", UDF1);
            cmd.Parameters.AddWithValue("@UDF2", cwp.Cwp);
            cmd.Parameters.AddWithValue("@UDF3", UDF3);
            cmd.Parameters.AddWithValue("@RespParty", RespParty);
            cmd.Parameters.AddWithValue("@WorkPackage", WorkPackage);
            cmd.Parameters.AddWithValue("@UpdatedBy", currentUser);
            cmd.Parameters.AddWithValue("@UpdatedUtcDate", now);
            cmd.Parameters.AddWithValue("@ActStart", actStart);
            cmd.Parameters.AddWithValue("@ActFin", actFin);

            cmd.ExecuteNonQuery();
        }
    }
}
