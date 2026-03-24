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

namespace ConstTfsMechUpdater
{
    // Data for a single spool from the CONST report
    internal class SpoolData
    {
        public string PieceMark { get; set; } = "";
        public string Contract { get; set; } = "";
        public string Spool { get; set; } = "";
        public string Isometric { get; set; } = "";
        public string RevNo { get; set; } = "";
        public string Area { get; set; } = "";
        public string Line { get; set; } = "";
        public string System { get; set; } = "";
        public string Insulation { get; set; } = "";
        public string Module { get; set; } = "";
        public string Class { get; set; } = "";
        public string PaintSystem { get; set; } = "";
        public string SpoolSize { get; set; } = "";
        public string Mtl { get; set; } = "";
        public double Weight { get; set; }
        public double PipeLength { get; set; }
        public double WldPercentAll { get; set; }
        public DateTime? RlsToFabDate { get; set; }
        public DateTime? FinalShipment { get; set; }
        public DateTime? EstimatedShipDate { get; set; }
    }

    // Existing activity data needed for comparison
    internal class ExistingRecord
    {
        public string UniqueID { get; set; } = "";
        public string SecondDwgNO { get; set; } = "";
        public string AssignedTo { get; set; } = "";
        public string Notes { get; set; } = "";
        public double Quantity { get; set; }
        public double PercentEntry { get; set; }
        public double Weight { get; set; }
        public string ActStart { get; set; } = "";
        public string ActFin { get; set; } = "";
        public string PlanFin { get; set; } = "";
    }

    internal class ConstImporter
    {
        private readonly IPluginHost _host;

        // Required columns in the CONST report
        private static readonly string[] RequiredHeaders = { "Piece Mark", "Pipe Length", "WLD % All" };

        // Hardcoded activity values
        private const string ROCStep = "4.SHP";
        private const string DescriptionPrefix = "Constellation Fabrication for PieceMark ";
        private const string DescriptionPattern = "Constellation Fabrication for PieceMark %";
        private const string CompType = "P";
        private const string PhaseCategory = "PIP";
        private const string PhaseCode = "xx.xxx.";
        private const string ProjectID = "25.005.";
        private const string RespParty = "SUMMIT - PM";
        private const string SchedActNO = "x";
        private const string UDF6 = "SPL";
        private const string Aux3 = "K209";
        private const string UOM = "LFP";
        private const double BudgetMHs = 0.001;

        public ConstImporter(IPluginHost host)
        {
            _host = host;
        }

        public async Task RunAsync(string filePath)
        {
            // Parse the Excel file
            var spoolDataList = ParseReport(filePath);
            if (spoolDataList == null) return;

            _host.LogInfo($"Parsed {spoolDataList.Count} spools from CONST report", "ConstImporter.RunAsync");

            // Look up existing activities by SecondDwgNO (Piece Mark)
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
                    $"Cannot update records. The following user(s) own existing CONST activities:\n\n" +
                    $"{string.Join(", ", foreignOwners)}\n\n" +
                    $"Only the original importer can update these records.",
                    "Ownership Conflict");
                return;
            }

            // Track Piece Marks in the new report for deletion detection
            var reportPieceMarks = new HashSet<string>(spoolDataList.Select(s => s.PieceMark), StringComparer.OrdinalIgnoreCase);

            // Process: create new, update existing, or mark deleted
            int created = 0;
            int updated = 0;
            int unchanged = 0;
            int deleted = 0;
            var dataWarnings = new List<string>();
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

                    // Process spools from the report
                    foreach (var spool in spoolDataList)
                    {
                        var description = $"{DescriptionPrefix}{spool.PieceMark} - Spool {spool.Spool}";

                        // Calculate percent entry: (WLD % All × 0.8) + (Final Shipment ? 20 : 0)
                        // Excel stores percentages as decimals (0-1), so convert to 0-100 range first
                        double wldPercent = spool.WldPercentAll <= 1 ? spool.WldPercentAll * 100 : spool.WldPercentAll;
                        double percentEntry = Math.Round(wldPercent * 0.8, 3);
                        bool hasShipped = spool.FinalShipment.HasValue;
                        if (hasShipped)
                            percentEntry += 20;

                        // Data warning: shipped but WLD < 100%
                        string notes = "";
                        if (hasShipped && wldPercent < 100)
                        {
                            notes = $"DATA WARNING: Shipped with WLD at {wldPercent}%";
                            dataWarnings.Add($"{spool.PieceMark}: WLD at {wldPercent}%");
                        }

                        // ActStart from RLS to Fab date
                        string actStart = "";
                        if (spool.RlsToFabDate.HasValue)
                            actStart = spool.RlsToFabDate.Value.ToString("yyyy-MM-dd HH:mm:ss");

                        // ActFin from Final Shipment
                        string actFin = "";
                        if (spool.FinalShipment.HasValue)
                            actFin = spool.FinalShipment.Value.ToString("yyyy-MM-dd HH:mm:ss");

                        // PlanFin from Estimated Ship Date
                        string planFin = "";
                        if (spool.EstimatedShipDate.HasValue)
                            planFin = spool.EstimatedShipDate.Value.ToString("yyyy-MM-dd HH:mm:ss");

                        if (existingRecords.TryGetValue(spool.PieceMark, out var existing))
                        {
                            // Preserve existing notes unless we have a new warning
                            var existingNotes = existing.Notes;
                            if (string.IsNullOrEmpty(notes) && !string.IsNullOrEmpty(existingNotes))
                            {
                                // Keep existing notes unless they were a previous data warning we should clear
                                if (!existingNotes.StartsWith("DATA WARNING:") || (hasShipped && spool.WldPercentAll < 100))
                                    notes = existingNotes;
                            }

                            // Check if anything actually changed
                            bool changed = Math.Abs(existing.PercentEntry - percentEntry) > 0.0001
                                || Math.Abs(existing.Quantity - spool.PipeLength) > 0.0001
                                || Math.Abs(existing.Weight - spool.Weight) > 0.0001
                                || existing.ActStart != actStart
                                || existing.ActFin != actFin
                                || existing.PlanFin != planFin
                                || notes != existingNotes;

                            if (changed)
                            {
                                UpdateActivity(connection, transaction, existing, spool, percentEntry, actStart, actFin, planFin, notes);
                                updated++;
                            }
                            else
                            {
                                unchanged++;
                            }
                        }
                        else
                        {
                            var uniqueId = $"i{timestamp}{sequence}{userSuffix}";
                            sequence++;
                            InsertActivity(connection, transaction, spool, description, percentEntry, uniqueId, currentUser, actStart, actFin, planFin, notes);
                            created++;
                        }
                    }

                    // Mark missing spools as deleted
                    foreach (var existing in existingRecords.Values)
                    {
                        if (!reportPieceMarks.Contains(existing.SecondDwgNO))
                        {
                            // Skip if already marked deleted
                            if (existing.Notes.Contains("DELETED"))
                            {
                                unchanged++;
                                continue;
                            }

                            var newNotes = string.IsNullOrEmpty(existing.Notes) ? "DELETED" : $"{existing.Notes} DELETED";
                            var actFin = string.IsNullOrEmpty(existing.ActFin) ? today.ToString("yyyy-MM-dd HH:mm:ss") : existing.ActFin;

                            MarkAsDeleted(connection, transaction, existing, actFin, newNotes);
                            deleted++;
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
            var message = $"CONST Import Complete\n\nCreated: {created}\nUpdated: {updated}\nUnchanged: {unchanged}\nDeleted: {deleted}";

            if (dataWarnings.Count > 0)
            {
                message += $"\n\nData Warnings ({dataWarnings.Count} spools shipped with WLD < 100%):\n" +
                           string.Join("\n", dataWarnings.Take(10));
                if (dataWarnings.Count > 10)
                    message += $"\n... and {dataWarnings.Count - 10} more";
            }

            _host.ShowInfo(message, "CONST TFS MECH Updater");
            _host.LogInfo($"CONST import: {created} created, {updated} updated, {unchanged} unchanged, {deleted} deleted, {dataWarnings.Count} warnings", "ConstImporter.RunAsync");

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

        // Parse the CONST Excel report - one record per spool
        private List<SpoolData>? ParseReport(string filePath)
        {
            try
            {
                using var workbook = new XLWorkbook(filePath);

                // Validate single worksheet
                if (workbook.Worksheets.Count > 1)
                {
                    _host.ShowError(
                        "The report contains multiple tabs. Please delete all tabs except 'Detailed Spool Report' and try again.",
                        "Invalid File Format");
                    return null;
                }

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

                // Get column indices (required)
                int pieceMarkCol = headers["Piece Mark"];
                int pipeLengthCol = headers["Pipe Length"];
                int wldPercentCol = headers["WLD % All"];

                // Get column indices (optional)
                int contractCol = headers.TryGetValue("Contract", out var c1) ? c1 : -1;
                int spoolCol = headers.TryGetValue("Spool", out var c2) ? c2 : -1;
                int isometricCol = headers.TryGetValue("Isometric", out var c3) ? c3 : -1;
                int revNoCol = headers.TryGetValue("Rev #", out var c4) ? c4 : -1;
                int areaCol = headers.TryGetValue("Area", out var c5) ? c5 : -1;
                int lineCol = headers.TryGetValue("Line", out var c6) ? c6 : -1;
                int systemCol = headers.TryGetValue("System", out var c7) ? c7 : -1;
                int insulationCol = headers.TryGetValue("Insulation", out var c8) ? c8 : -1;
                int moduleCol = headers.TryGetValue("Module", out var c9) ? c9 : -1;
                int classCol = headers.TryGetValue("Class", out var c10) ? c10 : -1;
                int paintSystemCol = headers.TryGetValue("Paint System", out var c11) ? c11 : -1;
                int spoolSizeCol = headers.TryGetValue("Spool Size", out var c12) ? c12 : -1;
                int mtlCol = headers.TryGetValue("MTL", out var c13) ? c13 : -1;
                int weightCol = headers.TryGetValue("Weight", out var c14) ? c14 : -1;
                int rlsToFabCol = headers.TryGetValue("RLS to Fab date", out var c15) ? c15 : -1;
                int finalShipCol = headers.TryGetValue("Final Shipment", out var c16) ? c16 : -1;
                int estShipCol = headers.TryGetValue("Estimated Ship Date", out var c17) ? c17 : -1;

                var spoolList = new List<SpoolData>();
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                for (int row = 2; row <= lastRow; row++)
                {
                    var pieceMark = worksheet.Cell(row, pieceMarkCol).GetString().Trim();
                    if (string.IsNullOrEmpty(pieceMark)) continue;

                    var spool = new SpoolData
                    {
                        PieceMark = pieceMark,
                        PipeLength = GetNumericValue(worksheet.Cell(row, pipeLengthCol)),
                        WldPercentAll = GetNumericValue(worksheet.Cell(row, wldPercentCol)),
                        Contract = contractCol > 0 ? worksheet.Cell(row, contractCol).GetString().Trim() : "",
                        Spool = spoolCol > 0 ? worksheet.Cell(row, spoolCol).GetString().Trim() : "",
                        Isometric = isometricCol > 0 ? worksheet.Cell(row, isometricCol).GetString().Trim() : "",
                        RevNo = revNoCol > 0 ? worksheet.Cell(row, revNoCol).GetString().Trim() : "",
                        Area = areaCol > 0 ? worksheet.Cell(row, areaCol).GetString().Trim() : "",
                        Line = lineCol > 0 ? worksheet.Cell(row, lineCol).GetString().Trim() : "",
                        System = systemCol > 0 ? worksheet.Cell(row, systemCol).GetString().Trim() : "",
                        Insulation = insulationCol > 0 ? worksheet.Cell(row, insulationCol).GetString().Trim() : "",
                        Module = moduleCol > 0 ? worksheet.Cell(row, moduleCol).GetString().Trim() : "",
                        Class = classCol > 0 ? worksheet.Cell(row, classCol).GetString().Trim() : "",
                        PaintSystem = paintSystemCol > 0 ? worksheet.Cell(row, paintSystemCol).GetString().Trim() : "",
                        SpoolSize = spoolSizeCol > 0 ? worksheet.Cell(row, spoolSizeCol).GetString().Trim() : "",
                        Mtl = mtlCol > 0 ? worksheet.Cell(row, mtlCol).GetString().Trim() : "",
                        Weight = weightCol > 0 ? GetNumericValue(worksheet.Cell(row, weightCol)) : 0,
                        RlsToFabDate = rlsToFabCol > 0 ? GetDateValue(worksheet.Cell(row, rlsToFabCol)) : null,
                        FinalShipment = finalShipCol > 0 ? GetDateValue(worksheet.Cell(row, finalShipCol)) : null,
                        EstimatedShipDate = estShipCol > 0 ? GetDateValue(worksheet.Cell(row, estShipCol)) : null
                    };

                    spoolList.Add(spool);
                }

                return spoolList;
            }
            catch (Exception ex)
            {
                _host.LogError(ex, "ConstImporter.ParseReport");
                _host.ShowError($"Failed to read CONST report:\n\n{ex.Message}", "File Error");
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

        // Find existing activities matching the CONST description pattern, keyed by SecondDwgNO (Piece Mark)
        private async Task<Dictionary<string, ExistingRecord>> FindExistingRecordsAsync()
        {
            var result = new Dictionary<string, ExistingRecord>(StringComparer.OrdinalIgnoreCase);

            await Task.Run(() =>
            {
                using var connection = DatabaseSetup.GetConnection();
                connection.Open();

                var cmd = connection.CreateCommand();
                cmd.CommandText = $"SELECT UniqueID, SecondDwgNO, AssignedTo, Notes, Quantity, PercentEntry, UDF7, ActStart, ActFin, PlanFin FROM Activities WHERE Description LIKE '{DescriptionPattern}'";

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var secondDwgNo = reader["SecondDwgNO"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(secondDwgNo))
                    {
                        result[secondDwgNo] = new ExistingRecord
                        {
                            UniqueID = reader["UniqueID"]?.ToString() ?? "",
                            SecondDwgNO = secondDwgNo,
                            AssignedTo = reader["AssignedTo"]?.ToString() ?? "",
                            Notes = reader["Notes"]?.ToString() ?? "",
                            Quantity = reader["Quantity"] != DBNull.Value ? Convert.ToDouble(reader["Quantity"]) : 0,
                            PercentEntry = reader["PercentEntry"] != DBNull.Value ? Convert.ToDouble(reader["PercentEntry"]) : 0,
                            Weight = reader["UDF7"] != DBNull.Value ? Convert.ToDouble(reader["UDF7"]) : 0,
                            ActStart = reader["ActStart"]?.ToString() ?? "",
                            ActFin = reader["ActFin"]?.ToString() ?? "",
                            PlanFin = reader["PlanFin"]?.ToString() ?? ""
                        };
                    }
                }
            });

            return result;
        }

        // Update an existing activity with new spool data (only progress fields + Quantity/Weight)
        private void UpdateActivity(SqliteConnection connection, SqliteTransaction transaction,
            ExistingRecord existing, SpoolData spool, double percentEntry, string actStart, string actFin, string planFin, string notes)
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
                    PlanFin = @PlanFin,
                    UDF7 = @UDF7,
                    LocalDirty = 1,
                    UpdatedBy = @UpdatedBy,
                    UpdatedUtcDate = @UpdatedUtcDate
                WHERE UniqueID = @UniqueID";

            cmd.Parameters.AddWithValue("@Quantity", spool.PipeLength);
            cmd.Parameters.AddWithValue("@PercentEntry", percentEntry);
            cmd.Parameters.AddWithValue("@Notes", notes);
            cmd.Parameters.AddWithValue("@ActStart", actStart);
            cmd.Parameters.AddWithValue("@ActFin", actFin);
            cmd.Parameters.AddWithValue("@PlanFin", planFin);
            cmd.Parameters.AddWithValue("@UDF7", spool.Weight);
            cmd.Parameters.AddWithValue("@UpdatedBy", _host.CurrentUsername);
            cmd.Parameters.AddWithValue("@UpdatedUtcDate", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@UniqueID", existing.UniqueID);

            cmd.ExecuteNonQuery();
        }

        // Mark an existing activity as deleted (missing from report)
        private void MarkAsDeleted(SqliteConnection connection, SqliteTransaction transaction,
            ExistingRecord existing, string actFin, string notes)
        {
            var cmd = connection.CreateCommand();
            cmd.Transaction = transaction;

            cmd.CommandText = @"
                UPDATE Activities SET
                    PercentEntry = 100,
                    Notes = @Notes,
                    ActFin = @ActFin,
                    LocalDirty = 1,
                    UpdatedBy = @UpdatedBy,
                    UpdatedUtcDate = @UpdatedUtcDate
                WHERE UniqueID = @UniqueID";

            cmd.Parameters.AddWithValue("@Notes", notes);
            cmd.Parameters.AddWithValue("@ActFin", actFin);
            cmd.Parameters.AddWithValue("@UpdatedBy", _host.CurrentUsername);
            cmd.Parameters.AddWithValue("@UpdatedUtcDate", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"));
            cmd.Parameters.AddWithValue("@UniqueID", existing.UniqueID);

            cmd.ExecuteNonQuery();
        }

        // Insert a new activity for a spool
        private void InsertActivity(SqliteConnection connection, SqliteTransaction transaction,
            SpoolData spool, string description, double percentEntry, string uniqueId, string currentUser,
            string actStart, string actFin, string planFin, string notes)
        {
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
                    @UniqueID, 0, '', @AssignedTo, '', '', @Aux3,
                    0, 0, 0, @BudgetMHs, @ChgOrdNO, 0.001,
                    0, 0, @CompType, @CreatedBy, 0, @Description,
                    @DwgNO, 0, 0, '', 0, '', '',
                    0, '', @InsulType, @LineNumber, 1, @MtrlSpec, @Notes, @PaintCode,
                    @PercentEntry, @PhaseCategory, @PhaseCode, @PipeGrade, @PipeSize1, 0,
                    0, 0, @ProjectID, @Quantity, @RevNO, '',
                    0, 0, 0, @ROCStep, @SchedActNO, @ActFin, @ActStart,
                    '', @SecondDwgNO, '', '', '', @SubArea, @PjtSystem, '',
                    '', '', @UDF2, '', '', @UDF5, @UDF6, @UDF7, '', '',
                    '', '', '', '', '', '', '', '', @RespParty, '',
                    @UpdatedBy, @UpdatedUtcDate, @UOM, @WorkPackage, 0, 0, '',
                    '', @PlanFin
                )";

            cmd.Parameters.AddWithValue("@UniqueID", uniqueId);
            cmd.Parameters.AddWithValue("@AssignedTo", currentUser);
            cmd.Parameters.AddWithValue("@Aux3", Aux3);
            cmd.Parameters.AddWithValue("@BudgetMHs", BudgetMHs);
            cmd.Parameters.AddWithValue("@ChgOrdNO", spool.Contract);
            cmd.Parameters.AddWithValue("@CompType", CompType);
            cmd.Parameters.AddWithValue("@CreatedBy", currentUser);
            cmd.Parameters.AddWithValue("@Description", description);
            cmd.Parameters.AddWithValue("@DwgNO", spool.Isometric);
            cmd.Parameters.AddWithValue("@InsulType", spool.Insulation);
            cmd.Parameters.AddWithValue("@LineNumber", spool.Line);
            cmd.Parameters.AddWithValue("@MtrlSpec", spool.Class);
            cmd.Parameters.AddWithValue("@Notes", notes);
            cmd.Parameters.AddWithValue("@PaintCode", spool.PaintSystem);
            cmd.Parameters.AddWithValue("@PercentEntry", percentEntry);
            cmd.Parameters.AddWithValue("@PhaseCategory", PhaseCategory);
            cmd.Parameters.AddWithValue("@PhaseCode", PhaseCode);
            cmd.Parameters.AddWithValue("@PipeGrade", spool.Mtl);
            cmd.Parameters.AddWithValue("@PipeSize1", spool.SpoolSize);
            cmd.Parameters.AddWithValue("@ProjectID", ProjectID);
            cmd.Parameters.AddWithValue("@Quantity", spool.PipeLength);
            cmd.Parameters.AddWithValue("@RevNO", spool.RevNo);
            cmd.Parameters.AddWithValue("@ROCStep", ROCStep);
            cmd.Parameters.AddWithValue("@SchedActNO", SchedActNO);
            cmd.Parameters.AddWithValue("@SecondDwgNO", spool.PieceMark);
            cmd.Parameters.AddWithValue("@SubArea", spool.Area);
            cmd.Parameters.AddWithValue("@PjtSystem", spool.System);
            cmd.Parameters.AddWithValue("@UDF2", spool.Module);
            cmd.Parameters.AddWithValue("@UDF5", spool.Spool);
            cmd.Parameters.AddWithValue("@UDF6", UDF6);
            cmd.Parameters.AddWithValue("@UDF7", spool.Weight);
            cmd.Parameters.AddWithValue("@RespParty", RespParty);
            cmd.Parameters.AddWithValue("@UOM", UOM);
            cmd.Parameters.AddWithValue("@WorkPackage", spool.Module);
            cmd.Parameters.AddWithValue("@UpdatedBy", currentUser);
            cmd.Parameters.AddWithValue("@UpdatedUtcDate", now);
            cmd.Parameters.AddWithValue("@ActStart", actStart);
            cmd.Parameters.AddWithValue("@ActFin", actFin);
            cmd.Parameters.AddWithValue("@PlanFin", planFin);

            cmd.ExecuteNonQuery();
        }
    }
}
