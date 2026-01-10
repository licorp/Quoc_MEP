using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Quoc_MEP.Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ScheduleManager
{
    /// <summary>
    /// Safely reads schedule data using async pattern without blocking Revit UI
    /// Uses RevitAsyncHelper (similar to Revit.Async) for thread-safe operations
    /// </summary>
    public class AsyncScheduleReader
    {
        private readonly UIApplication _uiApp;

        public class ScheduleData
        {
            public string ScheduleName { get; set; }
            public List<string> ColumnHeaders { get; set; } = new List<string>();
            public List<ScheduleRow> Rows { get; set; } = new List<ScheduleRow>();
            public int TotalElements { get; set; }
        }

        public class ReadProgress
        {
            public int Current { get; set; }
            public int Total { get; set; }
            public string Status { get; set; }
            public double PercentComplete => Total > 0 ? (double)Current / Total * 100 : 0;
        }

        public AsyncScheduleReader(UIApplication uiApp)
        {
            _uiApp = uiApp ?? throw new ArgumentNullException(nameof(uiApp));
        }

        /// <summary>
        /// Read schedule data asynchronously without blocking Revit UI
        /// Phase 1: Read structure on main thread (Revit API)
        /// Phase 2: Process data on background thread
        /// </summary>
        public async Task<ScheduleData> ReadScheduleAsync(
            ViewSchedule schedule,
            IProgress<ReadProgress> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (schedule == null)
                throw new ArgumentNullException(nameof(schedule));

            // Report initial status
            progress?.Report(new ReadProgress
            {
                Status = "Reading schedule structure...",
                Current = 0,
                Total = 100
            });

            // PHASE 1: Read structure on main thread using RevitAsyncHelper
            var structure = await ReadScheduleStructureAsync(schedule, cancellationToken);

            progress?.Report(new ReadProgress
            {
                Status = $"Processing {structure.ElementIds.Count} elements...",
                Current = 10,
                Total = 100
            });

            // PHASE 2: Process data in batches on background thread
            var result = await Task.Run(() =>
            {
                var scheduleData = new ScheduleData
                {
                    ScheduleName = structure.ScheduleName,
                    ColumnHeaders = structure.ColumnHeaders,
                    TotalElements = structure.ElementIds.Count
                };

                // Process in batches to allow cancellation
                const int batchSize = 100;
                int processed = 0;

                for (int i = 0; i < structure.ElementIds.Count; i += batchSize)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var batch = structure.ElementIds.Skip(i).Take(batchSize).ToList();
                    
                    // Load element data on main thread
                    var batchRows = LoadElementBatchAsync(
                        batch,
                        structure.ColumnHeaders,
                        structure.FieldIds,
                        cancellationToken
                    ).GetAwaiter().GetResult(); // Synchronous wait in background task

                    scheduleData.Rows.AddRange(batchRows);
                    processed += batchRows.Count;

                    progress?.Report(new ReadProgress
                    {
                        Status = $"Loaded {processed} / {structure.ElementIds.Count} elements",
                        Current = 10 + (int)(processed * 90.0 / structure.ElementIds.Count),
                        Total = 100
                    });
                }

                return scheduleData;
            }, cancellationToken);

            progress?.Report(new ReadProgress
            {
                Status = "Complete",
                Current = 100,
                Total = 100
            });

            return result;
        }

        /// <summary>
        /// Phase 1: Read schedule structure (column headers, element IDs) on main thread
        /// MUST run on Revit main thread via RevitAsyncHelper
        /// </summary>
        private async Task<ScheduleStructure> ReadScheduleStructureAsync(
            ViewSchedule schedule,
            CancellationToken cancellationToken)
        {
            return await RevitAsyncHelper.RunAsync<ScheduleStructure>(uiApp =>
            {
                cancellationToken.ThrowIfCancellationRequested();

                var doc = uiApp.ActiveUIDocument.Document;
                var structure = new ScheduleStructure
                {
                    ScheduleName = schedule.Name
                };

                // Read column headers and field IDs
                var definition = schedule.Definition;
                var fieldOrder = definition.GetFieldOrder();

                foreach (var fieldId in fieldOrder)
                {
                    var field = definition.GetField(fieldId);
                    if (!field.IsHidden)
                    {
                        structure.ColumnHeaders.Add(field.GetName());
                        structure.FieldIds.Add(fieldId);
                    }
                }

                // Get all element IDs from schedule
                var collector = new FilteredElementCollector(doc, schedule.Id)
                    .WhereElementIsNotElementType();
                
                structure.ElementIds = collector.ToElementIds().ToList();

                return structure;
            });
        }

        /// <summary>
        /// Load batch of elements on main thread
        /// MUST run on Revit main thread via RevitAsyncHelper
        /// </summary>
        private async Task<List<ScheduleRow>> LoadElementBatchAsync(
            List<ElementId> elementIds,
            List<string> columnHeaders,
            List<ScheduleFieldId> fieldIds,
            CancellationToken cancellationToken)
        {
            return await RevitAsyncHelper.RunAsync<List<ScheduleRow>>(uiApp =>
            {
                cancellationToken.ThrowIfCancellationRequested();

                var doc = uiApp.ActiveUIDocument.Document;
                var rows = new List<ScheduleRow>();

                for (int i = 0; i < elementIds.Count; i++)
                {
                    try
                    {
                        var element = doc.GetElement(elementIds[i]);
                        if (element == null) continue;

                        var row = new ScheduleRow(null, i, columnHeaders.Count)
                        {
                            ElementId = elementIds[i]
                        };

                        // Read parameter values for each field
                        for (int colIndex = 0; colIndex < fieldIds.Count; colIndex++)
                        {
                            var field = doc.ActiveView is ViewSchedule vs
                                ? vs.Definition.GetField(fieldIds[colIndex])
                                : null;

                            if (field == null) continue;

                            var parameter = GetParameterFromField(element, field);
                            var value = parameter?.AsValueString() ?? string.Empty;
                            
                            row.SetValue(colIndex, value);
                        }

                        rows.Add(row);
                    }
                    catch (Exception ex)
                    {
                        // Skip problematic elements
                        System.Diagnostics.Debug.WriteLine($"Error reading element {elementIds[i]}: {ex.Message}");
                    }
                }

                return rows;
            });
        }

        /// <summary>
        /// Get parameter from element based on schedule field
        /// </summary>
        private Parameter GetParameterFromField(Element element, ScheduleField field)
        {
            var parameterId = field.ParameterId;
            
            if (parameterId == ElementId.InvalidElementId)
                return null;

            // Try to get parameter by ParameterId (cast to BuiltInParameter for Revit 2020 compatibility)
            return element.get_Parameter((BuiltInParameter)parameterId.IntegerValue);
        }

        /// <summary>
        /// Internal structure to hold schedule metadata
        /// </summary>
        private class ScheduleStructure
        {
            public string ScheduleName { get; set; }
            public List<string> ColumnHeaders { get; set; } = new List<string>();
            public List<ScheduleFieldId> FieldIds { get; set; } = new List<ScheduleFieldId>();
            public List<ElementId> ElementIds { get; set; } = new List<ElementId>();
        }
    }
}
