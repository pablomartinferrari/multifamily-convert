using System;
using System.Collections.Generic;
using System.Linq;
using ExcelProcessor.Models;

namespace ExcelProcessor.Services
{
    public class XrfProcessingService
    {
        private const int AveragingThreshold = 40;
        private const double PositivityThresholdPercent = 2.5;

        public ProcessingResults ProcessShots(List<XrfShot> allShots)
        {
            var results = new ProcessingResults();

            // 1. Filter out calibration shots and normalize component names
            var validShots = allShots
                .Where(s => !s.IsCalibration && !string.IsNullOrWhiteSpace(s.Component))
                .ToList();

            // 2. Group by normalized component name
            var groups = validShots.GroupBy(s => s.Component.ToLower().Trim());

            foreach (var group in groups)
            {
                var componentName = group.Key;
                var shots = group.ToList();
                var totalCount = shots.Count;
                var positiveCount = shots.Count(s => s.IsPositive);
                var negativeCount = totalCount - positiveCount;

                var posPercent = (double)positiveCount / totalCount * 100;
                var negPercent = (double)negativeCount / totalCount * 100;

                // Rule: Averaged Results (>= 40 shots)
                if (totalCount >= AveragingThreshold)
                {
                    results.AveragedResults.Add(new ComponentSummary
                    {
                        Component = shots.First().Component,
                        Count = totalCount,
                        PositivePercentage = Math.Round(posPercent, 2),
                        NegativePercentage = Math.Round(negPercent, 2),
                        // 2.5% rule for averaging
                        LeadContent = posPercent > PositivityThresholdPercent ? "Positive" : "Negative"
                    });
                }
                // Rule: Individually Tested ( < 40 shots)
                else
                {
                    // Uniform Results: All positive or all negative
                    if (positiveCount == 0 || positiveCount == totalCount)
                    {
                        results.UniformResults.Add(new ComponentSummary
                        {
                            Component = shots.First().Component,
                            Count = totalCount,
                            PositivePercentage = Math.Round(posPercent, 2),
                            NegativePercentage = Math.Round(negPercent, 2),
                            LeadContent = positiveCount > 0 ? "Positive" : "Negative"
                        });
                    }
                    // Conflicting Results: Mix of pos/neg and < 40 shots
                    else
                    {
                        results.ConflictingResults.AddRange(shots);
                    }
                }
            }

            // Sort results
            results.AveragedResults = results.AveragedResults.OrderBy(r => r.Component).ToList();
            results.UniformResults = results.UniformResults.OrderBy(r => r.Component).ToList();
            results.ConflictingResults = results.ConflictingResults.OrderBy(s => s.Reading).ToList();

            return results;
        }
    }
}



