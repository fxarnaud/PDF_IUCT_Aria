using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using OxyPlot;
using OxyPlot.Axes;
using System.Windows;
//using CommunityToolkit.Mvvm.ComponentModel;
//using CommunityToolkit.Mvvm.Input;

namespace PDF_IUCT
{
    public partial class MainViewModel
    {
        private readonly PlanSetup _plan;
        public ScriptContext _ctx { get; private set; }
        public PlotModel PlotModel { get; private set; }
        public IEnumerable<StructureStatistics> Structures { get; private set; }


        public MainViewModel(ScriptContext ctx)
        {
            _ctx = ctx;
            _plan = ctx.PlanSetup;
            Structures = GetPlanStructuresAndComputeStatistics(ctx.PlanSetup);
            PlotModel = CreatePlotModel();
        }

        private List<StructureStatistics> GetPlanStructuresAndComputeStatistics(PlanSetup plan)
        {
            List<StructureStatistics> sstatistics = new List<StructureStatistics>();

            foreach (var structure in plan.StructureSet.Structures)
            {
                if ((structure.DicomType.ToUpper() == "PTV" || structure.DicomType.ToUpper() == "ORGAN" || structure.DicomType.ToUpper() == "CTV" || structure.DicomType.ToUpper() == "GTV") && !structure.IsEmpty)
                {
                    if (!structure.Id.ToUpper().Contains("OVER") && !structure.Id.ToUpper().Contains("-") && !structure.Id.ToUpper().Contains("GRAIN") && !structure.Id.ToUpper().Contains("PROTH")
                        && !structure.Id.ToUpper().Contains("BLOC") && !structure.Id.ToUpper().Contains("PAC") && !structure.Id.ToUpper().Contains("ARTEF") && !structure.Id.ToUpper().Contains("TOTAL")
                        && !structure.Id.ToUpper().Contains("ISOLE") && !structure.Id.ToUpper().Contains("PLANNING"))
                    {
                        StructureStatistics stat = new StructureStatistics();

                        stat.structure = structure;
                        stat.structure_id = structure.Id;
                        stat.volume = structure.Volume;
                        //Get structure statistics to fill a table                  
                        stat.d1cc = plan.GetDoseAtVolume(structure, 1, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
                        DoseValuePresentation dvp = DoseValuePresentation.Absolute;
                        DVHData dvh = plan.GetDVHCumulativeData(structure, dvp, VolumePresentation.Relative, 0.01);
                        stat.maxdose = dvh.MaxDose.Dose;
                        stat.meandose = dvh.MeanDose.Dose;
                        stat.isChecked = true;
                        sstatistics.Add(stat);

                    }
                }
            }

            return sstatistics != null
            ? sstatistics
            : null;
        }

        private PlotModel CreatePlotModel()
        {
            var plotModel = new PlotModel();
            AddAxes(plotModel);
            SetupLegend(plotModel);
            return plotModel;
        }
        private void SetupLegend(PlotModel pm)
        {

            pm.LegendBorder = OxyColors.Black;
            pm.LegendBackground = OxyColor.FromAColor(32, OxyColors.Black);
            pm.LegendPosition = LegendPosition.BottomCenter;
            pm.LegendOrientation = LegendOrientation.Horizontal;
            pm.LegendPlacement = LegendPlacement.Outside;
        }
        private static void AddAxes(PlotModel plotModel)
        {
            plotModel.Axes.Add(new LinearAxis
            {
                Title = " Dose [Gy]",
                TitleFontSize = 14,
                TitleFontWeight = OxyPlot.FontWeights.Bold,
                AxisTitleDistance = 15,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Solid,
                Position = AxisPosition.Bottom
            });
            plotModel.Axes.Add(new LinearAxis
            {
                Title = " Volume [%]",
                TitleFontSize = 14,
                TitleFontWeight = OxyPlot.FontWeights.Bold,
                AxisTitleDistance = 15,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Solid,
                Position = AxisPosition.Left
            });
        }
        public void AddDvhCurve(Structure structure)
        {
            var dvh = CalculateDvh(structure);
            PlotModel.Series.Add(CreateDvhSeries(structure.Id, dvh));
            UpdatePlot();
        }
        public void RemoveDvhCurve(Structure structure)
        {
            var series = FindSeries(structure.Id);
            PlotModel.Series.Remove(series);
            UpdatePlot();
        }
        private DVHData CalculateDvh(Structure structure)
        {
            return _plan.GetDVHCumulativeData(structure,
            DoseValuePresentation.Absolute,
            VolumePresentation.Relative, 0.01);
        }

        private OxyPlot.Series.Series CreateDvhSeries(string structureId, DVHData dvh)
        {
            var series = new OxyPlot.Series.LineSeries
            {
                Title = structureId,
                Tag = structureId,
                Color = GetStructureColor(structureId),  //Color set from Plan color directly
                StrokeThickness = GetLineThickness(structureId),
            };
            var points = dvh.CurveData.Select(CreateDataPoint);
            series.Points.AddRange(points);
            return series;
        }

        private OxyColor GetStructureColor(string structureId)
        {
            //var structures = _plan.StructureSet.Structures;
            var structure = Structures.First(x => x.structure_id.ToUpper() == structureId.ToUpper());
            var color = structure.structure.Color;
            return OxyColor.FromRgb(color.R, color.G, color.B);
        }

        private double GetLineThickness(string structureId)
        {
            if (structureId.ToUpper().Contains("PTV"))
                return 5;
            return 2;
        }

        private DataPoint CreateDataPoint(DVHPoint p)
        {
            return new DataPoint(p.DoseValue.Dose, p.Volume);
        }

        private OxyPlot.Series.Series FindSeries(string structureId)
        {
            return PlotModel.Series.FirstOrDefault(x =>
            (string)x.Tag == structureId);
        }

        private void UpdatePlot()
        {
            PlotModel.InvalidatePlot(true);
        }

    }
}
