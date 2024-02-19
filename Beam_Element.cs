using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace PDF_IUCT
{
    internal class Beam_Element
    {
        Dictionary<string, string> datas = new Dictionary<string, string>();

        public void Collect_datas(ScriptContext ctx, Beam beam)
        {

            datas.Add("ID champ", beam.Id);
            datas.Add("ID machine", beam.TreatmentUnit.Id);
            datas.Add("Energie", beam.EnergyModeDisplayName);
            datas.Add("Debit de dose", beam.DoseRate.ToString() + " UM/min");

            //*********************************************
            double X1 = double.MinValue;
            double X2 = double.MinValue;
            double Y1 = double.MinValue;
            double Y2 = double.MinValue;

            double x1t = double.MinValue;
            double x2t = double.MinValue;
            double y1t = double.MinValue;
            double y2t = double.MinValue;
            foreach (ControlPoint cp in beam.ControlPoints)
            {
                X1 = cp.JawPositions.X1;
                X2 = cp.JawPositions.X2;
                Y1 = cp.JawPositions.Y1;
                Y2 = cp.JawPositions.Y2;

                if (Math.Abs(X1) >= x1t)
                    x1t = Math.Abs(X1);
                if (Math.Abs(X2) >= x2t)
                    x2t = Math.Abs(X2);
                if (Math.Abs(Y1) >= y1t)
                    y1t = Math.Abs(Y1);
                if (Math.Abs(Y2) >= y2t)
                    y2t = Math.Abs(Y2);
            }

            datas.Add("X1", Convert.ToString(x1t / 10));
            datas.Add("X2", Convert.ToString(x2t / 10));
            datas.Add("Y1", Convert.ToString(y1t / 10));
            datas.Add("Y2", Convert.ToString(y2t / 10));
            datas.Add("Taille de champ", (x1t + x2t) / 10 + " x " + (y1t + y2t) / 10 + " cm²");

            //************************************
            datas.Add("Angle du bras", Math.Round(beam.ControlPoints.FirstOrDefault().GantryAngle, 1) + "°");

            //*************************************
            if (beam.GantryDirection == GantryDirection.Clockwise)
            {
                datas.Add("Sens de rotation du bras", "SH");
            }
            else if (beam.GantryDirection == GantryDirection.CounterClockwise)
            {
                datas.Add("Sens de rotation du bras", "SAH");
            }
            else
            {
                datas.Add("Sens de rotation du bras", "-");
            }

            //*************************************
            datas.Add("Angle d'arrêt du bras", beam.ControlPoints.Last().GantryAngle == beam.ControlPoints.FirstOrDefault().GantryAngle ? "-" : beam.ControlPoints.Last().GantryAngle + "°");

            //*************************************
            datas.Add("Rot. collimateur", Math.Round(beam.ControlPoints.FirstOrDefault().CollimatorAngle, 1) + "°");
            //************************************
            datas.Add("Rot. de table", beam.ControlPoints.FirstOrDefault().PatientSupportAngle != 0 ? Math.Round(360 - beam.ControlPoints.FirstOrDefault().PatientSupportAngle, 1) + "°" : beam.ControlPoints.FirstOrDefault().PatientSupportAngle + "°");
            //************************************
            datas.Add("Bolus", (beam.Boluses == null || beam.Boluses.Count() == 0) ? "-" : beam.Boluses.FirstOrDefault().Id);
            //************************************
            datas.Add("Isocentre X", Math.Round((beam.IsocenterPosition.x - ctx.Image.UserOrigin.x) / 10, 2, MidpointRounding.AwayFromZero) + " cm");
            //************************************
            datas.Add("Isocentre Y", Math.Round((beam.IsocenterPosition.y - ctx.Image.UserOrigin.y) / 10, 2, MidpointRounding.AwayFromZero) + " cm");
            //************************************
            datas.Add("Isocentre Z", Math.Round((beam.IsocenterPosition.z - ctx.Image.UserOrigin.z) / 10, 2, MidpointRounding.AwayFromZero) + " cm");
            //************************************
            datas.Add("DSP", Math.Round(beam.SSD / 10, 1) + " cm");
            //************************************
            datas.Add("Ponderation", Convert.ToString(Math.Round(beam.WeightFactor, 3)));
            //************************************
            datas.Add("Point de référence", ctx.PlanSetup.PrimaryReferencePoint.Id);

            //*************************************
            //Elements qui dépendent de l'énergie
            switch (beam.EnergyModeDisplayName.Contains("X"))
            {
                case true:  //Si energie est Photons
                    datas.Add("Filtre en coin", beam.Wedges.Count() == 0 ? "-" : beam.Wedges.FirstOrDefault().Id);
                    datas.Add("Bloc", "-");
                    datas.Add("Plaque", "-");
                    datas.Add("Dose par fraction", Convert.ToString(beam.FieldReferencePoints.Where(x => x.IsPrimaryReferencePoint).FirstOrDefault().FieldDose));
                    break;

                default:  //Si energie est Electrons
                    datas.Add("Filtre en coin", "-");
                    datas.Add("Bloc", Convert.ToString(beam.Blocks.FirstOrDefault()));
                    datas.Add("Plaque", Convert.ToString(ctx.PlanSetup.Beams.FirstOrDefault().Blocks.First().Tray.Id));
                    datas.Add("Dose par fraction", Convert.ToString(ctx.PlanSetup.PlannedDosePerFraction));
                    break;
            }

            //************************************
            datas.Add("UM", Math.Round(beam.Meterset.Value, 1).ToString("N1") + " UM");
        }

        public Dictionary<string, string> Datas { get { return datas; } }


    }
}
