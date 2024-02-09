using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Fields;
using MigraDoc.Rendering;
using OxyPlot.Wpf;
using OxyPlot;
using System.Reflection;
using System.IO;
using System.Windows;

using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using System.Windows.Media;


using ImageMagick;


namespace PDF_IUCT
{
    public class DocumentGenerator
    {
        private static string _working_folder;
        private static string _image_folder;

        public DocumentGenerator(string working_folder)
        {
            _working_folder = working_folder;
            _image_folder = working_folder + "images\\";
        }
        public string GeneratePDF(ScriptContext ctx, OxyPlot.PlotModel plotmodel, IEnumerable<StructureStatistics> structures, string screenshotpath, ToDelete files)
        {

            // Creation du nouveau document migradoc
            Document document = new Document();
            document.Info.Title = "Rapport planification de radiothérapie Eclipse";
            document.Info.Author = "Claudius Regaud - Oncopole";
            DefineStyles(document);


            //Mise en page par défaut du nouveau document
            Section section = document.AddSection();
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Landscape;
            section.PageSetup.PageFormat = PageFormat.A4;
            section.PageSetup.FooterDistance = 0.8;
            section.PageSetup.LeftMargin = 20.0;

            //Ecriture des données dans le document      
            SetHeaders(document, section, ctx);
            FillPage1(ctx, document, section);
            FillPage2(ctx, document);
            FillDVH_Image(document, plotmodel, files);
            FillDVH_Statistics(document, structures);
            FillDecalages(ctx, document);
            FillScreenShot(document, screenshotpath, files);

            //Ecriture et sauvegarde du pdf
            const bool unicode = true;
            var pdfRenderer = new PdfDocumentRenderer(unicode);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            string filename = _working_folder + Guid.NewGuid().ToString() + ".pdf";
            files.files_pathes["pdf"] = filename;
            //MessageBox.Show(string.Format("{0}", filePath));
            pdfRenderer.PdfDocument.Save(filename);

            return filename;
        }
        private static void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            MigraDoc.DocumentObjectModel.Style style = document.Styles["Normal"];
            // Because all styles are derived from Normal, the next line changes the
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Times New Roman";
            // Heading1 to Heading9 are predefined styles with an outline level. An outline level
            // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks)
            // in PDF.
            style = document.Styles["Heading1"];
            style.Font.Name = "Times New Roman";
            style.Font.Size = 14;
            style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;
            style = document.Styles["Heading2"];
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 6;
            style = document.Styles["Heading3"];
            style.Font.Size = 10;
            style.Font.Bold = true;
            style.Font.Italic = true;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 3;
            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", MigraDoc.DocumentObjectModel.TabAlignment.Right);
            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", MigraDoc.DocumentObjectModel.TabAlignment.Center);
        }

        #region En tete
        private static void SetHeaders(Document document, Section section, ScriptContext ctx)
        {
            // Section section = document.AddSection();
            section.PageSetup.DifferentFirstPageHeaderFooter = false;

            //Top of the page margin
            section.PageSetup.TopMargin = "0.5cm"; // Replace "1cm" with your desired header height


            //Footers
            var footer = section.Footers.Primary.AddParagraph();
            footer.Format.Borders.Top.Visible = true;

            MigraDoc.DocumentObjectModel.Tables.Table table = section.Footers.Primary.AddTable();
            table.Borders.Visible = false;

            MigraDoc.DocumentObjectModel.Tables.Column column = table.AddColumn();
            column.Width = Unit.FromPoint(20);
            column.Format.Alignment = ParagraphAlignment.Left;
            MigraDoc.DocumentObjectModel.Tables.Column column1 = table.AddColumn();
            column1.Width = Unit.FromPoint(340);
            // column1.Format.Alignment = ParagraphAlignment.Left;
            MigraDoc.DocumentObjectModel.Tables.Column column2 = table.AddColumn();
            column.Width = Unit.FromPoint(340);
            column.Format.Alignment = ParagraphAlignment.Right;
            MigraDoc.DocumentObjectModel.Tables.Row row = table.AddRow();


            //Remplir la cellule de gauche
            MigraDoc.DocumentObjectModel.Tables.Cell cell = row.Cells[0];
            var paragraph1 = cell.AddParagraph("Rappport de planification de radiothérapie (v" + ctx.VersionInfo + ") de " + ctx.Patient.LastName + ", " + ctx.Patient.FirstName + "\n Plan : " + ctx.PlanSetup.Id + " \n Rapport créé le : " + DateTime.Now.ToString() + " par : " + ctx.CurrentUser.Name);
            paragraph1.Format.Alignment = ParagraphAlignment.Left;

            cell = row.Cells[1];
            var paragraph2 = row.Cells[1].AddParagraph();
            paragraph2.Format.Alignment = ParagraphAlignment.Center;
            string filename = _image_folder + "logo.jpg";
            paragraph2.AddImage(filename);

            cell = row.Cells[2];
            var paragraph3 = row.Cells[2].AddParagraph();
            paragraph3.Format.Alignment = ParagraphAlignment.Right;
            paragraph3.AddPageField();
            paragraph3.AddText("/");
            paragraph3.AddNumPagesField();

        }
        #endregion

        #region Page 1 : Nom patient et données du plan
        static void FillPage1(ScriptContext ctx, Document document, Section section)
        {
            // var section = document.AddSection();

            //Creation du premier tableau avec les infos du patient
            MigraDoc.DocumentObjectModel.Tables.Table table = new MigraDoc.DocumentObjectModel.Tables.Table();
            table.Borders.Visible = false;

            MigraDoc.DocumentObjectModel.Tables.Column column0 = table.AddColumn(Unit.FromCentimeter(12));
            MigraDoc.DocumentObjectModel.Tables.Column column1 = table.AddColumn(Unit.FromCentimeter(10));
            MigraDoc.DocumentObjectModel.Tables.Column column2 = table.AddColumn(Unit.FromCentimeter(5));
            column0.Format.Alignment = ParagraphAlignment.Left;
            column1.Format.Alignment = ParagraphAlignment.Left;
            column2.Format.Alignment = ParagraphAlignment.Left;

            MigraDoc.DocumentObjectModel.Tables.Row row1 = table.AddRow();
            row1.Format.Font.Bold = true;
            row1.Format.Font.Size = 13;
            row1.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[0].AddParagraph("PATIENT");
            MigraDoc.DocumentObjectModel.Tables.Row row2 = table.AddRow();
            MigraDoc.DocumentObjectModel.Tables.Row row3 = table.AddRow();
            row3.Cells[0].AddParagraph("Nom : " + ctx.Patient.Name);
            row3.Cells[1].AddParagraph("Date de Naissance : " + ctx.Patient.DateOfBirth.ToString().Substring(0, 10));
            string sexe = ctx.Patient.Sex == "Female" ? "Féminin" : "Masculin";
            row3.Cells[2].AddParagraph("Sexe : " + sexe);
            document.LastSection.Add(table);

            section.AddParagraph("");
            section.AddParagraph("");

            //Creation du deuxieme tableau avec les données du plan et de prescription
            table = new MigraDoc.DocumentObjectModel.Tables.Table();
            table.Borders.Visible = false;
            column0 = table.AddColumn(Unit.FromCentimeter(6));
            column1 = table.AddColumn(Unit.FromCentimeter(7.5));
            column2 = table.AddColumn(Unit.FromCentimeter(1));
            MigraDoc.DocumentObjectModel.Tables.Column column3 = table.AddColumn(Unit.FromCentimeter(5));
            MigraDoc.DocumentObjectModel.Tables.Column column4 = table.AddColumn(Unit.FromCentimeter(7));
            row1 = table.AddRow();

            row1.Cells[0].AddParagraph("PLAN");
            row1.Cells[0].Format.Font.Bold = true;
            row1.Cells[0].Format.Font.Size = 13;
            row1.Cells[0].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[1].Format.Font.Bold = true;
            row1.Cells[1].Format.Font.Size = 13;
            row1.Cells[1].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);

            row1.Cells[3].AddParagraph("PRESCRIPTION DE DOSE");
            row1.Cells[3].Format.Font.Bold = true;
            row1.Cells[3].Format.Font.Size = 13;
            row1.Cells[3].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[4].Format.Font.Bold = true;
            row1.Cells[4].Format.Font.Size = 13;
            row1.Cells[4].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1 = table.AddRow();

            Dictionary<string, string> plan_datas = new Dictionary<string, string>()
            {
                {"ID plan", ctx.PlanSetup.Id },
                {"ID dossier de traitement : ", ctx.Course.Id},
                {"ID image : ", ctx.Image.Id},
                {"Système d'imagerie : ", ctx.Image.Series.ImagingDeviceId},
                {"Orientation du traitement : ", Translate( ctx.Image.ImagingOrientation.ToString())},
                {"Décalage de l'origine/à DICOM : ","(" + (ctx.Image.UserOrigin.x / 10).ToString("N2") + " cm," + (ctx.Image.UserOrigin.y / 10).ToString("N2") + " cm," + (ctx.Image.UserOrigin.z / 10).ToString("N2") + "cm)" },
                {"Type de traitement : ", GetTreatmentType(ctx.PlanSetup)}
            };
            Dictionary<string, string> prescription_datas = new Dictionary<string, string>()
            {
                { "Méthode normalisation ", ctx.PlanSetup.PlanNormalizationMethod },
                { "Volume cible : ", ctx.PlanSetup.TargetVolumeID},
                { "Point de reference principal ", ctx.PlanSetup.PrimaryReferencePoint.Id},
                { "Valeur de normalisation : ", ctx.PlanSetup.PlanNormalizationValue.ToString("N1") + " %"},
                { "Pourcentage de dose prescrite ",(ctx.PlanSetup.TreatmentPercentage * 100).ToString("N1") + "%" },
                { "Fractionnement : ", "" },
                { "Dose Prescrite : ", ctx.PlanSetup.TotalDose.ValueAsString + "Gy (" + ctx.PlanSetup.DosePerFraction.ValueAsString + "Gy/fraction)"},
                { "Nombre de fractions : ", ctx.PlanSetup.NumberOfFractions.ToString()},
            };


            for (int i = 0; i < prescription_datas.Count(); i++)
            {
                row1 = table.AddRow();
                row1.Cells[3].AddParagraph(prescription_datas.ElementAt(i).Key);
                row1.Cells[3].Format.Font.Bold = false;
                row1.Cells[4].AddParagraph(prescription_datas.ElementAt(i).Value);
                if (i < 6)
                {
                    row1.Cells[3].Format.Font.Bold = true;
                }
                else
                {
                    row1.Cells[3].Format.Alignment = ParagraphAlignment.Right;
                }

                //MessageBox.Show(string.Format("{0} et {1}", prescription_datas.ElementAt(i).Key, prescription_datas.ElementAt(i).Value));
                if (i < plan_datas.Count())
                {
                    row1.Cells[0].AddParagraph(plan_datas.ElementAt(i).Key);
                    row1.Cells[0].Format.Font.Bold = true;
                    row1.Cells[1].AddParagraph(plan_datas.ElementAt(i).Value);
                }
            }

            document.LastSection.Add(table);
            section.AddParagraph("");
            section.AddParagraph("");

            //TABLEAU OPTIONS DE CALCUL
            table = new MigraDoc.DocumentObjectModel.Tables.Table();
            table.Borders.Visible = false;
            column0 = table.AddColumn(Unit.FromCentimeter(7.5));
            column1 = table.AddColumn(Unit.FromCentimeter(6));
            column1.Format.Alignment = ParagraphAlignment.Right;
            row1 = table.AddRow();
            row1.Cells[0].AddParagraph("OPTIONS DE CALCUL");
            row1.Cells[0].Format.Font.Bold = true;
            row1.Cells[0].Format.Font.Size = 13;
            row1.Cells[0].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[1].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1 = table.AddRow();


            var algoX = ctx.PlanSetup.Beams.Where(x => x.IsSetupField == false).FirstOrDefault().EnergyModeDisplayName.Contains("X") ? true : false;
            row1 = table.AddRow();
            row1.Cells[0].AddParagraph(algoX == true ? "Algorithme photon : " : "Algorithme electron : ");
            row1.Cells[0].Format.Font.Bold = true;
            row1.Cells[1].AddParagraph(algoX == true ? ctx.PlanSetup.PhotonCalculationModel : ctx.PlanSetup.ElectronCalculationModel);

            var cmodel = algoX == true ? ctx.PlanSetup.PhotonCalculationOptions : ctx.PlanSetup.ElectronCalculationOptions;
            foreach (var option in cmodel)
            {
                row1 = table.AddRow();
                row1.Cells[0].AddParagraph(option.Key);
                row1.Cells[0].Format.Alignment = ParagraphAlignment.Right;
                row1.Cells[1].AddParagraph(option.Value);
            }

            document.LastSection.Add(table);

            //Creation du tableau avec les approbations
            section.AddParagraph("");
            section.AddParagraph("");

            table = new MigraDoc.DocumentObjectModel.Tables.Table();
            table.Borders.Visible = false;
            column0 = table.AddColumn(Unit.FromCentimeter(5));
            column1 = table.AddColumn(Unit.FromCentimeter(8.5));
            row1 = table.AddRow();
            row1.Cells[0].AddParagraph("APPROBATIONS");
            row1.Cells[0].Format.Font.Bold = true;
            row1.Cells[0].Format.Font.Size = 13;
            row1.Cells[0].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[1].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1 = table.AddRow();
            row1 = table.AddRow();
            row1.Cells[0].AddParagraph("Plan créé :");
            row1.Cells[1].AddParagraph(ctx.PlanSetup.CreationUserName + " (" + ctx.PlanSetup.CreationDateTime + ")");
            row1 = table.AddRow();
            row1.Cells[0].AddParagraph("Appro. plan :");
            row1.Cells[1].AddParagraph(ctx.PlanSetup.PlanningApproverDisplayName == "" ? "-" : ctx.PlanSetup.PlanningApproverDisplayName + " (" + ctx.PlanSetup.PlanningApprovalDate + ")");
            row1 = table.AddRow();
            row1.Cells[0].AddParagraph("Appro. traitement :");
            row1.Cells[1].AddParagraph(ctx.PlanSetup.TreatmentApproverDisplayName == "" ? "-" : ctx.PlanSetup.TreatmentApproverDisplayName + " (" + ctx.PlanSetup.TreatmentApprovalDate + ")");

            document.LastSection.Add(table);

        }

        private static string Translate(string to_translate)
        {
            string translated;
            switch (to_translate.ToUpper())
            {
                case ("HEADFIRSTSUPINE"):
                    translated = "Tete devant - DD";
                    break;
                case ("FEETFIRSTSUPINE"):
                    translated = "Pieds devant - DD";
                    break;
                case ("HEADFIRSTPRONE"):
                    translated = "Tete devant - DV";
                    break;
                default:
                    translated = "Indefinie";
                    break;
            }
            return translated;
        }

        private static string GetTreatmentType(PlanSetup plan)
        {
            string TT_type = "indefini";

            if (plan.Beams.Any(b => (b.MLCPlanType == VMS.TPS.Common.Model.Types.MLCPlanType.ArcDynamic)))
            {
                TT_type = "ArcThérapie Dynamique (DCA)";
            }

            else if (plan.Beams.Any(b => (b.MLCPlanType == VMS.TPS.Common.Model.Types.MLCPlanType.VMAT)))
            {
                TT_type = "VMAT (RA)";
            }

            else if (plan.Beams.Any(b => (b.MLCPlanType == VMS.TPS.Common.Model.Types.MLCPlanType.DoseDynamic)))
            {
                TT_type = "IMRT statique";
            }
            else
            {
                TT_type = "RTC 3D";
            }
            return TT_type;
        }
        #endregion

        #region Page 2 : Parametres des champs de traitement
        static void FillPage2(ScriptContext ctx, Document document)
        {

            //Recuperation des donées des champs puis découpage en sous tableau pour avoir 6 elements max dans chaque tableau.
            //Chaque sous tableau représente une page du documetn avec 6 champs.
            List<Beam_Element> Beam_Elements = new List<Beam_Element>();
            foreach (Beam beam in ctx.PlanSetup.Beams)
            {
                if (!beam.IsSetupField)
                {
                    Beam_Element beam_el = new Beam_Element();
                    beam_el.Collect_datas(ctx, beam);
                    Beam_Elements.Add(beam_el);
                }
            }
            //Selon le nombre de champs, splitte le tableau global avec les champs en sous tableaux ayant 6 elements max
            List<List<Beam_Element>> resultLists = SplitList(Beam_Elements, 6);

            int comp = 1;
            foreach (List<Beam_Element> lstbeams in resultLists)
            {
                //Ajout de la section et saut de page
                var section = document.AddSection();
                section.AddPageBreak();

                //Sur chaque page, mettre un tableau pour le titre principal de la page
                MigraDoc.DocumentObjectModel.Tables.Table table_title = new MigraDoc.DocumentObjectModel.Tables.Table();
                MigraDoc.DocumentObjectModel.Tables.Column column = table_title.AddColumn(Unit.FromCentimeter(27));
                column.Format.Alignment = ParagraphAlignment.Left;
                column.Borders.Visible = false;
                MigraDoc.DocumentObjectModel.Tables.Row row1 = table_title.AddRow();
                row1.Format.Font.Bold = true;
                row1.Format.Font.Size = 13;
                row1.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
                string titre = "PARAMETRES DES FAISCEAUX DE TRAITEMENT " + "(" + comp + "/" + resultLists.Count().ToString() + ")";
                row1.Cells[0].AddParagraph(titre);
                row1 = table_title.AddRow();
                row1.Format.Font.Size = 16;
                row1.Cells[0].AddParagraph("");
                document.LastSection.Add(table_title);
             


                //TABLEAU AVEC DONNEES DES CHAMPS
                MigraDoc.DocumentObjectModel.Tables.Table beam_table = new MigraDoc.DocumentObjectModel.Tables.Table();
                //Pour migradoc il faut faire le slignes et colonnes et ensuite remplir le tableau
                for (int colonne = 0; colonne < lstbeams.Count()+1; colonne++)  //je fais +1 car je vais ajouter une colonne au debut de chaque tableau pour avoir le nom des items
                {
                    MigraDoc.DocumentObjectModel.Tables.Column beam_column = beam_table.AddColumn(Unit.FromCentimeter(4));
                    beam_column.Format.Alignment = ParagraphAlignment.Left;
                    beam_column.Borders.Visible = true;
                }
                //Then add rows
                for (int ligne = 0; ligne < lstbeams[0].Datas.Count(); ligne++)
                {
                    MigraDoc.DocumentObjectModel.Tables.Row beam_row = beam_table.AddRow();
                }

                // Then fill the table
                // fill first column with beam items
                for (int ligne = 0; ligne < lstbeams[0].Datas.Count(); ligne++)
                {
                    beam_table.Rows[ligne].Cells[0].AddParagraph(lstbeams[0].Datas.ElementAt(ligne).Key);
                    beam_table.Rows[ligne].Cells[0].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(245, 235, 235);
                    beam_table.Rows[ligne].Cells[0].Format.Font.Bold = true;
                }
                // Then fill other columns with beam datas
                int colo = 1;
                foreach (var beam in lstbeams)
                {
                    for (int ligne = 0; ligne < beam.Datas.Count(); ligne++)
                    {   
                        beam_table.Rows[ligne].Cells[colo].AddParagraph(beam.Datas.ElementAt(ligne).Value);

                        if (ligne == 0)
                        {
                            beam_table.Rows[ligne].Cells[colo].Borders.Visible = true;
                            beam_table.Rows[ligne].Cells[colo].Format.Font.Bold = true;
                            beam_table.Rows[ligne].Cells[colo].Format.Font.Size = 14;
                            beam_table.Rows[ligne].Cells[colo].Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(245, 235, 235);
                        }
                        if (beam.Datas.ElementAt(ligne).Key.Equals("UM"))
                        {
                            beam_table.Rows[ligne].Cells[colo].Format.Font.Bold = true;
                        }
                        if (beam.Datas.ElementAt(ligne).Key.Equals("Bolus"))
                        {
                            beam_table.Rows[ligne].Cells[colo].Format.Font.Bold = true;
                            beam_table.Rows[ligne].Cells[colo].Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Red;
                        }
                    }
                    colo++;
                }


                // Add the table to the document
                document.LastSection.Add(beam_table);
                comp++;
            }
        }


        static List<List<T>> SplitList<T>(List<T> sourceList, int chunkSize)
        {
            List<List<T>> result = new List<List<T>>();

            for (int i = 0; i < sourceList.Count; i += chunkSize)
            {
                List<T> chunk = sourceList.Skip(i).Take(chunkSize).ToList();
                result.Add(chunk);
            }

            return result;
        }

        static List<List<T>> SplitListEnumerable<T>(IEnumerable<T> sourceEnumerable, int chunkSize)
        {
            List<T> sourceList = sourceEnumerable.ToList(); // Convert to List
            List<List<T>> result = new List<List<T>>();

            for (int i = 0; i < sourceList.Count; i += chunkSize)
            {
                List<T> chunk = sourceList.Skip(i).Take(chunkSize).ToList();
                result.Add(chunk);
            }

            return result;
        }

        #endregion

        #region DVH et statistiques
        static void FillDVH_Image(Document document, OxyPlot.PlotModel plotmodel, ToDelete files_to_delete)
        {
            Section section = document.AddSection();
            section.AddPageBreak();

            //Tableau pour mettre un titre à la page
            MigraDoc.DocumentObjectModel.Tables.Table table_title = new MigraDoc.DocumentObjectModel.Tables.Table();
            MigraDoc.DocumentObjectModel.Tables.Column column = table_title.AddColumn(Unit.FromCentimeter(27));
            column.Format.Alignment = ParagraphAlignment.Left;
            column.Borders.Visible = false;
            MigraDoc.DocumentObjectModel.Tables.Row row1 = table_title.AddRow();
            row1.Format.Font.Bold = true;
            row1.Format.Font.Size = 13;
            row1.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[0].AddParagraph("HISTOGRAMME DOSE VOLUME");
            row1 = table_title.AddRow();
            row1.Format.Font.Size = 16;
            row1.Cells[0].AddParagraph("");
            document.LastSection.Add(table_title);
            string path = _working_folder + Guid.NewGuid().ToString() +".png";
            files_to_delete.files_pathes["dvh"] = path; // recupere tous les files to deletepour les supprimer à la toute fin
            PngExporter.Export(plotmodel, path, 900, 500, OxyColors.White);
            section.AddImage(path);
        }

        static void FillDVH_Statistics(Document document, IEnumerable<StructureStatistics> structures)
        {

            //ICI FAIRE SPLIT LIST    
            List<List<StructureStatistics>> resultLists = SplitListEnumerable(structures, 35);
            int compteur = 1;

            Section section = document.AddSection();

            foreach (List<StructureStatistics> list in resultLists)
            {
                section.AddPageBreak();
                //First create first header table
                MigraDoc.DocumentObjectModel.Tables.Table table_header = CreateHeaderTable(compteur, resultLists.Count());
                document.LastSection.Add(table_header);

                //Then create empty table
                //Add an empty table  
                MigraDoc.DocumentObjectModel.Tables.Table table_stat = Create_DVHStat_Empty_Table();
                MigraDoc.DocumentObjectModel.Color migraDocColor = new MigraDoc.DocumentObjectModel.Color(243, 228, 228);

                foreach (StructureStatistics structure in list)
                {
                    if (structure.isChecked)
                    {
                        MigraDoc.DocumentObjectModel.Tables.Row row = table_stat.AddRow();
                        row.Cells[0].Shading.Color = migraDocColor;
                        row.Cells[1].AddParagraph(structure.structure_id);
                        row.Cells[2].AddParagraph(structure.volume.ToString("N1"));
                        row.Cells[3].AddParagraph(structure.maxdose.ToString("N1"));
                        row.Cells[4].AddParagraph(structure.meandose.ToString("N1"));
                        row.Cells[5].AddParagraph(structure.d1cc.ToString("N1"));
                        row.Cells[6].AddParagraph(structure.d0035cc.ToString("N1"));                        
                    }
                }
                document.LastSection.Add(table_stat);

                compteur++;
            }

            //int i = 0;
            //foreach (StructureStatistics structure in structures)
            //{
            //    if (structure.isChecked)
            //    {
            //        if (i == 20)  //On est en bout depage, faire sur une nouvelle page
            //        {
            //            document.LastSection.Add(table);
            //            section.AddPageBreak();
            //            MigraDoc.DocumentObjectModel.Tables.Table table_title2 = CreateHeaderTable();
            //            document.LastSection.Add(table_title2);
            //            //document.LastSection.AddParagraph("Statistiques du DVH", "Heading1");
            //            table = Create_DVHStat_Empty_Table();
            //            i = 0;
            //        }




            //        // Get the RGB values from the SolidColorBrush
            //        byte r = structure.BackgroundColor.Color.R;
            //        byte g = structure.BackgroundColor.Color.G;
            //        byte b = structure.BackgroundColor.Color.B;

            //        // Create a MigraDoc color with the RGB values

            //        //Color migraDocColor = new Color(r, g, b);


            //    }
            //}

        }
        static MigraDoc.DocumentObjectModel.Tables.Table CreateHeaderTable(int compteur, int count)
        {
            //Tableau pour mettre un titre à la page
            MigraDoc.DocumentObjectModel.Tables.Table table_title = new MigraDoc.DocumentObjectModel.Tables.Table();
            MigraDoc.DocumentObjectModel.Tables.Column column = table_title.AddColumn(Unit.FromCentimeter(27));
            column.Format.Alignment = ParagraphAlignment.Left;
            column.Borders.Visible = false;
            MigraDoc.DocumentObjectModel.Tables.Row row1 = table_title.AddRow();
            row1.Format.Font.Bold = true;
            row1.Format.Font.Size = 13;
            row1.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[0].AddParagraph("STATISTIQUES HISTOGRAMME DOSE-VOLUME (" + compteur + "/" + count + ")");
            row1 = table_title.AddRow();
            row1.Format.Font.Size = 16;
            row1.Cells[0].AddParagraph("");

            return table_title;
        }

        static MigraDoc.DocumentObjectModel.Tables.Table Create_DVHStat_Empty_Table()
        {
            MigraDoc.DocumentObjectModel.Tables.Table table = new MigraDoc.DocumentObjectModel.Tables.Table();
            table.Borders.Width = 0.75;
            MigraDoc.DocumentObjectModel.Tables.Column column0 = table.AddColumn(Unit.FromCentimeter(1));//
            MigraDoc.DocumentObjectModel.Tables.Column column1 = table.AddColumn(Unit.FromCentimeter(5));
            MigraDoc.DocumentObjectModel.Tables.Column column2 = table.AddColumn(Unit.FromCentimeter(4));
            MigraDoc.DocumentObjectModel.Tables.Column column3 = table.AddColumn(Unit.FromCentimeter(4));
            MigraDoc.DocumentObjectModel.Tables.Column column4 = table.AddColumn(Unit.FromCentimeter(4));
            MigraDoc.DocumentObjectModel.Tables.Column column5 = table.AddColumn(Unit.FromCentimeter(4));
            MigraDoc.DocumentObjectModel.Tables.Column column6 = table.AddColumn(Unit.FromCentimeter(4));
            column0.Format.Alignment = ParagraphAlignment.Center;
            column1.Format.Alignment = ParagraphAlignment.Center;
            column2.Format.Alignment = ParagraphAlignment.Center;
            column3.Format.Alignment = ParagraphAlignment.Center;
            column4.Format.Alignment = ParagraphAlignment.Center;
            column5.Format.Alignment = ParagraphAlignment.Center;
            column6.Format.Alignment = ParagraphAlignment.Center;

            MigraDoc.DocumentObjectModel.Tables.Row row1 = table.AddRow();
            row1.Format.Font.Bold = true;
            row1.Format.Font.Size = 12;
            row1.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(200, 200, 200);
            row1.Cells[0].AddParagraph("");  //
            row1.Cells[1].AddParagraph("Structure");
            row1.Cells[2].AddParagraph("Volume(cc)");
            row1.Cells[3].AddParagraph("Dose Max(Gy)");
            row1.Cells[4].AddParagraph("Dose Moyenne(Gy)");
            row1.Cells[5].AddParagraph("Dose 1cc(Gy)");
            row1.Cells[6].AddParagraph("Dose 0.035cc(Gy)");

            return table;
        }


        #endregion

        #region Decalages
        private static void FillDecalages(ScriptContext ctx, Document document)
        {
            //Ajout de la section et saut de page
            var section = document.AddSection();
            section.AddPageBreak();

            //Tableau pour mettre un titre à la page
            MigraDoc.DocumentObjectModel.Tables.Table table_title = new MigraDoc.DocumentObjectModel.Tables.Table();
            MigraDoc.DocumentObjectModel.Tables.Column column = table_title.AddColumn(Unit.FromCentimeter(27));
            column.Format.Alignment = ParagraphAlignment.Left;
            column.Borders.Visible = false;
            MigraDoc.DocumentObjectModel.Tables.Row row1 = table_title.AddRow();
            row1.Format.Font.Bold = true;
            row1.Format.Font.Size = 13;
            row1.Shading.Color = MigraDoc.DocumentObjectModel.Color.FromRgb(175, 220, 232);
            row1.Cells[0].AddParagraph("DECALAGES DE LA TABLE PAR RAPPORT A LA POSITION DE REFERENCE");
            row1 = table_title.AddRow();
            row1.Format.Font.Size = 16;
            row1.Cells[0].AddParagraph("");
            document.LastSection.Add(table_title);

            //Recupere les différents isocentres sils existent
            List<Beam_isocenter> beams_isos = new List<Beam_isocenter>();
            foreach (Beam beam in ctx.PlanSetup.Beams)
            {
                if (!beam.IsSetupField)
                {
                    if (!beams_isos.Any(o => o.Isocenter.Equals(beam.IsocenterPosition)))
                    {
                        Beam_isocenter bm = new Beam_isocenter();
                        bm.Isocenter = beam.IsocenterPosition;
                        bm.beamid.Add(beam.Id);
                        bm.iso_x = Math.Round((beam.IsocenterPosition.x - ctx.Image.UserOrigin.x) / 10, 2, MidpointRounding.AwayFromZero);
                        bm.iso_y = Math.Round((beam.IsocenterPosition.y - ctx.Image.UserOrigin.y) / 10, 2, MidpointRounding.AwayFromZero);
                        bm.iso_z = Math.Round((beam.IsocenterPosition.z - ctx.Image.UserOrigin.z) / 10, 2, MidpointRounding.AwayFromZero);
                        beams_isos.Add(bm);
                    }
                    else
                    {
                        Beam_isocenter bm = beams_isos.FirstOrDefault(x => x.Isocenter.Equals(beam.IsocenterPosition));
                        bm.beamid.Add(beam.Id);
                    }

                }
            }

            //Pour chacun des iscocentre détecté, mettre les schéma des décalages
            MigraDoc.DocumentObjectModel.Tables.Table table = new MigraDoc.DocumentObjectModel.Tables.Table();
            table.Borders.Visible = false;
            column = table.AddColumn(Unit.FromCentimeter(3));
            column = table.AddColumn(Unit.FromCentimeter(8));
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn(Unit.FromCentimeter(8));
            column.Format.Alignment = ParagraphAlignment.Center;
            column = table.AddColumn(Unit.FromCentimeter(8));
            column.Format.Alignment = ParagraphAlignment.Center;

            foreach (Beam_isocenter beami in beams_isos)
            {
                //Ajout d'une premiere lige pour lister les champs concernés
                MigraDoc.DocumentObjectModel.Tables.Row row = table.AddRow();
                row.Cells[0].MergeRight = 3;
                row.Borders.Visible = true;
                row.Format.Font.Bold = true;
                row.Format.Font.Size = 12;

                string mess = "Isocentre des champs : ";
                for (int i = 0; i < beami.beamid.Count(); i++)
                {
                    mess = mess + ", " + beami.beamid[i].ToString();
                }

                row.Cells[0].AddParagraph(mess);
                row.Cells[0].MergeRight = 3;
                row = table.AddRow();
                row.Cells[0].AddParagraph("");

                //prendre l'opposé de l'isocentre pour qu'on ait les decalages à réaliser
                beami.iso_x = -beami.iso_x;
                beami.iso_y = -beami.iso_y;
                beami.iso_z = beami.iso_z;  //Decalage à faire dans le même sens que la position de table

                //Ajout d'une nouvelle ligne pour indiquer les valeurs des décalages
                row = table.AddRow();
                row.Cells[0].AddParagraph("Décaler :");
                row.Cells[1].AddParagraph(beami.iso_x.ToString() + " cm");
                row.Cells[2].AddParagraph(beami.iso_y.ToString() + " cm");
                row.Cells[3].AddParagraph(beami.iso_z.ToString() + " cm");
                row = table.AddRow();
                row.Cells[0].AddParagraph("");

                //Nouvelle ligne pour mettre les directions de décalages
                row = table.AddRow();
                row.Cells[0].AddParagraph("Direction :");
                string dirx = "aucune";
                dirx = beami.iso_x > 0 ? "Table vers la droite en regardant le bras" : "Table vers la  gauche en regardant le bras";
                string diry = "aucune";
                diry = beami.iso_y > 0 ? "Table vers le bas" : "Table vers le haut";
                string dirz = "aucune";
                dirz = beami.iso_z > 0 ? "Table en sortie (out)" : "Table en long vers le bras (in)";
                row.Cells[1].AddParagraph(dirx);
                row.Cells[2].AddParagraph(diry);
                row.Cells[3].AddParagraph(dirz);
                row = table.AddRow();
                row.Cells[0].AddParagraph("");

                //Dernière ligne pour mettre les schéma des décalages à faire
                row = table.AddRow();
                row.Format.Alignment = ParagraphAlignment.Center;
                string imagex = beami.iso_x == 0 ? _image_folder + @"Images\aucun.png" : //_image_folder + aucun.png" :
                                    beami.iso_x > 0 ?
                                        _image_folder + "x_negatif.png" : _image_folder + "x_positif.png";
                string imagey = beami.iso_y == 0 ? _image_folder + "aucun.png" :
                                    beami.iso_y > 0 ?
                                        _image_folder + "y_negatif.png" : _image_folder + "y_positif.png";
                string imagez = beami.iso_z == 0 ? _image_folder + "aucun.png" :
                                    beami.iso_z > 0 ?
                                        _image_folder + "z_positif.png" : _image_folder + "z_negatif.png";


                var cell = row.Cells[1];
                Paragraph paragraph = cell.AddParagraph();
                var image = paragraph.AddImage(imagex);
                paragraph.Format.Alignment = ParagraphAlignment.Center;

                cell = row.Cells[2];
                paragraph = cell.AddParagraph();
                image = paragraph.AddImage(imagey);
                paragraph.Format.Alignment = ParagraphAlignment.Center;

                cell = row.Cells[3];
                paragraph = cell.AddParagraph();
                image = paragraph.AddImage(imagez);
                paragraph.Format.Alignment = ParagraphAlignment.Center;

                row = table.AddRow();
                row.Cells[0].AddParagraph("");
            }
            document.LastSection.Add(table);
        }

        public class Beam_isocenter
        {
            public Beam_isocenter()
            {
                beamid = new List<string>();
            }
            public VMS.TPS.Common.Model.Types.VVector Isocenter { get; set; }
            public List<string> beamid { get; set; }

            public double iso_x { get; set; }

            public double iso_y { get; set; }
            public double iso_z { get; set; }

        }

        #endregion

        #region Vue triangulaire
        private static void FillScreenShot(Document document, string screenshotpath_in, ToDelete files_to_delete)
        {
            Section section = document.AddSection();
            section.AddPageBreak();

            string screenshotpath_out = _working_folder + Guid.NewGuid().ToString() + ".png";
            RescaleImage(screenshotpath_in, screenshotpath_out, 1024, 768);
            section.AddImage(screenshotpath_out);
            files_to_delete.files_pathes["sshot"] = screenshotpath_out; //ajout ici pour recuperer tous les fichiers à supprimer et les supprimer à la fin
        }

        public static void RescaleImage(string sourceImagePath, string outputImagePath, int newWidth, int newHeight)
        {

            using (var image = new MagickImage(sourceImagePath))
            {
                image.FilterType = FilterType.Lanczos;
                image.Resize(newWidth, newHeight);
                image.Write(outputImagePath); // Save the rescaled image to the output path
            }
        }

        //public static void SaveScreenshotToFile(Bitmap screenshot, string outputPath)
        //{
        //    screenshot.Save(outputPath, System.Drawing.Imaging.ImageFormat.Png);
        //}

        #endregion



    }

}
