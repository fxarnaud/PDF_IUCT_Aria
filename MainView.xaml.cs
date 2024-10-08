﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OxyPlot.Wpf;
using VMS.TPS.Common.Model.API;
using System.Threading;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Interop;


using Point = System.Windows.Point;
using System.IO;



//using System.Windows.Forms;





namespace PDF_IUCT
{
    /// <summary>
    /// Logique d'interaction pour MainView.xaml
    /// </summary>
    public partial class MainView : UserControl
    {


        private static readonly PlotView PlotView = new PlotView();
        private readonly MainViewModel _vm;
        private readonly string _screenshotpath;
        private readonly string _working_folder;
        private readonly string _backup_pc_adress;
        private ToDelete files_to_delete;
        public MainView(MainViewModel viewModel, string working_folder, string screenshotpath, string backupPC_adress, ToDelete files_delete)
        {
            files_to_delete = files_delete;
            _vm = viewModel;
            _screenshotpath = screenshotpath;
            _working_folder = working_folder;
            _backup_pc_adress = backupPC_adress;
            InitializeComponent();
            DataContext = viewModel;
        }

        private Structure GetStructure(object checkBoxObject)
        {
            var checkbox = (CheckBox)checkBoxObject;
            Structure structure = _vm._ctx.StructureSet.Structures.FirstOrDefault(o => o.Id.ToUpper() == checkbox.Content.ToString().ToUpper());
            // Structure structure = _vm.Structures.FirstOrDefault(o => o.structure_id.ToUpper() == checkbox.Content.ToString().ToUpper()).structure;

            return structure;
        }

        private void CheckBox_Checked(object checkBoxObject, RoutedEventArgs e)
        {         
            StructureStatistics el = _vm.Structures.Where(x => x.structure == GetStructure(checkBoxObject)).FirstOrDefault();
            el.isChecked = true;
            _vm.AddDvhCurve(GetStructure(checkBoxObject));

        }

        private void CheckBox_Unchecked(object checkBoxObject, RoutedEventArgs e)
        {
            StructureStatistics el = _vm.Structures.Where(x => x.structure == GetStructure(checkBoxObject)).FirstOrDefault();
            el.isChecked = false;
            _vm.RemoveDvhCurve(GetStructure(checkBoxObject));
        }

        private void btn_print_Click(object sender, RoutedEventArgs e)
        {
            string message = null;
            string filepath = null;
            //ToDelete files_to_delete = new ToDelete(); //creation objet pour lister les chemins de fichiers a supprimer
            //*****Géneration du doc pdf 
            try
            {
                DocumentGenerator pdfdoc = new DocumentGenerator(_working_folder);
                
                filepath = pdfdoc.GeneratePDF(_vm._ctx, _vm.PlotModel, _vm.Structures, _screenshotpath, files_to_delete);
            }
            catch (Exception ex)
            {
                message += "Erreur lors de la génération du fichier pdf : " + ex.Message + "\n";
            }

            System.Windows.Application.Current.Windows[0].Close();

            //*****Envoi sous Aria
            try
            {
                //A DECOMMENTER POUR ENVOYER SOUS ARIA
                AriaSender asender = new AriaSender(_vm._ctx, filepath);
                message += "- Rapport dosimétrique correctement envoyé sous Aria \n";
            }
            catch (Exception ex)
            {
                message += "- Erreur : Pas d'envoi sous Aria : " + ex.Message + "\n";
            }

            //******Envoi sous PC tiers de sauvegarde    
            string todayDateString = DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString();
            string currentpath = System.IO.Path.Combine(_backup_pc_adress, todayDateString);
            try
            {
                if (!Directory.Exists(currentpath))
                {
                    Directory.CreateDirectory(currentpath);
                }
            }
            catch (Exception ex)
            {
                message += "- Warning : Veuillez noter qu'il y a une erreur lors de la creation du repertoire /Datedujour sur le poste PC0367 \n" +
                    ex;
            }
            try
            {
                string nom_fichier = _vm._ctx.Patient.LastName + "_" + _vm._ctx.Patient.FirstName + "_" + "_("+_vm._ctx.Patient.Id+")_" + _vm._ctx.PlanSetup.Id + "-"
                                    + _vm._ctx.PlanSetup.TotalDose.Dose + "Gy(" + _vm._ctx.PlanSetup.NumberOfFractions + "fr).pdf";
                string chemin_source = System.IO.Path.Combine(currentpath, nom_fichier);

                File.Copy(filepath, chemin_source, true);
                message += "- Copie de backup correctement réalisée sous PC Tiers : " + chemin_source + "\n";
                
            }
            catch
            {
                message += "- Warning ; Veuillez noter qu'il y a une erreur lors du backup du pdf \n";
            }




            //*********Netoyage des fichiers temporaires utilisés
            foreach (string path in files_to_delete.files_pathes.Values)
            {
                try
                {
                    File.Delete(path);
                }
                catch (Exception ex)
                {
                    message += "- Warning - Veuillez noter qu'il y a une erreur lors de la suppression des fichiers temporaires => Allo fxa \n";
                }
            }

            MessageBox.Show(string.Format(message));


        }
    }
}





