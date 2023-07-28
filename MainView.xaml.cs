using System;
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
        public MainView(MainViewModel viewModel, string working_folder, string screenshotpath)
        {
            _vm = viewModel;
            _screenshotpath = screenshotpath;
            _working_folder = working_folder;
            InitializeComponent();
            DataContext = viewModel;
        }

        private Structure GetStructure(object checkBoxObject)
        {
            var checkbox = (CheckBox)checkBoxObject;
            Structure structure = _vm.Structures.FirstOrDefault(o => o.structure_id.ToUpper() == checkbox.Content.ToString().ToUpper()).structure;
            return structure;
        }

        private void CheckBox_Checked(object checkBoxObject, RoutedEventArgs e)
        {
            _vm.AddDvhCurve(GetStructure(checkBoxObject));
        }

        private void CheckBox_Unchecked(object checkBoxObject, RoutedEventArgs e)
        {
            _vm.RemoveDvhCurve(GetStructure(checkBoxObject));
        }

        private void btn_print_Click(object sender, RoutedEventArgs e)
        {

            string filepath = null;
            //Géneration du doc pdf 
            try
            {
                DocumentGenerator pdfdoc = new DocumentGenerator(_working_folder);
                filepath = pdfdoc.GeneratePDF(_vm._ctx, _vm.PlotModel, _vm.Structures, _screenshotpath);
                //MessageBox.Show(string.Format("{0}", filepath));
            }
            catch (Exception ex)
            {
                // Le bloc catch capture l'exception et affiche un message d'erreur
                MessageBox.Show(string.Format("Erreur lors de la génération du fichier pdf : " + ex.Message));
            }

            System.Windows.Application.Current.Windows[0].Close();

            //Envoi sous Aria
            try
            {
                //A DECOMMENTER POUR ENVOYER SOUS ARIA
                AriaSender asender = new AriaSender(_vm._ctx, filepath);
                MessageBox.Show(string.Format("Rapport dosimétrique envoyé sous Aria"));
            }
            catch (Exception ex)
            {
                // Le bloc catch capture l'exception et affiche un message d'erreur
                MessageBox.Show(string.Format("Erreur lors de l'envoi sous Aria : " + ex.Message));
            }

        }
    }
}





