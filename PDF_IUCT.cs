using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows;
using PDF_IUCT;

using System.Drawing;



namespace VMS.TPS
{
    public class Script
    {

        public void Execute(ScriptContext context)//, Window window)
        {

            //METTRE PLUTOT CA !!! ici
            string WORKBOOK_TEMPLATE_DIR = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //A MODIFIER POUR MODIFIER LE CHEMIN PAR DEFAUT
            string WORKBOOK_RESULT_DIR = System.IO.Path.GetTempPath();


            string working_folder = @"\\srv015\SF_COM\ARNAUD_FX\impression_pdf\";

            if (context == null)
            {
                throw new ApplicationException("Merci de charger un patient et un plan");
            }

            if (context.PlanSetup == null)
            {
                throw new ApplicationException("Merci de charger un plan");
            }
            if (context.PlanSetup.PlanIntent.Equals("VERIFICATION"))
            {
                throw new ApplicationException("Merci de charger un plan qui ne soit pas un plan de vérification");
            }
            if (!context.PlanSetup.IsDoseValid)
            {
                throw new ApplicationException("Merci de charger un plan avec une dose");
            }


            //First capture screen without any window
            string filePath = working_folder + "screenshotbis.png";
            System.Drawing.Rectangle screenBounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;

            using (Bitmap bitmap = new Bitmap(screenBounds.Width, screenBounds.Height - 35))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(screenBounds.Location, System.Drawing.Point.Empty, screenBounds.Size);
                }

                // Save the captured screenshot to a file or perform any other operations

                bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
            }

            //Then show windows and plots for printing

            Window window = new Window();
            //Then print window
            var mainViewModel = new MainViewModel(context);
            var mainView = new MainView(mainViewModel, working_folder, filePath);
            window.Title = "DVH  Plots ";
            window.Content = mainView;

            window.ShowDialog();

            //---------------------------------------------------------------------------------------------------------

        }

    }
}

