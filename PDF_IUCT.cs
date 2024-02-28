using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows;
using PDF_IUCT;
using System.IO;
using System.Windows.Forms;
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
            string working_folder = @"\\srv015\radiotherapie\SCRIPTS_ECLIPSE\PDF_IUC\datas_nepaseffacer\";
            string backupPC_adress = @"\\PC0367\Users\IUCTO_LIMBUS\Desktop\Backup_pdf\";

            PreliminaryTests.Test(context);
             
            
            string filePath = working_folder + Guid.NewGuid().ToString() + ".png";
            ToDelete files_to_delete = new ToDelete();
            //files.files_pathes["pdf"] = filename;
            files_to_delete.files_pathes["sshot1"] = filePath;

            #region GESTION DES ECRANS
            //***************************************
            //First capture screen depending if i'm launching my screen in secondary or primary screen

            Screen[] screens = Screen.AllScreens;
            if (screens.Length > 2)
            {
                throw new ApplicationException("Plus de 2 écrans non pris en charge");
            }
            else
            {
                Rectangle screenBounds = new Rectangle();

                switch (screens.Length)
                {
                    case 2: //two screens used

                        Screen primaryScreen = screens[0];
                        Screen secondaryScreen = screens[1];
                        if (primaryScreen.Bounds.Right > secondaryScreen.Bounds.Left || primaryScreen.Bounds.Left > secondaryScreen.Bounds.Right)
                        {
                            // Swap screens if they are inverted
                            secondaryScreen = screens[0];
                            primaryScreen = screens[1];
                        }
                        Screen currentScreen = Screen.FromPoint(Cursor.Position);
                        if (currentScreen.Primary)
                        {
                            screenBounds = primaryScreen.Bounds;                            
                        }
                        else
                        {
                            screenBounds = secondaryScreen.Bounds;
                        }

                        break;

                    case 1:  //one screen used
                        Screen primaryScreen_alone = screens[0];
                        Screen currentScreen_alone = Screen.FromPoint(Cursor.Position);
                        screenBounds = primaryScreen_alone.Bounds;
                        break;
                    default:
                        break;
                }              


                using (Bitmap bitmap = new Bitmap(screenBounds.Width, screenBounds.Height - 35))
                {
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(screenBounds.Location, System.Drawing.Point.Empty, screenBounds.Size);
                    }
                    // Save the captured screenshot to a file or perform any other operations
                    bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                }
            }

            #endregion 

            //********
            //Then show windows and plots for printing

            Window window = new Window();
            var mainViewModel = new MainViewModel(context, working_folder);
            var mainView = new MainView(mainViewModel, working_folder, filePath, backupPC_adress, files_to_delete);

            window.Title = "DVH  Plots ";
            window.Content = mainView;
            window.Height = 900;

            window.ShowDialog();

            //---------------------------------------------------------------------------------------------------------

        }

    }
}

