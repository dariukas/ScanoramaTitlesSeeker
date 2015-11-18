using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

using Microsoft.Office.Interop.PowerPoint;


namespace ScanoramaTitlesSeeker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();
        public static string savePath = FilesController.folderPath + "\\text.txt";

        public MainWindow()
        {
            InitializeComponent();
            InitPPT();
            run();
        }

        public static void run()
        {
            FilesController.createDeaultFolder("Scano");
            Presentation oPre = SlidesRun.openPresentation(app);
            System.Console.WriteLine(oPre.Name);
            savePath = FilesController.saveTxtFile(oPre.Name.Split('.')[0]); //split is necessary to cut .pptx
            SlidesRun.startShow(oPre);
            //Presentation newoPre = SlidesRun.createAdditionalPresentation(app);
            //FilesController.saveSlides(newoPre);
        }

        public static void run2()
        {
            FilesController.createDeaultFolder("SelfTitling");
            String txtFilePath = FilesController.openTxtFile();
            FilesController.saveSlides(NewSlides.createPresentation2(SlidesRun.fileToList(txtFilePath)));
        }

        //private UCOMIConnectionPoint m_oConnectionPoint;
        //private int m_Cookie;
        //private Microsoft.Office.Interop.PowerPoint.ApplicationClass oPPT;

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            /*
            // QI for IConnectionPointContainer.
            UCOMIConnectionPointContainer oConnPointContainer = (UCOMIConnectionPointContainer)app;
            // Get the GUID of the EApplication interface.
            Guid guid = typeof(Microsoft.Office.Interop.PowerPoint.EApplication).GUID;

            // Find the connection point.
            oConnPointContainer.FindConnectionPoint(ref guid, out m_oConnectionPoint);
            // Call Advise to sink up the connection.
            m_oConnectionPoint.Advise(this, out m_Cookie);*/
        }

        //Registering events
        private void InitPPT()
        {
            app.SlideShowNextSlide += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowNextSlideEventHandler(SlidesRun.app_SlideShowNextSlide);
            app.SlideShowBegin += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowBeginEventHandler(SlidesRun.app_SlideShowBegin);
            app.SlideShowEnd += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowEndEventHandler(SlidesRun.app_SlideShowEnd);
            app.PresentationClose += new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationCloseEventHandler(SlidesRun.app_PresentationClose);
            //app.SlideShowEnd += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowEndEventHandler(SlidesRun.app_PresentationClose);
        }
    }
}
