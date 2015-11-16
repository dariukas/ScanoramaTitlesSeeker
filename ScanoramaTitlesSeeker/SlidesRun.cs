using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.IO;
using System.Globalization;
using System.Collections;

namespace ScanoramaTitlesSeeker
{
    class SlidesRun
    {
        static List<KeyValuePair<string, string>> initialList = new List<KeyValuePair<string, string>>();
        static List<KeyValuePair<string, float>> list = new List<KeyValuePair<string, float>>();


        //MANIPULATIONS WITH LISTS
        public static void addToList(String text, float timecode) {
            list.Add(new KeyValuePair<string, float>(text.Trim(), timecode));
        }

        public static void addToInitialList(String timestamp, String text)
        {
            initialList.Add(new KeyValuePair<string, string>(timestamp.Trim(), text.Trim()));
        }

        public static void initialListToList(List<KeyValuePair<string, string>> inLs) {
            float key1 = 0;
            //double key1 = 0;
            String value = "";
            foreach (var pair in inLs)
            {

                //double key2 = timestampToDoubleSeconds(pair.Key.Trim());
                //float key = Convert.ToSingle(key2 - key1);
                float key2 = timestampToFloatSeconds(pair.Key.Trim());
                float key = key2 - key1;

                addToList(value, key);
                //System.Console.WriteLine(key + " - " + value);
                key1 = key2;
                value = pair.Value;
              
            }
            //return list;
        }





        public static Application createApplication() {
            Microsoft.Office.Interop.PowerPoint.Application oPowerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
            oPowerPoint.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            System.Console.WriteLine("PowerPoint application created.");
            return oPowerPoint;
        }

        public static Presentation openPresentation(Application oPowerPoint)
        {
            Presentations oPres = oPowerPoint.Presentations;
            Presentation oPre = oPres.Open(FilesController.openFile());
            System.Console.WriteLine("PowerPoint application opened.");
            return oPre;
        }

        public static Presentation startShow (Presentation oPre) {
            oPre.SlideShowSettings.ShowPresenterView = MsoTriState.msoTrue;
            //oPre.SlideShowSettings.ShowType
            oPre.SlideShowSettings.Run();
            System.Console.WriteLine("PowerPoint presentation {0} is running.", oPre.Name);
            return oPre;
        }

        public static float readSlideTimecode(SlideShowWindow Wn) {
            String timeStamp = GetTimestamp(DateTime.Now);
            float timecode = Wn.View.SlideElapsedTime;
            return timecode;
        }

        public static String readSlideText(SlideShowWindow Wn)
        {
            SlideShowView ssv = Wn.View;
            Slide currentSlide = ssv.Slide;
            if (currentSlide.Shapes.Count==0)
            {
                System.Console.WriteLine("Error");
                return "";
            }
            else {
                TextRange tr = currentSlide.Shapes[1].TextFrame.TextRange;
                return italicFilter(tr);
        }
        }

        public static String italicFilter(TextRange tr) {
            if (tr.Font.Italic == MsoTriState.msoTrue) {
                return "<i>" + tr.Text + "</i>";
            }
            return tr.Text;
        }

        public static Slide readSlide(SlideShowWindow Wn)
        {
            //SlideShowView ssv  = oPre.SlideShowWindow.View;
            SlideShowView ssv = Wn.View;
            return ssv.Slide;
        }

 
        public static double timestampToDoubleSeconds(String timestamp)
        {
            TimeSpan span = TimeSpan.ParseExact(timestamp, "hh\\:mm\\:ss\\:ffff", CultureInfo.InvariantCulture);
            return span.TotalSeconds;
        }

        public static float timestampToFloatSeconds(String timestamp)
        {
            try {
                //printChars(timestamp);
                TimeSpan span = TimeSpan.ParseExact(timestamp, "hh\\:mm\\:ss\\:ffff", CultureInfo.InvariantCulture);
                double seconds = span.TotalSeconds;
                //System.Console.WriteLine(timestamp + " to " + Convert.ToSingle(seconds));
                return Convert.ToSingle(seconds);
            }
            catch (Exception e) {
                //System.Console.WriteLine(e);
                //printChars(timestamp);
                System.Environment.Exit(1);
            }
            return 0;
        }

        /* public static void addTimecodes(Presentation oPre)
         {
             Slides oSlides = oPre.Slides;
         }*/

  

        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("hh:mm:ss:ffff");
        }

        //SYNCHRONICALLY CREATE SLIDE (NOT WORKING)

        public static Presentation createAdditionalPresentation(Application oPowerPoint)
        {
            Presentations oPres = oPowerPoint.Presentations;
            //System.Console.WriteLine(oPres.Count);
            Presentation oPreCopy = oPres.Add(MsoTriState.msoFalse);
            //System.Console.WriteLine(oPres.Count);
            oPreCopy = oPres[2];
            return oPreCopy;
        }

        //FILE TO SLIDES
        public static List<KeyValuePair<string, float>> fileToList(String filePath)
        {

            System.Console.WriteLine("Reading the titles from the file {0}...", filePath);
            string[] lines = System.IO.File.ReadAllLines(filePath);
            //exception for double lines
            /*ArrayList array1 = new ArrayList();
            for (int i = 0; i<lines.Length; i=i+4) {
                array1.Add(lines[i]);
                array1.Add(lines[i+1]);
            }*/
            
            bool sw = true;
            String timestamp = "";
            //put file lines into the dictionary
            foreach (string line in lines)
            {
                //System.Console.WriteLine(line);
                //printChars(line);
                if (sw == true)
                {
                    timestamp = line;
                    sw = false;
                }
                else
                {
                    addToInitialList(timestamp, line);
                    sw = true;
                }
                //if (Regex.IsMatch(line, @"[0-9]>[0-9]"))
            }
            initialListToList(initialList);
            return list;
        }

        //ADDITIONAL METHODS
        public static void printList(List<KeyValuePair<string, float>> ls)
        {
            foreach (var pair in ls)
            {
                System.Console.WriteLine(pair.Value);
                System.Console.WriteLine(pair.Key);
            }
        }

        public static void printChars(string text)
        {
            char[] myChars = text.ToCharArray();
            foreach (char ch in myChars)
            {
                System.Console.Write(ch + @" - \u" + ((int)ch).ToString("X4") + ", ");
            }
            System.Console.WriteLine();
        }

        //Events
        public static void app_SlideShowNextSlide(SlideShowWindow Wn)
        {
            //System.Console.WriteLine(readSlideTimecode(Wn));
            //System.Console.WriteLine(readSlideText(Wn));
            StreamWriter sw = new StreamWriter(MainWindow.savePath, true);
            sw.WriteLine(GetTimestamp(DateTime.Now));
           // sw.WriteLine(readSlideTimecode(Wn));
            sw.WriteLine(readSlideText(Wn));
            sw.Close();
            addToInitialList(GetTimestamp(DateTime.Now), readSlideText(Wn));
            //addToList(readSlideText(Wn), readSlideTimecode(Wn));
            //crashing
           //addTimecodeToSlide(readSlide(Wn), readSlideTimecode(Wn));
        }

        public static void app_SlideShowBegin(SlideShowWindow Wn)
        {
            System.Console.WriteLine("Slide Show begins");
        }

        public static void app_SlideShowEnd(Presentation oPre)
        {
            //NewSlides.createPresentation(list);
            initialListToList(initialList);
            //printList(list);
            NewSlides.createPresentation(list, oPre.Name.Split('.')[0]+"Mod");
            System.Console.WriteLine("Slide Show ended.");
        }

        public static void app_PresentationClose(Presentation oPre)
        {
            //FilesController.saveSlides(oPre);
            //NewSlides.createPresentation(list);
            System.Console.WriteLine("Presentation closed.");
        }

    }
}
