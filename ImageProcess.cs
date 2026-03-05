
using System;
using System.Drawing;
// To fix error : type bitmap could not be found...
// You have to add System.Drawing to your references so in solution Explorer rightclick on References
// and click on Add References and in assemblies find System.Drawing and click OK

namespace ImageProcess
{
    class Program
    {
        class ColorBarBinCount
        {
            public int White;
            public int Yellow;
            public int Cyan;
            public int Green;
            public int Magenta;
            public int Red;
            public int Blue;
            public int Black;
            public int Unknown;

            // Main constructor
            public ColorBarBinCount()
            {
                Yellow = 0;
                Cyan = 0;
                Green = 0;
                Magenta = 0;
                Red = 0;
                Blue = 0;
                Black = 0;
                Unknown = 0;
            }
        }
        class ColorBarBinRatio
        {
            public double White;
            public double Yellow;
            public double Cyan;
            public double Green;
            public double Magenta;
            public double Red;
            public double Blue;
            public double Black;
            public double Unknown;

            // Main constructor
            public ColorBarBinRatio()
            {
                Yellow = 0;
                Cyan = 0;
                Green = 0;
                Magenta = 0;
                Red = 0;
                Blue = 0;
                Black = 0;
                Unknown = 0;
            }
        }

        // Color bar detection function
        // Gray(White)  Yellow   Cyan  Green  Magenta  Red  Blue

        static float DetectWhiteBar(int R, int G, int B)
        {
            // White Bar computed statistic for 10 pixels RGB value on pattern captured with webcam
            //int Wbar_Rmean = (215 + 193 + 203 + 220 + 209 + 174 + 210 + 200 + 197 + 214) / 10;
            //int Wbar_Gmean = (220 + 203 + 215 + 221 + 217 + 180 + 223 + 201 + 200 + 223) / 10;
            //int Wbar_Bmean = (255 + 255 + 255 + 255 + 255 + 252 + 255 + 255 + 255 + 255) / 10;
            //int Wbar_Rmean = 203;
            //int Wbar_Gmean = 210;
            //int Wbar_Bmean = 254;
            int Wbar_Rmean = 118;
            int Wbar_Gmean = 108;
            int Wbar_Bmean = 198;
            // Compute Red Bar similarity
            int d1 = Math.Abs(R - Wbar_Rmean);
            int d2 = Math.Abs(G - Wbar_Gmean);
            int d3 = Math.Abs(B - Wbar_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float WhiteBarSimilarity = a * b * c;
            return WhiteBarSimilarity;
        }
        static float DetectYellowBar(int R, int G, int B)
        {
            // Yellow Bar computed statistic for 10 pixels RGB value on pattern captured with webcam
            //int Ybar_Rmean = (193 + 165 + 184 + 173 + 166 + 196 + 176 + 197 + 176 + 183) / 10;
            //int Ybar_Gmean = (209 + 180 + 196 + 189 + 171 + 203 + 192 + 214 + 192 + 199) / 10;
            //int Ybar_Bmean = (107 +  87 +  87 +  91 +  84 + 104 +  90 + 109 +  90 +  99) / 10;
            //int Ybar_Rmean = 180;
            //int Ybar_Gmean = 194;
            //int Ybar_Bmean = 94;
            int Ybar_Rmean = 170;
            int Ybar_Gmean = 154;
            int Ybar_Bmean = 24;
            // Compute Red Bar similarity
            int d1 = Math.Abs(R - Ybar_Rmean);
            int d2 = Math.Abs(G - Ybar_Gmean);
            int d3 = Math.Abs(B - Ybar_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float YellowBarSimilarity = a * b * c;
            return YellowBarSimilarity;
        }

        static float DetectCyanBar(int R, int G, int B)
        {
            // Cyan Bar computed statistic for 10 pixels RGB value on pattern captured with webcam
            //int Cbar_Rmean = ( 15 +  19 +  31 +  34 +  30 +  38 +  12 +  23 +  34 +  15) / 10;
            //int Cbar_Gmean = (148 + 188 + 218 + 235 + 219 + 227 + 150 + 179 + 219 + 179) / 10;
            //int Cbar_Bmean = (230 + 255 + 255 + 255 + 255 + 255 + 233 + 255 + 255 + 255) / 10;
            int Cbar_Rmean = 25;
            int Cbar_Gmean = 196;
            int Cbar_Bmean = 250;
            // Compute Red Bar similarity
            int d1 = Math.Abs(R - Cbar_Rmean);
            int d2 = Math.Abs(G - Cbar_Gmean);
            int d3 = Math.Abs(B - Cbar_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float CyanBarSimilarity = a * b * c;
            return CyanBarSimilarity;
        }
        static float DetectGreenBar(int R, int G, int B)
        {
            // Green Bar computed statistic for 10 pixels RGB value on pattern captured with webcam
            //int Gbar_Rmean = (71 + 39 + 50 + 102 + 103 + 87 + 37 + 118 + 44 + 101) / 10;
            //int Gbar_Gmean = (255 + 255 + 255 + 255 + 255 + 255 + 255 + 254 + 255 + 255) / 10;
            //int Gbar_Bmean = (128 + 85 + 95 + 177 + 183 + 165 + 79 + 197 + 89 + 182) / 10;
            //int Gbar_Rmean = 75;
            //int Gbar_Gmean = 254;
            //int Gbar_Bmean = 138;
            int Gbar_Rmean = 34;
            int Gbar_Gmean = 247;
            int Gbar_Bmean = 26;
            // Compute Green Bar similarity
            int d1 = Math.Abs(R - Gbar_Rmean);
            int d2 = Math.Abs(G - Gbar_Gmean);
            int d3 = Math.Abs(B - Gbar_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float GreenBarSimilarity = a * b * c;
            return GreenBarSimilarity;
        }
        static float DetectMagentaBar(int R, int G, int B)
        {
            // magenta Bar computed statistic for 10 pixels RGB value on pattern captured with webcam
            //int Mbar_Rmean = (255 + 211 + 255 + 255 + 255 + 255 + 255 + 255 + 255 + 255) / 10;
            //int Mbar_Gmean = ( 60 +  29 +  64 +  56 +  62 +  47 +  46 +  58 +  49 +  65) / 10;
            //int Mbar_Bmean = (255 + 254 + 255 + 255 + 255 + 255 + 255 + 255 + 255 + 255) / 10;
            //int Mbar_Rmean = 250;
            //int Mbar_Gmean = 53;
            //int Mbar_Bmean = 254;
            int Mbar_Rmean = 209;
            int Mbar_Gmean = 0;
            int Mbar_Bmean = 123;
            // Compute Blue Bar similarity
            int d1 = Math.Abs(R - Mbar_Rmean);
            int d2 = Math.Abs(G - Mbar_Gmean);
            int d3 = Math.Abs(B - Mbar_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float MagentaBarSimilarity = a * b * c;
            return MagentaBarSimilarity;
        }
        static float DetectRedBar(int R, int G, int B)
        {
            // Red Bar computed statistic for 10 pixels RGB value on pattern captured with webcam
            //int Rbar_Rmean = (255 + 255 + 255 + 255 + 255 + 255 + 255 + 255 + 255 + 255) / 10;
            //int Rbar_Gmean = (48 + 55 + 55 + 37 + 58 + 55 + 41 + 49 + 49 + 53) / 10;
            //int Rbar_Bmean = (75 + 76 + 77 + 62 + 76 + 73 + 69 + 65 + 77 + 83) / 10;
            //int Rbar_Rmean = 255;
            //int Rbar_Gmean = 50;
            //int Rbar_Bmean = 73;
            int Rbar_Rmean = 185;
            int Rbar_Gmean = 0;
            int Rbar_Bmean = 9;
            // Compute Red Bar similarity
            int d1 = Math.Abs(R - Rbar_Rmean);
            int d2 = Math.Abs(G - Rbar_Gmean);
            int d3 = Math.Abs(B - Rbar_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float RedBarSimilarity = a * b * c;
            return RedBarSimilarity;
        }
        static float DetectBlueBar(int R, int G, int B)
        {
            // Blue Bar computed statistic for 10 pixels RGB value on pattern captured with webcam
            //int Bbar_Rmean = (48 + 48 + 46 + 43 + 47 + 50 + 45 + 48 + 46 + 55) / 10;
            //int Bbar_Gmean = (76 + 79 + 74 + 65 + 74 + 83 + 67 + 79 + 68 + 88) / 10;
            //int Bbar_Bmean = (239 + 230 + 229 + 230 + 232 + 229 + 231 + 230 + 230 + 228) / 10;
            //int Bbar_Rmean = 47;
            //int Bbar_Gmean = 75;
            //int Bbar_Bmean = 230;
            int Bbar_Rmean = 4;
            int Bbar_Gmean = 0;
            int Bbar_Bmean = 187;
            // Compute Blue Bar similarity
            int d1 = Math.Abs(R - Bbar_Rmean);
            int d2 = Math.Abs(G - Bbar_Gmean);
            int d3 = Math.Abs(B - Bbar_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float BlueBarSimilarity = a * b * c;
            return BlueBarSimilarity;
        }
        static float DetectBlack(int R, int G, int B)
        {
            //  Black screen computed statistic for 10 pixels RGB value on webcam capture took inside area box
            //int Black_Rmean = (36 + 29 + 32 + 28 + 42 + 36 + 23 + 29 + 31 + 26) / 10;
            //int Black_Gmean = (38 + 34 + 32 + 30 + 43 + 36 + 28 + 31 + 33 + 31) / 10;
            //int Black_Bmean = (44 + 39 + 39 + 36 + 46 + 36 + 31 + 37 + 39 + 36) / 10;
            int Black_Rmean = 31;
            int Black_Gmean = 33;
            int Black_Bmean = 38;
            // Compute Blue Bar similarity
            int d1 = Math.Abs(R - Black_Rmean);
            int d2 = Math.Abs(G - Black_Gmean);
            int d3 = Math.Abs(B - Black_Bmean);

            float a = 1 - (float)d1 / 256;
            float b = 1 - (float)d2 / 256;
            float c = 1 - (float)d3 / 256;

            float BlackSimilarity = a * b * c;
            return BlackSimilarity;
        }

        static ColorBarBinRatio ColorBarBinCountRatio(ColorBarBinCount a)
        {
            //Gray(White)  Yellow  Cyan  Green  Magenta  Red  Blue  Black  Unknown
            int TotalPixel = a.White + a.Yellow + a.Cyan + a.Green + a.Magenta + a.Red + a.Blue + a.Black + a.Unknown;
            ColorBarBinRatio BinRatio; // The ratio between 0-1 for each color band
            BinRatio = new ColorBarBinRatio();
            BinRatio.Unknown = 1.0;
            
            if (TotalPixel != 0)
            {
                BinRatio.White = (double)a.White / TotalPixel;
                BinRatio.Yellow = (double)a.Yellow / TotalPixel;
                BinRatio.Cyan = (double)a.Cyan / TotalPixel;
                BinRatio.Green = (double)a.Green / TotalPixel;
                BinRatio.Magenta = (double)a.Magenta / TotalPixel;
                BinRatio.Red = (double)a.Red / TotalPixel;
                BinRatio.Blue = (double)a.Blue / TotalPixel;
                BinRatio.Black = (double)a.Black / TotalPixel;
                BinRatio.Unknown = (double)a.Unknown / TotalPixel;
            }
            return BinRatio;
        }

        static void Main(string[] aoArgs)
        {
            Console.WriteLine("ImageProcess");
            int NumberOfArea = 1;
            //string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\temp\mymbp.png";
            //string path = @"C:\Sylvain\TestImage\image_good.bmp";
            //string path = @"C:\HostRhesus\image_good.bmp";
            string oInpPath  = @"C:\HostRhesus\image.bmp";
            string oOutpPath = @"C:\HostRhesus\image_processed.bmp";
            string oSwitch   = string.Empty;
            
            foreach (string oItem in aoArgs)
            {
                if ((oItem.ToLower().Equals("-i"))||(oItem.Equals("--Input")))
                {
                    oSwitch = "-i";
                }
                else if (oItem.ToLower().Equals("-o") || (oItem.Equals("--Output")))
                {
                    oSwitch = "-o";
                }
                else if (oSwitch.Equals("-i"))
                {
                    oInpPath = oItem;
                    oSwitch  = string.Empty;
                }
                else if (oSwitch.Equals("-o"))
                {
                    oOutpPath = oItem;
                    oSwitch   = string.Empty;
                }
                else if (oItem.ToLower().Equals("help"))
                {
                    string oIndent = "    ";
                    Console.WriteLine("Usage:");
                    Console.WriteLine(oIndent + "ImageProcess [-i/--Input] inputFilePath [-o/--Output] outputFilePath");
                    Console.WriteLine(oIndent + "ImageProcess     (no parameters: use default input/output file path)");
                    Console.WriteLine(oIndent + oIndent + "Default: Input: " + oInpPath + " / Output: " + oOutpPath);
                    Console.WriteLine("Press CR to Exit...");
                    Console.Read();
                    return;
                }
            }

            Bitmap image = new Bitmap(oInpPath, true);
            for (int x = 0; x < image.Width; x++)
            {
                for (int y = 0; y < image.Height; y++)
                {
                    Color clr = image.GetPixel(x, y);
                    //Color newclr = Color.FromArgb(clr.R, 0, 0);
                    //Color newclr = Color.FromArgb(0, clr.G, 0);
                    //Color newclr = Color.FromArgb(0, 0, clr.B);
                    //Color newclr = Color.FromArgb(clr.R, 0, clr.B);
                    Color newclr = Color.FromArgb(clr.R, clr.G, clr.B);
                    image.SetPixel(x, y, newclr);
                }
            }

            // Definition of up to two area to study
            //int BoxWidth = 200;     // Box size is the same for all area
            //int BoxWidth = (int)(0.104166 * image.Width);
            int BoxWidth = (int)(0.95 * image.Width);

            //int BoxHeight = 200;
            //int BoxHeight = (int)(0.1851 * image.Height);
            int BoxHeight = (int)(0.5 * image.Height);

            // area position
            //int x1 = 790;           // top left coordinate
            //int y1 = 320;
            //int x2 = 1350;
            //int y2 = 350;

            //int x1 = (int)(0.4114 * image.Width);   // top left coordinate
            //int y1 = (int)(0.1666 * image.Height);
            //int x2 = (int)(0.7031 * image.Width);
            //int y2 = (int)(0.3241 * image.Height);

            int x1 = (int)(0 * image.Width);      // top left coordinate
            int y1 = (int)(0.2 * image.Height);
            int x2 = 0;
            int y2 = (int)(0.5 * image.Height);

            // Compute some RGB pattern detection

            //float RedBarSimilarity = DetectRedBar(255, 0, 0);     // 0.5753
            //float RedBarSimilarity = DetectRedBar(255, 48, 75);     // 0.9844
            //float GreenBarSimilarity = DetectGreenBar(255, 48, 75);     // 0.043
            //float BlueBarSimilarity = DetectBlueBar(255, 48, 75);       // 0.066

            //int SuccessDetectRate;
            int TotalNumberOfPixelInArea = BoxWidth * BoxHeight;

            // Define Bar Similarity Thresholds
            double ThresholdWhiteBarSimilarity = 0.5;
            double ThresholdYellowBarSimilarity = 0.5;
            double ThresholdCyanBarSimilarity = 0.5;
            double ThresholdRedBarSimilarity = 0.5;
            double ThresholdGreenBarSimilarity = 0.45;
            double ThresholdBlueBarSimilarity = 0.5;
            double ThresholdMagentaBarSimilarity = 0.5;

            //ColorBarBinCount[] colorBarBinCount;
            //colorBarBinCount = new ColorBarBinCount[2];

            ColorBarBinCount colorBarBinCount0;
            ColorBarBinCount colorBarBinCount1;
            colorBarBinCount0 = new ColorBarBinCount();
            colorBarBinCount1 = new ColorBarBinCount();

            // Highlight Detection
            // area 1
            int DetColBarPixCnt1 = 0;
            int DetBlackPixCnt1 = 0;
            int DetUnknownCnt1 = 0;
            for (int x = x1; x < (x1 + BoxWidth); x++)
            {
                for (int y = y1; y < (y1 + BoxHeight); y++)
                {
                    Color clr = image.GetPixel(x, y);
                    if (DetectWhiteBar(clr.R, clr.G, clr.B) > ThresholdWhiteBarSimilarity)
                    {
                        Color newclr = Color.FromArgb(255, 255, 255);   // Transform to full white
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.White += 1;
                        DetColBarPixCnt1 += 1;
                    }
                    else if (DetectYellowBar(clr.R, clr.G, clr.B) > ThresholdYellowBarSimilarity)
                    {
                        Color newclr = Color.FromArgb(255, 255, 0);   // Transform to full yellow
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.Yellow += 1;
                        DetColBarPixCnt1 += 1;
                    }
                    else if (DetectCyanBar(clr.R, clr.G, clr.B) > ThresholdCyanBarSimilarity)
                    {
                        Color newclr = Color.FromArgb(0, 255, 255);   // Transform to full cyan
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.Cyan += 1;
                        DetColBarPixCnt1 += 1;
                    }
                    else if (DetectRedBar(clr.R, clr.G, clr.B) > ThresholdRedBarSimilarity)
                    {
                        Color newclr = Color.FromArgb(255, 0, 0);   // Transform to full red
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.Red += 1;
                        DetColBarPixCnt1 += 1;
                    }
                    else if (DetectGreenBar(clr.R, clr.G, clr.B) > ThresholdGreenBarSimilarity)
                    {
                        Color newclr = Color.FromArgb(0, 255, 0);   // Transform to full green
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.Green += 1;
                        DetColBarPixCnt1 += 1;
                    }
                    else if (DetectBlueBar(clr.R, clr.G, clr.B) > ThresholdBlueBarSimilarity)
                    {
                        Color newclr = Color.FromArgb(0, 0, 255);   // Transform to full blue
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.Blue += 1;
                        DetColBarPixCnt1 += 1;
                    }
                    else if (DetectMagentaBar(clr.R, clr.G, clr.B) > ThresholdMagentaBarSimilarity)
                    {
                        Color newclr = Color.FromArgb(255, 0, 255);   // Transform to full magenta
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.Magenta += 1;
                        DetColBarPixCnt1 += 1;
                    }
                    else if (DetectBlack(clr.R, clr.G, clr.B) > 0.5)
                    {
                        Color newclr = Color.FromArgb(0, 0, 0);     // Transform to full black
                        image.SetPixel(x, y, newclr);
                        colorBarBinCount0.Black += 1;
                        DetBlackPixCnt1 += 1;
                    }
                    else
                    {
                        colorBarBinCount0.Unknown += 1;
                        DetUnknownCnt1 += 1;
                    }
                }
            }

            int DetColBarPixCnt2 = 0;
            int DetBlackPixCnt2 = 0;
            int DetUnknownCnt2 = 0;
            if (NumberOfArea == 2)
            {
                // Highlight Detection
                // area 2
                for (int x = x2; x < (x2 + BoxWidth); x++)
                {
                    for (int y = y2; y < (y2 + BoxHeight); y++)
                    {
                        Color clr = image.GetPixel(x, y);
                        if (DetectWhiteBar(clr.R, clr.G, clr.B) > ThresholdWhiteBarSimilarity)
                        {
                            Color newclr = Color.FromArgb(255, 255, 255);   // Transform to full white
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.White += 1;
                            DetColBarPixCnt2 += 1;
                        }
                        else if (DetectYellowBar(clr.R, clr.G, clr.B) > ThresholdYellowBarSimilarity)
                        {
                            Color newclr = Color.FromArgb(255, 255, 0);   // Transform to full yellow
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.Yellow += 1;
                            DetColBarPixCnt2 += 1;
                        }
                        else if (DetectCyanBar(clr.R, clr.G, clr.B) > ThresholdCyanBarSimilarity)
                        {
                            Color newclr = Color.FromArgb(0, 255, 255);   // Transform to full cyan
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.Cyan += 1;
                            DetColBarPixCnt2 += 1;
                        }
                        else if (DetectRedBar(clr.R, clr.G, clr.B) > ThresholdRedBarSimilarity)
                        {
                            Color newclr = Color.FromArgb(255, 0, 0);   // Transform to full red
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.Red += 1;
                            DetColBarPixCnt2 += 1;
                        }
                        else if (DetectGreenBar(clr.R, clr.G, clr.B) > ThresholdGreenBarSimilarity)
                        {
                            Color newclr = Color.FromArgb(0, 255, 0);   // Transform to full green
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.Green += 1;
                            DetColBarPixCnt2 += 1;
                        }
                        else if (DetectBlueBar(clr.R, clr.G, clr.B) > ThresholdBlueBarSimilarity)
                        {
                            Color newclr = Color.FromArgb(0, 0, 255);   // Transform to full blue
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.Blue += 1;
                            DetColBarPixCnt2 += 1;
                        }
                        else if (DetectMagentaBar(clr.R, clr.G, clr.B) > ThresholdMagentaBarSimilarity)
                        {
                            Color newclr = Color.FromArgb(255, 0, 255);   // Transform to full magenta
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.Magenta += 1;
                            DetColBarPixCnt2 += 1;
                        }
                        else if (DetectBlack(clr.R, clr.G, clr.B) > 0.5)
                        {
                            Color newclr = Color.FromArgb(0, 0, 0);     // Transform to full black
                            image.SetPixel(x, y, newclr);
                            colorBarBinCount1.Black += 1;
                            DetBlackPixCnt2 += 1;
                        }
                        else
                        {
                            colorBarBinCount1.Unknown += 1;
                            DetUnknownCnt2 += 1;
                        }
                    }
                }
            }

            ColorBarBinRatio BinRatio; // The ratio between 0-1 for each color band
            BinRatio = new ColorBarBinRatio();
            BinRatio = ColorBarBinCountRatio(colorBarBinCount0);

            // Compute percentage of pixel identified as Colorbar
            float DetectPattern1 = 0;
            float DetectPattern2 = 0;
            float DetectBlack1 = 0;
            float DetectBlack2 = 0;
            float DetectUnknown1 = 100;
            float DetectUnknown2 = 100;
            if (TotalNumberOfPixelInArea != 0) // DIV0 protection
            {
                DetectPattern1 = (float)100 * DetColBarPixCnt1 / (float)TotalNumberOfPixelInArea;
                DetectPattern2 = (float)100 * DetColBarPixCnt2 / (float)TotalNumberOfPixelInArea;
                DetectBlack1 = (float)100 * DetBlackPixCnt1 / (float)TotalNumberOfPixelInArea;
                DetectBlack2 = (float)100 * DetBlackPixCnt2 / (float)TotalNumberOfPixelInArea;
                DetectUnknown1 = (float)100 * DetUnknownCnt1 / (float)TotalNumberOfPixelInArea;
                DetectUnknown2 = (float)100 * DetUnknownCnt2 / (float)TotalNumberOfPixelInArea;
            }
            Console.WriteLine(string.Format("ColorBar {0,3:0} {1,3:0}", DetectPattern1, DetectPattern2));
            Console.WriteLine(string.Format("Black    {0,3:0} {1,3:0}", DetectBlack1, DetectBlack2));
            Console.WriteLine(string.Format("Unknown  {0,3:0} {1,3:0}", DetectUnknown1, DetectUnknown2));

            // Show area delimitation
            // Do it after the pixel measurement because it alter the image buffer...
            Color delimiterclr = Color.FromArgb(255, 0, 0);
            // area1
            for (int x = x1; x < (x1 + BoxWidth); x++)
            {
                //Color clr = image.GetPixel(x, y1);
                //Color newclr = Color.FromArgb(255-(clr.R), 255-clr.G, 255-clr.B);
                image.SetPixel(x, y1, delimiterclr);        // top line
                //clr = image.GetPixel(x, y1+h1);
                //newclr = Color.FromArgb(255 - (clr.R), 255 - clr.G, 255 - clr.B);
                image.SetPixel(x, y1 + BoxHeight, delimiterclr);     // bottom line
            }
            for (int y = y1; y < (y1 + BoxHeight); y++)
            {
                image.SetPixel(x1, y, delimiterclr);        // left line
                image.SetPixel(x1 + BoxWidth, y, delimiterclr);     // right line
            }

            if (NumberOfArea == 2)
            {
                // area2
                for (int x = x2; x < (x2 + BoxWidth); x++)
                {
                    image.SetPixel(x, y2, delimiterclr);        // top line
                    image.SetPixel(x, y2 + BoxHeight, delimiterclr);   // bottom line
                }
                for (int y = y2; y < (y2 + BoxHeight); y++)
                {
                    image.SetPixel(x2, y, delimiterclr);        // left line
                    image.SetPixel(x2 + BoxWidth, y, delimiterclr);   // right line
                }
            }

            /*
            // Copy all pixels
            for (int x = x1; x < (x1+w1); x++)
            {
                for (int y = y1; y < (y1+h1); y++)
                {
                    Color clr = image.GetPixel(x, y);
                    //Color newclr = Color.FromArgb(clr.R, 0, 0);
                    //Color newclr = Color.FromArgb(0, clr.G, 0);
                    //Color newclr = Color.FromArgb(0, 0, clr.B);
                    //Color newclr = Color.FromArgb(clr.R, 0, clr.B);
                    Color newclr = Color.FromArgb(clr.R, clr.G, clr.B);
                    image.SetPixel(x, y, newclr);
                }
            }
            */

            // Test pattern generator
            // Toggle between R,G,B
            // ex for a band width of 100
            // 0-99     Red
            // 100-199  Green
            // 200-299  Blue
            /*
            int BandWidth = 100;    // Width of a color bar in pixels
            for (int x = 0; x < image.Width; x++)
            {
                int a = x % (3 * BandWidth);    // take remainder and provide each band a color
                Color PatterColor;
                if (a<1*BandWidth)
                    PatterColor = Color.FromArgb(255, 0, 0);
                else if (a < (2 * BandWidth))
                    PatterColor = Color.FromArgb(0, 255, 0);
                else
                    PatterColor = Color.FromArgb(0, 0, 255);

                for (int y = 0; y < image.Height; y++)
                {
                    image.SetPixel(x, y, PatterColor);
                }
            }
            */

            //image.Save(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\temp\newbmp.png");
            //image.Save(@"C:\Sylvain\TestImage\image_good_processed.bmp");
            image.Save(@"C:\HostRhesus\image_processed.bmp");
            Console.WriteLine("process is done!");
            //Console.ReadKey();
        }
    }
}
