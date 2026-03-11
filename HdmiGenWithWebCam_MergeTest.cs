// --------------------------------------------------------------------------------------------------------------------------------
// (c) Copyright Matrox Graphics Inc.                                            __  __   ____  ___
//                                                                              |  \/  | / ___||_ _|
// This documents contains confidential proprietary information                 | |\/| || |  _  | |
// that may not be disclosed without prior written permission                   | |  | || |_| | | |
// from Matrox Graphics Inc.                                                    |_|  |_| \____||___|
//
// $Id: HdmiGenWithWebCam.cs 113977 2023-09-07 19:00:26Z jpaquett $
// $HeadURL: http://rex/mgi_svn/Projects/HdmiGen/Trunk/Script/HdmiGenWithWebCam.cs $
// $Revision: 113977 $
// $Date: 2023-09-07 15:00:26 -0400 (Thu, 07 Sep 2023) $
// Project: M7100 HDMI input certification
// Documentation:
// Description: This script control the Digilent Nexys Video board to produce HDMI output with some transition defects.
//              It control a webcam to capture the monitor output. It attempt to detect the test pattern and count black screen.
// Version: 31 May 2023
// --------------------------------------------------------------------------------------------------------------------------------
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;
using ScriptInterpreter;
using System.Text.RegularExpressions;
using System.IO.Ports;
using System.Diagnostics; // Process.Start Method

public class Script
{
    // ======================================================================
    // SVN Variables (do not change)
    // ======================================================================
    string strSvnRevision = "$Revision: 113977 $";
    string strSvnDate = "$Date: 2023-09-07 15:00:26 -0400 (Thu, 07 Sep 2023) $";

    // ======================================================================
    // Enum for next value calculation algorithm
    // ======================================================================
    enum EChangeAlgo
    {
        INCREMENTAL = 0, // Continuous incremental sweep of every values
        RANDOM      = 1, // Random resolution change
        STATIC      = 2, // Keep same initial value
        COUNT       = 3, // Number of values
    }

    // ======================================================================
    // Enum for resolutions
    // ======================================================================
    enum ESourceRes
    {
        DVI_720x480P60_27027_RGB_8          = 0,  //  27.027 MHz, 480p_8_60Hz_CBar_RGB_PC_v53_Dvi
        HDMI_640x480P60_25200_RGB_8         = 1,  //  25.200 MHz, VGA_8_60Hz_CBar_RGB_TV_v54
        HDMI_720x480P60_27027_RGB_8         = 2,  //  27.027 MHz, 480p_8_60Hz_CBar_RGB_TV_v54
        HDMI_720x576P50_27000_RGB_8         = 3,  //  27.000 MHz, 576p_8_50Hz_CBar_RGB_TV_v54
        HDMI_3840x2160P60_594000_YUV420_8   = 4,  // 594.000 MHz, 3840x2160p_60Hz_ColorBar_420_VIC_97
        HDMI_1280x720P60_74250_RGB_8        = 5,  //  74.250 MHz, 720p_8_60Hz_CBar_RGB_TV_v54
        HDMI_1280x720P50_92813_RGB_10       = 6,  //  92.813 MHz, 720p_10_50_CBar_RGB_TV_v90
        HDMI_1920x1080I50_74250_RGB_8       = 7,  //  74.250 MHz, 1080i_8_50Hz_CBar_RGB_TV_v54
        HDMI_1920x1080I60_74250_RGB_8       = 8,  //  74.250 MHz, 1080i_8_60Hz_CBar_RGB_TV_v54
        HDMI_1920x1080P30_74250_RGB_8       = 9,  //  74.250 MHz, 1080p_8_30Hz_CBar_RGB_TV_v99
        HDMI_1920x1080P50_148500_RGB_8      = 10, // 148.500 MHz, 1080p_8_50Hz_CBar_RGB_TV_v63
        HDMI_1920x1080P60_148500_RGB_8      = 11, // 148.500 MHz, 1080p_8_60Hz_CBar_RGB_TV_v54
        HDMI_1920x1080P60_185625_RGB_10     = 12, // 185.625 MHz, 1080p_10_60_CBar_RGB_TV_v90
        HDMI_1920x1080P60_222750_RGB_12     = 13, // 222.750 MHz, 1080p_12_60_CBar_RGB_TV_v90
        HDMI_3840x2160P24_297000_RGB_8      = 14, // 297.000 MHz, 4K2K_2160p_24Hz
        COUNT                               = 15, // Number of values
    }

    // ======================================================================
    // Enum for defects
    // ======================================================================
    enum EDefectType
    {
        CLEAN                   = 0, // Clean transition
        CLOCK_JITTER            = 1, // Clock Jitter 0.3 Tbit
        SYMBOL_SKEW_DELAY       = 2, // InterPair symbol skew random delay
        SEQUENTIAL_DATA         = 3, // Sequential Data (Lanes are enabled one after the other in random sequence with random delay)
        TMDS_OUTPUT_ENABLE      = 4, // Glitchy Output Enable on TMDS redriver  (Multiple random toggling)
        TMDS_SCRAMBLING         = 5, // TMDS scrambling. Done a random number of time.
        RANDOM_5V_DISCONNECT    = 6, // 5V random disconnect. Turn OFF-ON the +5V on HDMI a random number of time.
        TMDS_SWAP               = 7, // Swap TMDS Pair randomly
        ALL_DRESS               = 8, // Put it all
        BLACK                   = 9, // Black screen : used to verify capture is still alive
        COUNT                   = 10, // Number of values
    }

    // ======================================================================
    // User Variables (Start)
    // ======================================================================
    string  strHdmiSourceComPort    = "COM3";               // Nexsys HDMI source com port
    bool    bPing                   = false;                // Ping the DUT for response on each test iteration?
    string  IP_Address              = "192.168.182.218";    // Ping IP addresscmdcmd
    int     iStopAfterErrorCount    = 0;                    // Stop on N errors (0 to disable).
    int     iJitterAmplitude_mTbit  = 300;                  // Jitter amplitude in mTbit (0.300Tbit = 300)
    bool    bTransitionalJitter     = true;                 // Fixed transitional jitter time
    int     iJITTER_TIME_MS         = 500;                  // How much time the jitter is applyed
    bool    bTransitionalInterPairDelay = true              // Fixed transitional inter-pair delay
    int     iINTER_PAIR_DELAY_TIME_MS = 500;                // How much time the inter-pair delay is applyed

    // Source resolution selection
    ESourceRes  eSourceRes          = ESourceRes.HDMI_1920x1080P60_148500_RGB_8;    // Source initial value
    EChangeAlgo eSourceChangeAlgo   = EChangeAlgo.STATIC;                           // Source change algo
    int         iSourceChangeMask   = 0xFFFF;                                       // Source change mask enable (1 bit per source)

    // Defects selection
    EDefectType eDefectType         = EDefectType.CLEAN;                            // Defect Initial value
    EChangeAlgo eDefectChangeAlgo   = EChangeAlgo.INCREMENTAL;                      // Defect change algo
    int         iDefectChangeMask   = 0x0002;                                       // Defect change mask enable (1 bit per defect)
    int         iDefectImageMask    = 0xFF7D;                                       // Defect image verification mask enable (1 bit per defect).

    // LOOP_DELAY_MS holds the delay in ms before taking a webcam snapshot
    const int LOOP_DELAY_MS = 5000;
    // ======================================================================
    // User Variables (End)
    // ======================================================================

    // ======================================================================
    // Test variables
    // ======================================================================
    Thread m_oThread = null;
    dynamic m_oComTool = null;
    dynamic m_oNetwork = null;
    dynamic m_oResult = null;
    IScriptOptions m_oHostOptions = null;
    dynamic m_oSrcHandle = null;
    dynamic m_oAnaHandle = null;
    dynamic m_oIoHandle = null;

    bool bComPortTest = false;
    bool bIoControl = false;

    // string strComPortName_IoControl  = "COM87";

    const int TMDSOnTimeMsOffset = 1416; // Base TMDS active time in ms caused by the whole code between deactivation (measured with a scope)

    // Added delay in increment of 100ms
    const int RetryQtyMax = 20;

    // The variable bPromptForLogfile can be set to false to prevent the user
    // from  being prompted for the output filename. In this case, the output
    // data will be logged to the file given by the variable "LogFilename" and
    // placed on the Desktop
    const bool bPromptForLogfile = false;
    string DefaultLogFilename = @"\HDMILog.log";

    // string LogFilename = null;
    string LogFilename = @"\HDMILog.log";

    bool bTerminateRequest = false;
    bool bTerminateComplete = false;
    int iMonitorErrCount = 0;
    bool bRetryAfterDelay = true;
    int[] RetryQty = new int[RetryQtyMax + 1]; // Count Retry
    int RetryCouldNotFixCount = 0;
    bool bHaltOnMonitorNotDetect = false; // Allow to halt on monitor not detected after RetryQtyMax

    // Comm Error Stat
    int iCommErr1 = 0;
    int iCommErr2 = 0;
    int iCommErr3 = 0;

    // Mutex
    object m_oSync = new object();

    // =====================================================================
    // =====                                                          ======
    // =====           END OF PARAMETER DEFINITION BLOCK              ======
    // =====                                                          ======
    // =====================================================================
    int iLoopIdx = 0; // When test is interupted, will indicate qty of tested resolutions.
    string strErrorList = null;
    List<string> strMemInfo = new List<string>();
    List<string> strAVSLUsage = new List<string>();

    // ======================================================================
    // HDMI_Source class
    // ======================================================================
    public class HDMI_Source // Description of the parameters of the sources in the NEXYS FPGA
    {
        public int ResX;
        public int ResY;
        public int RefreshRate; // In KHz
        public bool IsInterlaced;
        public int PixelClock;
        public int Hfront;
        public int Hsync;
        public int Hback;
        public bool HpolPos;
        public int Vfront;
        public int Vsync;
        public int Vback;
        public bool VpolPos;
        public int VIC;
        public int DC;
        public int Y1Y0;
        public bool IsHdmi;

        public HDMI_Source(
            int X,
            int Y,
            int Rate,
            bool Inter,
            int PClk,
            int HfrontPorch,
            int Hsync_s,
            int HbackPorch,
            bool HpolarityPositive,
            int VfrontPorch,
            int Vsync_s,
            int VbackPorch,
            bool VpolarityPositive,
            int VICcode,
            int DeepCol,
            int Y1Y0_s,
            bool IsHdmi_s
        )
        {
            ResX = X;
            ResY = Y;
            RefreshRate = Rate;
            IsInterlaced = Inter;
            PixelClock = PClk;
            Hfront = HfrontPorch;
            Hsync = Hsync_s;
            Hback = HbackPorch;
            HpolPos = HpolarityPositive;
            Vfront = VfrontPorch;
            Vsync = Vsync_s;
            Vback = VbackPorch;
            VpolPos = VpolarityPositive;
            VIC = VICcode;
            DC = DeepCol;
            Y1Y0 = Y1Y0_s;
            IsHdmi = IsHdmi_s;
        }
    }

    public HDMI_Source[] HDMI_source;

    /************************************************************************************************************\
    Function:       GenerateReport
    Description:    Generate reports.
    \************************************************************************************************************/
    public string GenerateReport()
    {
        // =====================================================
        // Display final statistics
        // =====================================================
        string strReport = null;
        strReport += "\n";
        strReport += "Iterations,                                    TOTAL COUNT = " + iLoopIdx + "\n";
        float x;
        x = 100*iMonitorErrCount/(float)iLoopIdx;
        strReport += "Monitor Fail                                   ERROR COUNT = " + iMonitorErrCount + "    ERROR RATE = " + x.ToString("#0.00") + " %" + "\n";
        strReport += "Com Error1, Read command echo diffent than 'r'       COUNT = " + iCommErr1 + "\n";
        strReport += "Com Error2, Not enought character after 'r' command  COUNT = " + iCommErr2 + "\n";
        strReport += "Com Error3, Fail on WriteByte                        COUNT = " + iCommErr3 + "\n";
        for (int RetryCnt = 1 ; RetryCnt <= RetryQtyMax ; RetryCnt++)
        {
            strReport += "Retry " + RetryCnt + "   COUNT = " + RetryQty[RetryCnt] + "\n";
        }
        strReport += "Retry Could Not Fix  COUNT = " + RetryCouldNotFixCount + "\n";
        return strReport;
    }

    /************************************************************************************************************\
    Function:       Abort
    Description:    Abort the test.
    \************************************************************************************************************/
    public void Abort()
    {
        lock (m_oSync)
        {
            bTerminateRequest = true;
        }

        // bool bExit = false;
        // do
        // {
        //   lock(m_oSync)
        //   {
        //      bExit = bTerminateComplete;
        //   }
        //   Thread.Sleep(100);
        // }
        // while (!bExit);

        m_oThread.Abort();

        if (!string.IsNullOrEmpty(m_oMsgPostedByThread))
            m_oHostOptions.AppendToDebug(m_oMsgPostedByThread);

        m_oHostOptions.Result = "Aborted";
        m_oHostOptions.AppendResultToHistory();

        string strReport = "";
        strReport += "Test complete (user-terminated)";
        strReport += GenerateReport();
        m_oHostOptions.AppendToDebug(strReport, false);

        m_oComTool.CloseSession(m_oSrcHandle);
        m_oComTool.CloseSession(m_oAnaHandle);
        m_oComTool.CloseSession(m_oIoHandle);
    }

    private string m_oMsgPostedByThread = null;

    /************************************************************************************************************\
    Function:       Finished
    Description:    Finished function (required).
    \************************************************************************************************************/
    public bool Finished()
    {
        return false;
    }
    
    /************************************************************************************************************\
    Function:       Run
    Description:    Run the test thread.
    \************************************************************************************************************/
    public void Run(IScriptOptions oHostOptions)
    {
        // MessageBox.Show("Should have seen a checkedListBox by now...");
        m_oHostOptions = oHostOptions;
        m_oNetwork = m_oHostOptions.GetPlugInObject("RhesusNetwork");
        m_oComTool = m_oHostOptions.GetPlugInObject("ComTool");

        // MessageBox.Show("Got CommTool");
        // m_oSrcHandle = m_oComTool.OpenSession(strHdmiSourceComPort);
        // m_oSrcHandle = m_oComTool.OpenSession("COM5",115200,8, StopBits.Two, Parity.None,Handshake.None, true);
        m_oSrcHandle = m_oComTool.OpenSession(
            strHdmiSourceComPort,
            115200,
            8,
            StopBits.Two,
            Parity.None,
            Handshake.None,
            true);
        // MessageBox.Show("Got handle");
        if (m_oSrcHandle == null)
            MessageBox.Show("ERROR, Source ComTool port Handle invalid!...");
        // if( bIoControl == true )
        // {
        //   m_oIoHandle = m_oComTool.OpenSession(strComPortName_IoControl,460800,8, StopBits.Two, Parity.None,Handshake.None, true);
        //   if (m_oIoHandle == null)
        //      MessageBox.Show("ERROR, IoControl ComTool port Handle invalid!...");
        // }
        m_oHostOptions.SetDebugFont("Consolas", 10.0f);
        m_oComTool.SetXChgBufferMode(true, false, false); // (Boolean bAscii, Boolean bShowTxData, Boolean bShowRxData);

        CheckedListBox cbList = new CheckedListBox();
        cbList.Height = 600;
        cbList.Width = 800;
        cbList.Items.Add("Test 1 ; Some comment", false);
        cbList.Items.Add("Test 2 ; Some other comment", false);
        cbList.Show();

        // MessageBox.Show("Should have seen a checkedListBox by now...");
        if (
            m_oComTool.IsSessionOpenned(m_oSrcHandle)
            && m_oComTool.IsSessionOpenned(m_oIoHandle)
            && (m_oNetwork != null) )
        {
           m_oThread = new Thread(MainProcess);
           m_oThread.Start();
        }
        else
            MessageBox.Show("Not started: missing Networt PlugIn or ComTool session not openned.");
    }

    /************************************************************************************************************\
    Function:       GetFileToOpen
    Description:    Get file to open.
    \************************************************************************************************************/
    private string GetFileToOpen()
    {
        OpenFileDialog openFileDialog1 = new OpenFileDialog();
        string CfgFilename = null;
        openFileDialog1.InitialDirectory = @".\";

        openFileDialog1.RestoreDirectory = true;
        openFileDialog1.Title = "Browse Text Files";
        openFileDialog1.DefaultExt = "txt";
        openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
        openFileDialog1.FilterIndex = 2;
        openFileDialog1.CheckFileExists = false;
        openFileDialog1.CheckPathExists = true;

        DialogResult oResult = DialogResult.Cancel;
        var oThread = new Thread(
            new ParameterizedThreadStart(param =>
            {
                oResult = openFileDialog1.ShowDialog();
            }));
        oThread.SetApartmentState(ApartmentState.STA);

        oThread.Start();

        while (oThread.IsAlive)
            Thread.Sleep(100);

        if (oResult == DialogResult.OK)
        {
            CfgFilename = openFileDialog1.FileName;
        }

        return CfgFilename;
    }

    /************************************************************************************************************\
    Function:       ERROR
    Description:    Add error to list.
    \************************************************************************************************************/
    private void ERROR(string strError)
    {
        if (strErrorList.Contains("No Errors"))
        {
            strErrorList = "Errors:\r\n";
        }
        strErrorList += "     " + DateTime.Now.ToString("MM/dd/yy HH:mm:ss") + " ==> " + strError + "\r\n";
    }

    
    /************************************************************************************************************\
    Function:       WriteLogFile2
    Description:    Write to log files.
    \************************************************************************************************************/
    private void WriteLogFile2(string LogFile, string LogStr)
    {
        StreamWriter logFile = new StreamWriter(LogFile, true);

        if (logFile == null)
        {
            ERROR("Unable to open " + LogFile + " for writing!!!");
        }
        else
        {
            logFile.WriteLine(LogStr);
            logFile.Close();
        }
    }

    /************************************************************************************************************\
    Function:       SetHdmiSource
    Description:    Set HDMI source.
    \************************************************************************************************************/
    private void SetHdmiSource(ESourceRes eSourceRes)
    {
        // Command (general)
        string CMD_GEN_RESET = "!";
        string CMD_GEN_HDMI_DISABLE = ",";
        string CMD_GEN_HDMI_ENABLE = ".";

        // Command (clock)
        string CMD_CLK_25200 = "a";
        string CMD_CLK_27000 = "b";
        string CMD_CLK_27027 = "c";
        string CMD_CLK_74250 = "d";
        string CMD_CLK_92813 = "e";
        string CMD_CLK_111375 = "f";
        string CMD_CLK_148500 = "g";
        string CMD_CLK_185625 = "h";
        string CMD_CLK_222750 = "i";
        string CMD_CLK_297000 = "j";

        // Commands (resolutions)
        string CMD_RES_DVI_720x480P_RGB_8 = "l";
        string CMD_RES_HDMI_640x480P_RGB_8 = "m";
        string CMD_RES_HDMI_720x480P_RGB_8 = "n";
        string CMD_RES_HDMI_720x576P_RGB_8 = "o";
        string CMD_RES_HDMI_3840x2160P_YUV420_8 = "p";
        string CMD_RES_HDMI_1280x720P_RGB_8 = "q";
        string CMD_RES_HDMI_1280x720P_RGB_10 = "r";
        string CMD_RES_HDMI_1920x1080I50_RGB_8 = "s";
        string CMD_RES_HDMI_1920x1080I60_RGB_8 = "t";
        string CMD_RES_HDMI_1920x1080P30_RGB_8 = "u";
        string CMD_RES_HDMI_1920x1080P50_RGB_8 = "v";
        string CMD_RES_HDMI_1920x1080P60_RGB_8 = "w";
        string CMD_RES_HDMI_1920x1080P_RGB_10 = "x";
        string CMD_RES_HDMI_1920x1080P_RGB_12 = "y";
        string CMD_RES_HDMI_3840x2160P_RGB_8 = "z";
    
        string szCommand = "";

        if (eSourceRes == ESourceRes.DVI_720x480P60_27027_RGB_8)
            szCommand = CMD_CLK_27027 + CMD_RES_DVI_720x480P_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_640x480P60_25200_RGB_8)
            szCommand = CMD_CLK_25200 + CMD_RES_HDMI_640x480P_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_720x480P60_27027_RGB_8)
            szCommand = CMD_CLK_27027 + CMD_RES_DVI_720x480P_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_720x576P50_27000_RGB_8)
            szCommand = CMD_CLK_27000 + CMD_RES_HDMI_720x576P_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_3840x2160P60_594000_YUV420_8)
            szCommand = CMD_CLK_297000 + CMD_RES_HDMI_3840x2160P_YUV420_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1280x720P60_74250_RGB_8)
            szCommand = CMD_CLK_74250 + CMD_RES_HDMI_1280x720P_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1280x720P50_92813_RGB_10)
            szCommand = CMD_CLK_92813 + CMD_RES_HDMI_1280x720P_RGB_10 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1920x1080I50_74250_RGB_8)
            szCommand = CMD_CLK_74250 + CMD_RES_HDMI_1920x1080I50_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1920x1080I60_74250_RGB_8)
            szCommand = CMD_CLK_74250 + CMD_RES_HDMI_1920x1080I60_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1920x1080P30_74250_RGB_8)
            szCommand = CMD_CLK_74250 + CMD_RES_HDMI_1920x1080P30_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1920x1080P50_148500_RGB_8)
            szCommand = CMD_CLK_148500 + CMD_RES_HDMI_1920x1080P50_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1920x1080P60_148500_RGB_8)
            szCommand = CMD_CLK_148500 + CMD_RES_HDMI_1920x1080P60_RGB_8 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1920x1080P60_185625_RGB_10)
            szCommand = CMD_CLK_185625 + CMD_RES_HDMI_1920x1080P_RGB_10 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_1920x1080P60_222750_RGB_12)
            szCommand = CMD_CLK_222750 + CMD_RES_HDMI_1920x1080P_RGB_12 + CMD_GEN_RESET;
        else if (eSourceRes == ESourceRes.HDMI_3840x2160P24_297000_RGB_8)
            szCommand = CMD_CLK_297000 + CMD_RES_HDMI_3840x2160P_RGB_8 + CMD_GEN_RESET;

        m_oComTool.SendTextCommandEx(m_oSrcHandle, szCommand, false);
    }

    private int GetNextValue(EChangeAlgo eChangeAlgo, int iCurrent, int iCount, int iMask)
    {
        int iNext = iCurrent;

        if ((((1 << iCount) - 1) & iMask) != 0)
        {
            do
            {
                if (eChangeAlgo == EChangeAlgo.INCREMENTAL)
                {
                    iNext = (iNext + 1) % iCount;
                }
                else if (eChangeAlgo == EChangeAlgo.RANDOM)
                {
                    Random random = new Random();
                    iNext = random.Next(0, iCount); // Between 0 and (iCount-1)
                }
                else if (eChangeAlgo == EChangeAlgo.STATIC)
                {
                    iNext = iCurrent;
                }
            } while (((iMask >> iNext) & 0x1) == 0);
        }

        return iNext;
    }

    /************************************************************************************************************\
    Function:       ComPortTest
    Description:    Check echo for correct transmission. Infinite loop and error report.
    \************************************************************************************************************/
    private int ComPortTest()
    {
        m_oHostOptions.AppendToDebug("Com port test stated.", false);
        String oCmd;
        String oCmdStr1 = "0123456789"; // Test Data String. Avoid using 'r' or 'R' because they generate a response from the interface...
        String oCmdStr2 = "abcdefghij";
        String oCmdStr3 = "ABCDEFGHIJ";
        String oCmdStr4 = "DeadBeefA5";
        Byte[] abyCmd; // Command received
        bool bCrLf = false;
        Byte[] abyEcho = new Byte[10];
        int iNbToReceive = 10;
        Int32 iTimeoutMSec = 1000;
        bool bError = false;
        bool bNoTimeout = false;
        int iIteration = 1;
        int iSelCmd = 0;
        while (true)
        {
            if (iSelCmd == 0)
                oCmd = oCmdStr1;
            else if (iSelCmd == 1)
                oCmd = oCmdStr2;
            else if (iSelCmd == 2)
                oCmd = oCmdStr3;
            else
                oCmd = oCmdStr4;

            abyCmd = System.Text.Encoding.ASCII.GetBytes(oCmd);
            bNoTimeout = m_oComTool.SendTextCommandEx(
                m_oSrcHandle,
                oCmd,
                bCrLf,
                abyEcho,
                iNbToReceive,
                iTimeoutMSec
            );
            // bNoTimeout = m_oComTool.SendTextCommandEx(m_oAnaHandle, oCmd, bCrLf, abyEcho, iNbToReceive, iTimeoutMSec);
            if (bNoTimeout == true)
            {
                for (int iIdx = 0; !bError && (iIdx < abyCmd.Length); iIdx++)
                {
                    bError = abyCmd[iIdx] != abyEcho[iIdx];
                    if (bError == true)
                        m_oHostOptions.AppendToDebug("Character mismatch at iIdx = " + iIdx, false);
                }
            }
            else
            {
                m_oHostOptions.AppendToDebug("Timeout", false);
                bError = true;
            }
            if (bError == true)
            {
                m_oHostOptions.AppendToDebug(
                    "Com Port error detected at iteration = " + iIteration,
                    false
                );
                for (int i = 0; i < 10; i++)
                {
                    // m_oHostOptions.AppendToDebug( abyEcho[i] + "  0x" + abyEcho[i].ToString("X2"), false);
                    m_oHostOptions.AppendToDebug("0x" + abyEcho[i].ToString("X2") + "  ", false, false);
                    if ((abyEcho[i] > 32) && (abyEcho[i] < 127))
                        m_oHostOptions.AppendToDebug("" + (char)abyEcho[i], false);
                    else
                        m_oHostOptions.AppendToDebug("-", false);
                }
                // MessageBox.Show("Press key to continue");
                bError = false;
            }
            for (int i = 0; i < 10; i++) // Clear echo buffer
                abyEcho[i] = 0;
            iIteration += 1;

            if (iSelCmd == 3)
                iSelCmd = 0;
            else
                iSelCmd += 1;
        }
    }

    /************************************************************************************************************\
    Function:       ReadSourceInfo
    Description:    Read the source information and return a string
    \************************************************************************************************************/
    private string ReadSourceInfo(int SourceSel, HDMI_Source[] Hdmi_source)
    {
        // CStringEx strLine;
        // strLine = new CStringEx("");
        string strLine = "";

        strLine += String.Format("{0,4}", Hdmi_source[SourceSel].ResX);
        strLine += " x ";
        strLine += String.Format("{0,4}", Hdmi_source[SourceSel].ResY);
        if (Hdmi_source[SourceSel].IsInterlaced == true)
            strLine += "i";
        else
            strLine += "p";
        strLine += " @ ";
        strLine += String.Format("{0,2}", Hdmi_source[SourceSel].RefreshRate);
        strLine += " Hz";
        strLine += String.Format(" {0,6}", Hdmi_source[SourceSel].PixelClock);
        strLine += " kHz ";
        strLine += String.Format("{0,2}", Hdmi_source[SourceSel].DC);
        strLine += " bits ";
        if (Hdmi_source[SourceSel].Y1Y0 == 0)
            strLine += "RGB";
        else if (Hdmi_source[SourceSel].Y1Y0 == 1)
            strLine += "422";
        else if (Hdmi_source[SourceSel].Y1Y0 == 2)
            strLine += "444";
        else if (Hdmi_source[SourceSel].Y1Y0 == 3)
            strLine += "420";
        if (Hdmi_source[SourceSel].IsHdmi == true)
            strLine += " HDMI ";
        else
            strLine += "  DVI ";

        return (strLine);
    }

    /************************************************************************************************************\
    Function:       IoControlRead
    Description:    Read the com port response.
    \************************************************************************************************************/
    private byte IoControlRead()
    {
        bool bCrLf = false;
        byte[] abyResponse = new byte[16]; // Store response bytes when reading the com port.
        int iTimeoutMSec = 1000;
        int iNbToReceive = 2;
        bool bTimeRef = false;
        bool bNextLine = true;
        bool ComEchoReadSuccess = false;
        string oResponse;

        while (ComEchoReadSuccess == false)
        {
            ComEchoReadSuccess = m_oComTool.SendTextCommandEx(
                m_oIoHandle,
                "R",
                bCrLf,
                abyResponse,
                iNbToReceive,
                iTimeoutMSec
            );
            if (ComEchoReadSuccess == false)
            {
                iCommErr2 += 1;
                m_oHostOptions.AppendToDebug(
                    "Fail SendTextCommandEx in IoControlRead: Did not received enought character after 'R' command",
                    bTimeRef,
                    bNextLine
                );
                oResponse = System.Text.Encoding.UTF8.GetString(abyResponse);
                oResponse = oResponse.TrimEnd(new char[] { '\0' });
                m_oHostOptions.AppendToDebug("Response was " + oResponse, bTimeRef, bNextLine);
                // MessageBox.Show("Click OK to continue...");
            }
        }
        oResponse = System.Text.Encoding.UTF8.GetString(abyResponse);
        oResponse = oResponse.TrimEnd(new char[] { '\0' });
        string newstring = oResponse.Substring(oResponse.Length - 2, 2);
        byte byData = (byte)Convert.ToInt32(newstring, 16);
        return (byData);
    }

    /************************************************************************************************************\
    Function:       InitRegistersHdmiGen
    Description:    Init HDMI generator registers.
    \************************************************************************************************************/
    private void InitRegistersHdmiGen()
    {
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W300", false); // No Jitter
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400", false); // Enable lane OEn
        // OEn FPGA HDMI output enable before the redriver 0:enabled
        // <0> Data 0, <1> Data 1, <2> Data 2, <3> Data CLK
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W501", false); // Enable HDMI_+5V
        // <0> HDMI5V ouput  1: enabled, 0: disabled
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600", false); // Enable TMDS141 OEN
        // <0> tmds141_OEn  0:enabled, 1: disabled
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4", false); // No TMDS lane swap. Default CLK D2 D1 D0 "11100100"
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W800", false); // Scrambling disabled
        // <3>  <2>  <1>  <0>     0: scrambling disabled   1: scrambling enabled
        // CLK   D2   D1   D0
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W900", false); // Disable interpair delay
        // W9xx  Bit Delay TMDS D1 D0
        // <7:4> D1   <3:0> D0
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "WA00", false); // WAxx  Bit Delay TMDS CLK D2
        // <7:4> CLK  <3:0> D2
        // Maximum delay is 15 bits
        // W9XX must be followed by WAxx to activate the bit delay
    }

    /************************************************************************************************************\
    Function:       DisableHdmi
    Description:    Disable HDMI.
    \************************************************************************************************************/
    private void DisableHdmi()
    {
        // Create problem with switch
        m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);     // Reset Source
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40F", false); // Disable HDMI TMDS output (HI-Z)
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W500", false); // Disable HDMI output +5V
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W601", false); // Disable TMDS141 OEN
    }

    /************************************************************************************************************\
    Function:       EnableHdmi
    Description:    Enable HDMI.
    \************************************************************************************************************/
    private void EnableHdmi()
    {
        // Clean HDMI enable
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);     // Reset Source
        m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI TMDS
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400", false); // Enable HDMI TMDS output (HI-Z)
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W501", false); // Enable HDMI output +5V
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600", false); // Enable TMDS141 OEN
    }

    /************************************************************************************************************\
    Function:       SetClockJitter
    Description:    Defect CLOCK_JITTER.
    \************************************************************************************************************/
    private void SetClockJitter( int iPixFreqKHz, int iColorDepth )
    {
        // Calculate TMDS clock (rounded division)
        int iTmdsFreqKHz = iPixFreqKHz;
        if (iColorDepth == 10)
        {
            iTmdsFreqKHz = ((iTmdsFreqKHz * 125) + 50) / 100;
        }
        else if (iColorDepth == 12)
        {
            iTmdsFreqKHz = ((iTmdsFreqKHz * 150) + 50) / 100;
        }

        // Get VCO frequency
        const int FreqLimit25200  =  26000;
        const int FreqLimit27000  =  27010;
        const int FreqLimit27027  =  50000;
        const int FreqLimit74250  =  80000;
        const int FreqLimit92813  = 100000;
        const int FreqLimit148500 = 160000;
        const int FreqLimit185625 = 200000;
        const int FreqLimit222750 = 250000;
        // const int FreqLimit297000 = 300000;

        const int FreqVco25200  =  630000;
        const int FreqVco27000  =  945000;
        const int FreqVco27027  = 1081250;
        const int FreqVco74250  =  742500;
        const int FreqVco92813  =  928125;
        const int FreqVco148500 =  742500;
        const int FreqVco185625 =  928125;
        const int FreqVco222750 = 1113750;
        const int FreqVco297000 = 1485000;

        int iVcoFreqKHz = 0;
        
        if (iTmdsFreqKHz < FreqLimit25200)
        {
            iVcoFreqKHz = FreqVco25200;
        }
        else if (iTmdsFreqKHz < FreqLimit27000)
        {
            iVcoFreqKHz = FreqVco27000;
        }
        else if (iTmdsFreqKHz < FreqLimit27027)
        {
            iVcoFreqKHz = FreqVco27027;
        }
        else if (iTmdsFreqKHz < FreqLimit74250)
        {
            iVcoFreqKHz = FreqVco74250;
        }
        else if (iTmdsFreqKHz < FreqLimit92813)
        {
            iVcoFreqKHz = FreqVco92813;
        }
        else if (iTmdsFreqKHz < FreqLimit148500)
        {
            iVcoFreqKHz = FreqVco148500;
        }
        else if (iTmdsFreqKHz < FreqLimit185625)
        {
            iVcoFreqKHz = FreqVco185625;
        }
        else if (iTmdsFreqKHz < FreqLimit222750)
        {
            iVcoFreqKHz = FreqVco222750;
        }
        else // if (iTmdsFreqKHz < FreqLimit297000)
        {
            iVcoFreqKHz = FreqVco297000;
        }

        // Set clock jitter to the desired amount according to pixel rate
        // Refer to PhaseShiftCtrl.vhd for amplitude settings
        float fJitterNum = iJitterAmplitude_mTbit * 100000 / (float) iTmdsFreqKHz;
        float fJitterDen = 1000000000 / (float) iVcoFreqKHz / 56;
        int   iJitterAmp = (int) ((fJitterNum / fJitterDen) + 0.5);

        string strCommand = "W300";
        if (iJitterAmp < 256)
        {
            // Set clock jitter to desired Tbit amount
            strCommand = "W3" + iJitterAmp.ToString("X2");
        }
        m_oComTool.SendTextCommandEx(strCommand, false);
    }

    /************************************************************************************************************\
    Function:       InterPairDelayRandom
    Description:    Defect SYMBOL_SKEW_DELAY.
    \************************************************************************************************************/
    private void InterPairDelayRandom(int iMaxInterPairDelay)
    {
        bool PrintDelayValues = false;
        Random random = new Random();
        int iDelay_D0 = random.Next(0, iMaxInterPairDelay + 1); // Between 0 and iMaxInterPairDelay (rightmost is excluded maximum)
        int iDelay_D1 = random.Next(0, iMaxInterPairDelay + 1);
        int iDelay_D2 = random.Next(0, iMaxInterPairDelay + 1);
        int iDelay_CLK = random.Next(0, iMaxInterPairDelay + 1);
        if (PrintDelayValues == true)
        {
            m_oHostOptions.AppendToDebug(" D0=" + iDelay_D0, false, false);
            m_oHostOptions.AppendToDebug(" D1=" + iDelay_D1, false, false);
            m_oHostOptions.AppendToDebug(" D2=" + iDelay_D2, false, false);
            m_oHostOptions.AppendToDebug(" Dclk=" + iDelay_CLK, false, false);
        }
        int Reg_0x9_D1D0 = iDelay_D1 * 16 + iDelay_D0;
        int Reg_0xA_ClkD2 = iDelay_CLK * 16 + iDelay_D2;

        if (PrintDelayValues == true)
        {
            m_oHostOptions.AppendToDebug(" W9" + Reg_0x9_D1D0.ToString("X2"), false, false);
            m_oHostOptions.AppendToDebug(" WA" + Reg_0xA_ClkD2.ToString("X2") + " ", false, false);
        }
        // W9XX must be followed by WAxx to activate the bit delay
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W9" + Reg_0x9_D1D0.ToString("X2"), false); // W9xx  Bit Delay TMDS D1 D0
        // <7:4> D1   <3:0> D0
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "WA" + Reg_0xA_ClkD2.ToString("X2"), false); // WAxx  Bit Delay TMDS CLK D2
        // <7:4> CLK  <3:0> D2
        // Maximum delay is 15 bits
        
        
        if( bTransitionalInterPairDelay == true )
        {
           // Wait a bit
           Thread.Sleep(iINTER_PAIR_DELAY_TIME_MS); //
           
           // Remove the inter-pair delay
           // W9XX must be followed by WAxx to activate the bit delay
           m_oComTool.SendTextCommandEx(m_oSrcHandle, "W900", false); // W9xx  Bit Delay TMDS D1 D0
           // <7:4> D1   <3:0> D0
           m_oComTool.SendTextCommandEx(m_oSrcHandle, "WA00", false); // WAxx  Bit Delay TMDS CLK D2
           // <7:4> CLK  <3:0> D2
        }
        
    }

    /************************************************************************************************************\
    Function:       SequentialDataRandom
    Description:    Defect SEQUENTIAL_DATA.
    \************************************************************************************************************/
    private void SequentialDataRandom()
    {
        bool bPrintSequ = false;
        // Enable lane OEn
        // OEn FPGA HDMI output enable before the redriver 0:enabled
        // <3> Data CLK,  <2> Data 2,  <1> Data 1,  <0> Data 0
        m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40F", false); // Disable CLK, D2, D1 and D0
        Thread.Sleep(100);
        m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable FPGA HDMI Datapath output
        Thread.Sleep(100);
        // Enable each lane in random order, use mutiple of 100ms for sequencing
        int iMaxSequDelay_ms = 100;
        Random random = new Random();
        int iDelay_D0 = random.Next(0, iMaxSequDelay_ms + 1); // Between 0 and iMaxSequDelay_ms (rightmost is excluded maximum)
        int iDelay_D1 = random.Next(0, iMaxSequDelay_ms + 1);
        int iDelay_D2 = random.Next(1, iMaxSequDelay_ms + 1);
        int iDelay_CLK = random.Next(1, iMaxSequDelay_ms + 1);
        int iQuartet = 0x0F;
        int iTime_ms = 0;
        for (iTime_ms = 0; iTime_ms <= iMaxSequDelay_ms; iTime_ms++)
        {
            if (iTime_ms == iDelay_D0) // Time to enable D0 ?
            {
                iQuartet = iQuartet & 0x0E;
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40" + iQuartet.ToString("X1"), false);
                if (bPrintSequ == true)
                    m_oHostOptions.AppendToDebug(" D0=" + iDelay_D0, false, false);
            }
            if (iTime_ms == iDelay_D1) // Time to enable D1 ?
            {
                iQuartet = iQuartet & 0x0D;
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40" + iQuartet.ToString("X1"), false);
                if (bPrintSequ == true)
                    m_oHostOptions.AppendToDebug(" D1=" + iDelay_D1, false, false);
            }
            if (iTime_ms == iDelay_D2) // Time to enable D2 ?
            {
                iQuartet = iQuartet & 0x0B;
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40" + iQuartet.ToString("X1"), false);
                if (bPrintSequ == true)
                    m_oHostOptions.AppendToDebug(" D2=" + iDelay_D2, false, false);
            }
            if (iTime_ms == iDelay_CLK) // Time to enable CLK ?
            {
                iQuartet = iQuartet & 0x07;
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40" + iQuartet.ToString("X1"), false);
                if (bPrintSequ == true)
                    m_oHostOptions.AppendToDebug(" Clk=" + iDelay_CLK, false, false);
            }
            Thread.Sleep(1); // pause 1ms
        }
    }

    /************************************************************************************************************\
    Function:       GlitchyOutputEnableRandom
    Description:    Defect TMDS_OUTPUT_ENABLE.
    \************************************************************************************************************/
    private void GlitchyOutputEnableRandom() // Toggle the TMDS141 redriver output enable a multiple of time
    {
        bool bPrintNumberOfToggle = false; // for debug
        int iMaxNumberOfTogging = 15; // Maximum Number of Toggling allowed by the random generator
        int iMaxOffTime_ms = 250;
        int iMaxOnTime_ms = 250;
        //
        // Enable HDMI output
        m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI FPGA TMDS driver
        //
        Random random = new Random(); // Determine a random number of toggling to do within allowed maximum
        int NumberOfTogging = random.Next(2, iMaxNumberOfTogging + 1);
        if (bPrintNumberOfToggle == true)
        {
            m_oHostOptions.AppendToDebug(" T" + NumberOfTogging, false, false);
        }
        //
        for (int iToggleCount = 0; iToggleCount < NumberOfTogging; iToggleCount++) // Loop for each toogling
        {
            int iOffTime_ms = random.Next(0, iMaxOffTime_ms + 1); // Determine a random number for Off time
            //
            int iOnTime_ms = random.Next(0, iMaxOnTime_ms + 1); // Determine a random number for On time
            // Turn OFF TMDS141 redriver
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "W601", false); // Disable TMDS141 OEN
            // <0> tmds141_OEn  0:enabled, 1: disabled
            Thread.Sleep(iOffTime_ms);
            // Turn ON TMDS141 redriver
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600", false); // Enable TMDS141 OEN
            Thread.Sleep(iOffTime_ms);
        }
    }

    /************************************************************************************************************\
    Function:       ScramblingRandom
    Description:    Defect TMDS_SCRAMBLING.
    \************************************************************************************************************/
    private void ScramblingRandom() // insert tmds scrambling a random number of time
    {
        bool bPrintNumberOfScramblingEvents = false; // for debug
        int iMaxNumberOfEvents = 15; // Maximum Number of Toggling allowed by the random generator
        int iMaxOffTime_ms = 250;
        int iMaxOnTime_ms = 250;
        //
        // Enable HDMI output
        m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI FPGA TMDS driver
        //
        Random random = new Random(); // Determine a random number of scrambling events to do within allowed maximum
        int NumberOfEvents = random.Next(2, iMaxNumberOfEvents + 1);
        if (bPrintNumberOfScramblingEvents == true)
        {
            m_oHostOptions.AppendToDebug(" S" + NumberOfEvents, false, false);
        }
        //
        for (int iEventCount = 0; iEventCount < NumberOfEvents; iEventCount++) // Loop for each toogling
        {
            int iOffTime_ms = random.Next(0, iMaxOffTime_ms + 1); // Determine a random scrambling OFF time
            int iOnTime_ms = random.Next(0, iMaxOnTime_ms + 1); // Determine a random scrambling ON time
            // Turn ON scrambling on all lanes
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "W80F", false); // <3>  <2>  <1>  <0>   0: scrambling disabled
            // CLK   D2   D1   D0   1: scrambling enabled
            Thread.Sleep(iOnTime_ms);
            // Turn OFF scrambling on all lanes
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "W800", false);
            Thread.Sleep(iOffTime_ms);
        }
    }

    /************************************************************************************************************\
    Function:       Disconnect5VRandom
    Description:    Defect RANDOM_5V_DISCONNECT.
    \************************************************************************************************************/
    private void Disconnect5VRandom()
    {
        bool bPrintNumberOfDisconnectEvents = false; // for debug
        int iMaxNumberOfEvents = 15; // Maximum Number of Toggling allowed by the random generator
        int iMaxOffTime_ms = 250;
        int iMaxOnTime_ms = 250;
        //
        // Enable HDMI output
        m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI FPGA TMDS driver
        //
        Random random = new Random(); // Determine a random number of +5V disconnect events to do within allowed maximum
        int NumberOfEvents = random.Next(2, iMaxNumberOfEvents + 1);
        if (bPrintNumberOfDisconnectEvents == true)
        {
            m_oHostOptions.AppendToDebug(" D" + NumberOfEvents, false, false);
        }
        //
        for (int iEventCount = 0; iEventCount < NumberOfEvents; iEventCount++) // Loop for each event
        {
            int iOffTime_ms = random.Next(0, iMaxOffTime_ms + 1); // Determine a random 5V OFF time
            int iOnTime_ms = random.Next(0, iMaxOnTime_ms + 1); // Determine a random 5V ON time
            // Turn OFF HDMI +5V
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "W500", false); // <0> HDMI5V ouput  1: enabled, 0: disabled
            Thread.Sleep(iMaxOffTime_ms);
            // Turn ON HDMI +5V
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "W501", false);
            Thread.Sleep(iOnTime_ms);
        }
    }

    /************************************************************************************************************\
    Function:       SwapTmdsPairRandom
    Description:    Defect TMDS_SWAP.
    \************************************************************************************************************/
    private void SwapTmdsPairRandom()
    {
        bool bPrintNumberOfSwapEvents = false; // for debug
        int iMaxNumberOfEvents = 15; // Max Number of swap allowed by the random generator
        int iMaxOffTime_ms = 250; // Max time spend on the not swapped case (Normal lane assignment)
        int iMaxOnTime_ms = 250; // Max time spend on the swapped situation
        //
        // Enable HDMI output
        m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI FPGA TMDS driver
        //
        Random random = new Random(); // Determine a random number of lane swap events to do within allowed maximum
        int NumberOfEvents = random.Next(2, iMaxNumberOfEvents + 1);
        if (bPrintNumberOfSwapEvents == true)
        {
            m_oHostOptions.AppendToDebug(" S" + NumberOfEvents, false, false);
        }
        //
        for (int iEventCount = 0; iEventCount < NumberOfEvents; iEventCount++) // Loop for each event
        {
            int iOnTime_ms = random.Next(0, iMaxOnTime_ms + 1); // Determine a random time for swapped
            int iOffTime_ms = random.Next(0, iMaxOffTime_ms + 1); // Determine a random time for not swapped

            int iSwapCase = random.Next(0, 6); // Randomly pick one and apply the swap case from 0 to 5
            if (iSwapCase == 0) // 0: CLK D2 D1 D0
            {
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4", false); // No TMDS lane swap. Default CLK D2 D1 D0 "11100100"
            }
            else if (iSwapCase == 1) // 1: CLK D2 D0 D1
            {
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E1", false); // "11100001"
            }
            else if (iSwapCase == 2) // 2: CLK D1 D2 D0
            {
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7D8", false); // "11011000"
            }
            else if (iSwapCase == 3) // 3: CLK D1 D0 D2
            {
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7D2", false); // "11010010"
            }
            else if (iSwapCase == 4) // 4: CLK D0 D2 D1
            {
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C9", false); // "11001001"
            }
            else if (iSwapCase == 5) // 5: CLK D0 D1 D2
            {
                m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6", false); // "11000110"
            }
            Thread.Sleep(iOnTime_ms); // Apply this swap a time amount
            // Restore normal lane assignment
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4", false); // No TMDS lane swap. Default CLK D2 D1 D0 "11100100"
            Thread.Sleep(iOffTime_ms);
        }
    }

    /************************************************************************************************************\
    Function:       HdmiTransition
    Description:    Apply the selected resolution and defect.
    \************************************************************************************************************/
    private void HdmiTransition(EDefectType eDefectType, ESourceRes eSourceRes, int iPixFreqKHz, int iColorDepth)
    {
        const int TMDS_OFF_TIME_MS = 200;
        if (eDefectType == EDefectType.CLEAN) // Perfect transition
        {
            // ==========================================================
            // Disable HDMI, Set source, Wait TMDSOffTime, Enable HDMI
            // ==========================================================
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Enable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI TMDS
        }
        else if (eDefectType == EDefectType.CLOCK_JITTER) // Add clock jitter 0.3Tbit
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Add clock jitter
            SetClockJitter(iPixFreqKHz, iColorDepth);

            if( bTransitionalJitter == true )
            {
               // Wait a bit
               Thread.Sleep(iJITTER_TIME_MS);

               // Remove jitter
               m_oComTool.SendTextCommandEx(m_oSrcHandle, "W300", false);
            }

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Enable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI TMDS
        }
        else if (eDefectType == EDefectType.SYMBOL_SKEW_DELAY) // InterPair Delay random
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Add InterPair delay random
            InterPairDelayRandom(2);

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Enable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ".", false); // Enable HDMI TMDS
        }
        else if (eDefectType == EDefectType.SEQUENTIAL_DATA) // Sequential Data Random
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Add InterPair delay random
            SequentialDataRandom();
        }
        else if (eDefectType == EDefectType.TMDS_OUTPUT_ENABLE) // Glitchy Output Enable (Multiple random toggling)
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Glitchy HDMI output enable, Toggle a random number of time with random delay
            GlitchyOutputEnableRandom();
        }
        else if (eDefectType == EDefectType.TMDS_SCRAMBLING) // scrambling random
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Glitchy HDMI output enable, Toggle a random number of time with random delay
            ScramblingRandom();
        }
        else if (eDefectType == EDefectType.RANDOM_5V_DISCONNECT) // 5V Disconnect
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Turn OFF-ON the +5V on HDMI a random number of time
            Disconnect5VRandom();
        }
        else if (eDefectType == EDefectType.TMDS_SWAP) // Swap TMDS Pair randomly
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Do pair swap
            SwapTmdsPairRandom();
        }
        else if (eDefectType == EDefectType.ALL_DRESS) // Full defect
        {
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);

            // Add clock jitter
            SetClockJitter(iPixFreqKHz, iColorDepth);
            
             // Add InterPair delay random
            InterPairDelayRandom(2);
            
            // Add InterPair delay random
            SequentialDataRandom();
            
            // Glitchy HDMI output enable, Toggle a random number of time with random delay
            GlitchyOutputEnableRandom();
            
            ScramblingRandom();
            
            // Do pair swap
            SwapTmdsPairRandom();
            
            // Do pair swap
            SwapTmdsPairRandom();
        }
        if (eDefectType == EDefectType.BLACK) // Black screen
        {
            // ==========================================================
            // Disable HDMI, Set source, Wait TMDSOffTime
            // ==========================================================
            // Disable HDMI output
            m_oComTool.SendTextCommandEx(m_oSrcHandle, ",", false); // Disable HDMI

            // Set resolution and clock rate
            SetHdmiSource(eSourceRes);

            // Reset source
            m_oComTool.SendTextCommandEx(m_oSrcHandle, "!", false); // Reset source

            // Wait a bit
            Thread.Sleep(TMDS_OFF_TIME_MS);
        }

        // for( int iTroubleMakerCnt = 1; iTroubleMakerCnt < 3 ; iTroubleMakerCnt++)
        // {
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "W71B",false);   // 1B    D0  D1  D2  CLK   "00011011"  Swap All
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "d",false);      // 74.250000 MHz Clock
        //    Thread.Sleep(500);
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "i",false);      // 222.75 MHz Clock
        //    Thread.Sleep(500);
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "W76C",false);   // 6C    D1  D2  CLK  D0   "01101100"  Swap CLK and D1
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "s",false);      // 1080i_8_50Hz_CBar_RGB_TV_v54
        //    Thread.Sleep(1000);
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "h",false);     // 185.625 MHz Clock
        //    Thread.Sleep(500);
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7B4",false);  // B4     D2  CLK D1  D0    "10110100"  Swap CLK and D2
        //    Thread.Sleep(1000);
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "W700",false);  // 00     D0  D0  D0  D0    "00000000   D0 on all pairs
        //    // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W755",false);  // 55     D1  D1  D1  D1    "01010101   D1 on all pairs
        //    Thread.Sleep(1000);
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "z",false);     // 4K2K_2160p_24Hz
        //    Thread.Sleep(1500);
        //    m_oComTool.SendTextCommandEx(m_oSrcHandle, "f",false);     // 111.375 MHz
        //    Thread.Sleep(1000);
        //    SetHdmiSource(eSourceRes);
        // }

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W350",false);   // Put Jitter

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7A4",false);   // A4     D2  D2  D1  D0    "10100100"  CLK replace with D2
        // Thread.Sleep(250);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7D8",false);   // D8    CLK  D1  D2  D0    "11011000"  Swap D2 with D1
        // Thread.Sleep(250);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data
        // Thread.Sleep(250);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W300",false);   // Remove Jitter
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E5",false);   // E5    CLK  D2  D1  D1    "11100101"  D0 replace with D1
        // Thread.Sleep(250);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E1",false);   // E1    CLK  D2  D0  D1    "11100001"  Swap D0 with D1
        // Thread.Sleep(2000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W71B",false);   // 1B    D0  D1  D2  CLK   "00011011"  Swap All

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4",false);   // Default CLK D2 D1 D0 "11100100"
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // DisableHdmi();
        // SetHdmiSource(eSourceRes);
        // Thread.Sleep(TMDSOffTimeMs);
        // EnableHdmi();

        // Disable HDMI output
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ",",false);     // Disable HDMI
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40F",false);  // Disable HDMI TMDS output (HI-Z)
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W500",false);  // Disable HDMI output +5V
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W601",false);  // Disable TMDS141 OEN
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7D8",false);   // D8    CLK  D1  D2  D0    "11011000"  Swap D1 with D2
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W700",false);   // 00     D0  D0  D0  D0    "00000000"  D0 on all pairs
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E8",false);   // E8    CLK  D2  D2  D0    "11101000"  Replace D1 with D2

        // Enable HDMI output
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ".",false);     // Enable HDMI TMDS
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);  // Enable HDMI TMDS output (no HI-Z)
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W501",false);  // Enable HDMI output +5V
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600",false);  // Enable TMDS141 OEN

        // Thread.Sleep(2000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4",false);   // Default CLK D2 D1 D0 "11100100"
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // Thread.Sleep(320);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // Thread.Sleep(120);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // Thread.Sleep(96);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data
        // Thread.Sleep(128);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4",false);   // Default CLK D2 D1 D0 "11100100"
        // Thread.Sleep(256);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // Thread.Sleep(512);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data
        // Thread.Sleep(300);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // Thread.Sleep(2000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4",false);   // E4    CLK D2 D1 D0  "11100100"   Default
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK D0 D1 D2  "11000110"   Swap only Data
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4",false);   // E4    CLK D2 D1 D0  "11100100"   Default
        // Thread.Sleep(200);

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W601",false);   // Disable TMDS141 OEN
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40F",false);   // Stop D0, D1, D2, CLK
        // Thread.Sleep(30);                                           // 30 ms
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600",false);   // Enable TMDS141 OEN
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40B",false);   // Go D2
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W409",false);   // Go D1
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W408",false);   // Go D0
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);   // Go CLK
        // Thread.Sleep(800);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W601",false);   // Disable TMDS141 OEN
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40F",false);   // Stop D0, D1, D2, CLK
        // SetHdmiSource(eSourceRes);
        // Thread.Sleep(30);                                           // 30 ms
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600",false);   // Enable TMDS141 OEN
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40B",false);   // Go D2
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W409",false);   // Go D1
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W408",false);   // Go D0
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);   // Go CLK


        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40F",false);   // OEn Disable all lanes
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W807",false);   // Enable scrambling on data lanes

        // SetHdmiSource(eSourceRes);

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ",",false);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "gw",false);     // 148.5 MHz , 1080p_8_60Hz_CBar_RGB_TV_v54
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ".",false);

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // Thread.Sleep(2000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W408",false);   // OEn Activate lanes with Disabled Clock
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);   // OEn Reenable clock
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W800",false);   // Disable scrambling
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4!",false);   // Default CLK D2 D1 D0 "11100100"
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ",",false);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ".",false);

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data
        // SetHdmiSource(eSourceRes);
        // Thread.Sleep(2000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4",false);   // Default no lane swap CLK D2 D1 D0 "11100100"
        // Thread.Sleep(DelayVariable);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source

        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7C6",false);   // C6    CLK  D0  D1  D2    "11000110"   Swap only Data
        // SetHdmiSource(eSourceRes);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "gw",false);     // 148.5 MHz , 1080p_8_60Hz_CBar_RGB_TV_v54
        // Thread.Sleep(1000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source
        // Thread.Sleep(1000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W7E4!",false);   // Default CLK D2 D1 D0 "11100100"

        // Thread.Sleep(DelayVariable);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "!",false);      // Reset source

        // if( DelayVariable < 500 )
        // DelayVariable = DelayVariable + 1;
        // else
        // DelayVariable = 0;

        // // Create problem with switch

        // // Add weard sequencing
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W300",false);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "jl",false);     // 297MHz 480p
        // Thread.Sleep(300);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "cz",false);     // 27MHz 4K
        // Thread.Sleep(1000);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W601",false);  // Disable TMDS141 OEN
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600",false);  // Enable TMDS141 OEN
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W304",false);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "ip",false);     // 222.75MHz 720p
        // Thread.Sleep(300);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "d",false);     // 27MHz 1080p 10b
        // Thread.Sleep(300);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "jl",false);     // 297MHz 1080p 12b
        // // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W30A",false);   // Work
        // // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W30D",false);   // occasional failure
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W300",false);
        // 
        // SetHdmiSource(eSourceRes);  // put right pixel clock
        // Thread.Sleep(1000);

        // // Max Pix Clk before settling...
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "j",false);     // Max pix clock... 300 MHz
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ".",false);     // Enable HDMI TMDS
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W407",false);  // Enable only clock
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W501",false);  // Enable HDMI output +5V
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W600",false);  // Enable TMDS141 OEN
        // Thread.Sleep(1500);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);  // Enable all TMDS line
        // SetHdmiSource(eSourceRes); // set final pixel clock (with resolution)

        // // Mix of events...
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, ".",false);     // Enable HDMI
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "j",false);     // 300MHz clock...
        // Thread.Sleep(300);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "a",false);     // 25MHz clock...
        // Thread.Sleep(300);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W501",false);  // Enable HDMI output +5V
        // // Sequencing D0, D1, D2, CLK
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40E",false);  // HDMI D0 enabled     0:enable
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40C",false);  // HDMI D0 and D1 enabled
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W408",false);  // HDMI D0, D1 and D2 enabled
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);  // HDMI D0, D1, D2 and CLK enabled
        // Thread.Sleep(100);
        // SetHdmiSource(eSourceRes);  // put right pixel clock
        // Thread.Sleep(200);

        // // Sequencing D0, D1, D2, CLK
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40E",false);  // HDMI D0 enabled     0:enable
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40C",false);  // HDMI D0 and D1 enabled
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W408",false);  // HDMI D0, D1 and D2 enabled
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);  // HDMI D0, D1, D2 and CLK enabled

        // // Sequencing CLK D0, D1, D2
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W407",false);  // HDMI CLK enabled     0:enable
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W403",false);  // HDMI CLK and D2 enabled
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W401",false);  // HDMI CLK, D2 and D1 enabled
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);  // HDMI CLK, D2, D1 and D0 enabled

        // // Sequencing weardo
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40E",false);  // HDMI D0 enabled     0:enable
        // Thread.Sleep(300);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W405",false);  // HDMI CLK and D1 enabled
        // Thread.Sleep(400);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W40A",false);  // HDMI D2 and D0 enabled
        // Thread.Sleep(200);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W402",false);  // HDMI CLK, D2, and D0 enabled
        // Thread.Sleep(500);
        // m_oComTool.SendTextCommandEx(m_oSrcHandle, "W400",false);  // HDMI CLK, D2, D1 and D0 enabled
    }

    /************************************************************************************************************\
    Function:       CheckMonitorIsDisplaying
    Description:    Check monitor is displaying.
    \************************************************************************************************************/
    private bool CheckMonitorIsDisplaying()
    {
        int DataTemp = IoControlRead() & 0x01;
        if (DataTemp == 0)
            return (true);
        else
            return (false);
    }

    /************************************************************************************************************\
    Function:       CamSnapshot
    Description:    Take snapshot picture with webcam
    \************************************************************************************************************/
    private string CamSnapshot()
    {
        bool bHideCmdBox = true;
        Process oProcess = new Process();
        oProcess.StartInfo.FileName = "CommandCam";
        oProcess.StartInfo.Arguments = "";
        oProcess.StartInfo.UseShellExecute = false;
        // Redirect the output stream of the child process.
        oProcess.StartInfo.RedirectStandardError = true;
        oProcess.StartInfo.RedirectStandardOutput = true;
        if (bHideCmdBox == true) // to hide process cmd box
        {
            oProcess.StartInfo.CreateNoWindow = true;
            oProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
        }
        string stdError;
        try
        {
            oProcess.Start();
            // Read the output stream first and then wait.
            stdError = oProcess.StandardOutput.ReadToEnd();
            stdError += oProcess.StandardError.ReadToEnd();
            oProcess.WaitForExit();
        }
        catch (Exception oEx)
        {
            // m_oHostOptions.AppendToDebug(" Exeption in CamSnapshot : " + oEx.Message,false,false);
            string sError = " Exception in CamSnapshot : " + oEx.Message;
            // return oEx.Message;
            return sError;
        }

        if (oProcess.ExitCode == 0) // Normal exit?
        {
            return ""; // return nothing when normal
        }
        else
        {
            return "Not exit code 0, " + stdError;
        }
        // Thread.Sleep(4000);  // wait // unreachable code
    }

    /************************************************************************************************************\
    Function:       ImageProcessor
    Description:    Image processor load image and try to recognise what is displayed on each monitor
    \************************************************************************************************************/
    private string ImageProcessor()
    {
        bool bHideCmdBox = true;
        Process oProcess = new Process();
        // oProcess.StartInfo.FileName  = "ImageProcess";
        oProcess.StartInfo.FileName = "/../cprogram/ImageProcess/bin/Debug/ImageProcess";
        oProcess.StartInfo.Arguments = "";
        oProcess.StartInfo.UseShellExecute = false;
        // Redirect the output stream of the child process.
        oProcess.StartInfo.RedirectStandardError = true;
        oProcess.StartInfo.RedirectStandardOutput = true;
        if (bHideCmdBox == true) // to hide process cmd box
        {
            oProcess.StartInfo.CreateNoWindow = true;
            oProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
        }
        // string stdError;
        string stdOutput;
        try
        {
            oProcess.Start();
            // Read the output stream first and then wait.
            stdOutput = oProcess.StandardOutput.ReadToEnd();
            // stdError += oProcess.StandardError.ReadToEnd();
            oProcess.WaitForExit();
            return stdOutput;
        }
        catch (Exception oEx)
        {
            // m_oHostOptions.AppendToDebug(" Exeption in CamSnapshot : " + oEx.Message,false,false);
            string sError = " Exeption in ImageProcessor : " + oEx.Message;
            // return oEx.Message;
            return sError;
        }

        // if (oProcess.ExitCode == 0)   // Normal exit?
        // {
        //    return "";     // return nothing when normal
        //    // return stdOutput;
        // }
        // else
        // {
        //    return "Not exit code 0, " + stdError;
        // }
    }

    /************************************************************************************************************\
    Function:       PingDut
    Description:    Return the ping result on the ip adress.
    \************************************************************************************************************/
    private string PingDut(string IP_Address)
    {
        // bool bPingTestSuccess = false;  // By default we do not pass

        bool bHideCmdBox = true;
        Process oProcess = new Process();
        oProcess.StartInfo.FileName = "ping.exe";
        // oProcess.StartInfo.Arguments = "";
        oProcess.StartInfo.Arguments = "-n 1 " + IP_Address; // -n 1 is for only one reply for faster results
        oProcess.StartInfo.UseShellExecute = false;
        // Redirect the output stream of the child process.
        oProcess.StartInfo.RedirectStandardError = true; // if true i will see text in a command shell
        oProcess.StartInfo.RedirectStandardOutput = true;
        if (bHideCmdBox == true) // to hide process cmd box
        {
            oProcess.StartInfo.CreateNoWindow = true;
            oProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
        }

        // string stdError;
        string stdOutput;
        try
        {
            oProcess.Start();
            // Read the output stream first and then wait.
            stdOutput = oProcess.StandardOutput.ReadToEnd();
            // stdError += oProcess.StandardError.ReadToEnd();
            oProcess.WaitForExit();
            return stdOutput;
        }
        catch (Exception oEx)
        {
            // m_oHostOptions.AppendToDebug(" Exeption in CamSnapshot : " + oEx.Message,false,false);
            string sError = " Exeption in PingDut : " + oEx.Message;
            // return oEx.Message;
            return sError;
        }
    }

    /************************************************************************************************************\
    Function:       AnalysePingResult
    Description:    Parse Ping Result on DUT to conclude if DUT corectly replyed

                    Example:
                    Pinging 192.168.182.218 with 32 bytes of data:
                    Reply from 192.168.182.218: bytes=32 time<1ms TTL=61
    
                    Ping statistics for 192.168.182.218:
                    Packets: Sent = 1, Received = 1, Lost = 0 (0% loss),
                    Approximate round trip times in milli-seconds:
                    Minimum = 0ms, Maximum = 0ms, Average = 0ms

     Return Values: true=PASS, false=FAIL
    \************************************************************************************************************/
    private bool AnalysePingResult(string PingResult)
    {
        bool bTestPass = false; // by default test is failing
        string[] sLines = PingResult.Split(
            new char[] { '\n', '\r' },
            StringSplitOptions.RemoveEmptyEntries
        );

        // m_oHostOptions.AppendToDebug("0 " + sLines[0],false,true);  // Pinging 192.168.182.218 with 32 bytes of data:
        // m_oHostOptions.AppendToDebug("1 " + sLines[1],false,true);  // Reply from 192.168.182.218: bytes=32 time<1ms TTL=61
        // m_oHostOptions.AppendToDebug("2 " + sLines[2],false,true);  // Ping statistics for 192.168.182.218:
        // m_oHostOptions.AppendToDebug("3 " + sLines[3],false,true);  //      Packets: Sent = 1, Received = 1, Lost = 0 (0% loss),
        // m_oHostOptions.AppendToDebug("4 " + sLines[4],false,true);  // Approximate round trip times in milli-seconds:
        // m_oHostOptions.AppendToDebug("5 " + sLines[5],false,true);  //      Minimum = 1ms, Maximum = 1ms, Average = 1ms

        // We will just search for the word "Reply" on line 1 to conlude that the DUT pass the ping test
        string[] sLineOfInterest = sLines[1].Split(
            new char[] { ' ' },
            StringSplitOptions.RemoveEmptyEntries
        );
        // m_oHostOptions.AppendToDebug("0 " + sLineOfInterest[0],false,true);  // Reply
        // m_oHostOptions.AppendToDebug("1 " + sLineOfInterest[1],false,true);  // from
        // m_oHostOptions.AppendToDebug("2 " + sLineOfInterest[2],false,true);  // 192.168.182.218:
        // m_oHostOptions.AppendToDebug("3 " + sLineOfInterest[3],false,true);  // bytes=32
        // m_oHostOptions.AppendToDebug("4 " + sLineOfInterest[4],false,true);  // time<1ms

        if (sLineOfInterest[0] == "Reply")
        {
            bTestPass = true; // we found the condition that make the test pass
        }
        return bTestPass;
    }

    /************************************************************************************************************\
    Function:       MainProcess
    Description:    Run the test.
    \************************************************************************************************************/
    private void MainProcess()
    {
        if (bComPortTest == true)
            ComPortTest(); // Com port infinite loop test...

        // List of source resolutions from HdmiGen FPGA on the NEXYS VIDEO board.
        HDMI_source = new HDMI_Source[(int)ESourceRes.COUNT];
        // LEGEND .....................................................                      X     Y  Rate Interl   PClk    Hf  Hs  Hbp  HpolPos Vf  Vs Vbp  VpolPos VIC  DC Y1Y0 HDMI
        HDMI_source[(int) ESourceRes.DVI_720x480P60_27027_RGB_8]        = new HDMI_Source(  720,  480, 60, false,  27027,   16, 62,  60, false,   9,  6, 30, false,   0,   8,  0, false);  // 480p_8_60Hz_CBar_RGB_PC_v53_Dvi
        HDMI_source[(int) ESourceRes.HDMI_640x480P60_25200_RGB_8]       = new HDMI_Source(  640,  480, 60, false,  25203,   16, 96,  48, false,  10,  2, 33, false,   1,   8,  0, true );  // VGA_8_60Hz_CBar_RGB_TV_v54
        HDMI_source[(int) ESourceRes.HDMI_720x480P60_27027_RGB_8]       = new HDMI_Source(  720,  480, 60, false,  27027,   16, 62,  60, false,   9,  6, 30, false,   2,   8,  0, true );  // 480p_8_60Hz_CBar_RGB_TV_v54
        HDMI_source[(int) ESourceRes.HDMI_720x576P50_27000_RGB_8]       = new HDMI_Source(  720,  576, 50, false,  27000,   12, 64,  68, false,   5,  5, 39, false,  17,   8,  0, true );  // 576p_8_50Hz_CBar_RGB_TV_v54
        HDMI_source[(int) ESourceRes.HDMI_3840x2160P60_594000_YUV420_8] = new HDMI_Source( 3840, 2160, 60, false, 594000,  176, 88, 296, true,    8, 10, 72, true,   97,   8,  3, true );  // 3840x2160p_60Hz_ColorBar_420_VIC_97
        HDMI_source[(int) ESourceRes.HDMI_1280x720P50_92813_RGB_10]     = new HDMI_Source( 1280,  720, 50, false,  74250,  440, 40, 220, true,    5,  5, 20, true,   19,  10,  0, true );  // 720p_10_50_CBar_RGB_TV_v90
        HDMI_source[(int) ESourceRes.HDMI_1280x720P60_74250_RGB_8]      = new HDMI_Source( 1280,  720, 60, false,  74250,  110, 40, 220, true,    5,  5, 20, true,    4,   8,  0, true );  // 720p_8_60Hz_CBar_RGB_TV_v54
        HDMI_source[(int) ESourceRes.HDMI_1920x1080I50_74250_RGB_8]     = new HDMI_Source( 1920, 1080, 25, true,   74250,  528, 44, 148, true,    2,  5, 15, true,   20,   8,  0, true );  // 1080i_8_50Hz_CBar_RGB_TV_v54
        HDMI_source[(int) ESourceRes.HDMI_1920x1080I60_74250_RGB_8]     = new HDMI_Source( 1920, 1080, 30, true,   74250,   88, 44, 148, true,    2,  5, 15, true,    5,   8,  0, true );  // 1080i_8_60Hz_CBar_RGB_TV_v54 
        HDMI_source[(int) ESourceRes.HDMI_1920x1080P30_74250_RGB_8]     = new HDMI_Source( 1920, 1080, 30, false,  74250,   88, 44, 148, true,    4,  5, 36, true,   34,   8,  0, true );  // 1080p_8_30Hz_CBar_RGB_TV_v99
        HDMI_source[(int) ESourceRes.HDMI_1920x1080P50_148500_RGB_8]    = new HDMI_Source( 1920, 1080, 50, false, 148500,  528, 44, 148, true,    4,  5, 36, true,   31,   8,  0, true );  // 1080p_8_50Hz_CBar_RGB_TV_v63
        HDMI_source[(int) ESourceRes.HDMI_1920x1080P60_148500_RGB_8]    = new HDMI_Source( 1920, 1080, 60, false, 148500,   88, 44, 148, true,    4,  5, 36, true,   16,   8,  0, true );  // 1080p_8_60Hz_CBar_RGB_TV_v54
        HDMI_source[(int) ESourceRes.HDMI_1920x1080P60_185625_RGB_10]   = new HDMI_Source( 1920, 1080, 60, false, 148500,   88, 44, 148, true,    4,  5, 36, true,   16,  10,  0, true );  // 1080p_10_60_CBar_RGB_TV_v90   MMCM output is 148.5   MHz    LDetect give 148.493 MHz
        HDMI_source[(int) ESourceRes.HDMI_1920x1080P60_222750_RGB_12]   = new HDMI_Source( 1920, 1080, 60, false, 148333,   88, 44, 148, true,    4,  5, 36, true,   16,  12,  0, true );  // 1080p_12_60_CBar_RGB_TV_v90   MMCM output is 148.333 MHz    LDetect give 148.328 MHz
        HDMI_source[(int) ESourceRes.HDMI_3840x2160P24_297000_RGB_8]    = new HDMI_Source( 3840, 2160, 24, false, 297484, 1276, 88, 296, true,    8, 10, 72, true,    0,   8,  0, true );  // 4K2K_2160p_24Hz

        string oStepMsg = null;

        // Lists to hold the data (and errors) to diplay on screen...
        List<CStringEx> strDebugDisplayStrings = new List<CStringEx>();
        List<string> strErrorDisplayStrings = new List<string>();

        string strHdrLine;
        // string strBoardInfo;
        DateTime TestStartTime;

        if (bPromptForLogfile == true)
            LogFilename = GetFileToOpen();
        else
            LogFilename = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + DefaultLogFilename;

        TestStartTime = DateTime.Now;

        // Now print out the header so we know what's what...
        strHdrLine = "\r\nTest started at: " + TestStartTime.ToString() + "\r\n";
        strHdrLine += "Logging data to file: " + LogFilename + "\r\n";
        strHdrLine += "HDMI Monitor test with Nexys Source " + "\r\n";
        strHdrLine += "SVN rev : " + strSvnRevision.Trim('$') + "\r\n";
        strHdrLine += "SVN date: " + strSvnDate.Trim('$') + "\r\n";
        strHdrLine += "Test parameters : Sources (start=" + eSourceRes +", algo=" + eSourceChangeAlgo + ", source_mask=0x" + Convert.ToString(iSourceChangeMask, 16) + "\r\n";
        strHdrLine += "Test parameters : Defects (start=" + eDefectType +", algo=" + eDefectChangeAlgo + ", defect_mask=0x" + Convert.ToString(iDefectChangeMask, 16) + ", image_mask=0x" + Convert.ToString(iDefectImageMask, 16) +"\r\n";
        strHdrLine += "Test parameters : StopAfterErrorCount=" + iStopAfterErrorCount + ", Defect_Jitter_mTbit=" + iJitterAmplitude_mTbit;

        m_oHostOptions.AppendToDebug(strHdrLine, false, true);
        WriteLogFile2(LogFilename, strHdrLine);
        string StrLogLine = "";
        string StrLog = "";

        // Thread.Sleep(1000);

        // ================================================
        // Initial HDMI output setting
        // ================================================
        m_oHostOptions.AppendToDebug("Init Source HDMI.", false, true);
        DisableHdmi();
        InitRegistersHdmiGen();
        SetHdmiSource(eSourceRes);
        Thread.Sleep(1000);
        EnableHdmi();
        Thread.Sleep(5000);

        int DelayVariable = 0;

        string sOutput;
        int iFailCnt = 0;
        int iPassCnt = 0;
        int iNoCheckCnt = 0;
        do
        {
            m_oHostOptions.ShowProgress(iLoopIdx);
            m_oHostOptions.Section(iLoopIdx, "Testing DUT with Nexys Video HDMI source. (" + oStepMsg + ")");

            StrLogLine = "";

            // Date, Time, iteration and defect
            StrLog = DateTime.Now.ToString() + "  " + (iLoopIdx + 1) + " " + "D" + (int) eDefectType + " ";
            m_oHostOptions.AppendToDebug(StrLog, false, false);
            StrLogLine = StrLogLine + StrLog;

            // Source resolution info
            StrLog = ReadSourceInfo((int)eSourceRes, HDMI_source);
            m_oHostOptions.AppendToDebug(StrLog, false, false);
            StrLogLine = StrLogLine + StrLog;

            // Fresh Register Clear on Hdmi Generator
            InitRegistersHdmiGen();

            // HDMI transition generation (this will apply a defect, then turn on the output)
            HdmiTransition(eDefectType, eSourceRes, HDMI_source[(int)eSourceRes].PixelClock, HDMI_source[(int)eSourceRes].DC);

            // Wait LOOP_DELAY_MS in 100ms increment.
            for (int d = 0; d < LOOP_DELAY_MS; d += 100) // Note LOOP_DELAY_MS in increment of 100ms
            {
                if (bTerminateComplete)
                    break;
                Thread.Sleep(100);
            }

            // Do we need to check the image?
            bool bCheckImage = ((iDefectImageMask >> (int) eDefectType) & 0x1) != 0;

            // Take a webcamCam snapshot (Output file : image.bmp)
            int[] DetectColorbar = new int[2];
            int[] DetectBlack = new int[2];
            int[] DetectUnknown = new int[2];

            if (bCheckImage)
            {
                bool bSnapShotSuccess = false;
                while (bSnapShotSuccess == false)
                {
                    StrLog = " shoot";
                    m_oHostOptions.AppendToDebug(StrLog, false, false);
                    StrLogLine = StrLogLine + StrLog;

                    sOutput = CamSnapshot();
                    if (sOutput == "") // command was send correctly
                    {
                        bSnapShotSuccess = true;
                        StrLog = " OK  ";
                        m_oHostOptions.AppendToDebug(StrLog, false, false);
                        StrLogLine = StrLogLine + StrLog;
                    }
                    else
                    { // something went wrong
                        StrLog = " FAIL: " + sOutput; // show error
                        // StrLog = " FAIL ";
                        m_oHostOptions.AppendToDebug(StrLog, false, false);
                        StrLogLine = StrLogLine + StrLog;

                        // kill process
                        StrLog = " pskill CommandCam";
                        m_oHostOptions.AppendToDebug(StrLog, false, false);
                        StrLogLine = StrLogLine + StrLog;

                        Process oProcessK = new Process();
                        oProcessK.StartInfo.FileName = "pskill";
                        oProcessK.StartInfo.Arguments = "CommandCam";
                        oProcessK.Start();
                        oProcessK.WaitForExit();
                        Thread.Sleep(5000); // wait
                    }
                }

                // Image Processor Recognition
                StrLog = " Recog";
                m_oHostOptions.AppendToDebug(StrLog, false, false);
                StrLogLine = StrLogLine + StrLog;

                // List<string> oResponse = null;
                // string oResponse = "";
                string oResp = ImageProcessor();
                // Sample of the ImageProcessor output
                // ImageProcessor
                // ColorBar  95  94
                // Black      0   0
                // Unknown    5   6
                string[] subs = oResp.Split(
                    new char[] { ' ', '\n', '\r' },
                    StringSplitOptions.RemoveEmptyEntries
                );

                // m_oHostOptions.AppendToDebug("0 " + subs[0],false,true);  // ImageProcess
                // m_oHostOptions.AppendToDebug("1 " + subs[1],false,true);  // ColorBar
                // m_oHostOptions.AppendToDebug("2 " + subs[2],false,true);
                // m_oHostOptions.AppendToDebug("3 " + subs[3],false,true);
                // m_oHostOptions.AppendToDebug("4 " + subs[4],false,true);    // Black
                // m_oHostOptions.AppendToDebug("5 " + subs[5],false,true);
                // m_oHostOptions.AppendToDebug("6 " + subs[6],false,true);
                // m_oHostOptions.AppendToDebug("7 " + subs[7],false,true);    // Unknown
                // m_oHostOptions.AppendToDebug("8 " + subs[8],false,true);
                // m_oHostOptions.AppendToDebug("9 " + subs[9],false,true);

                DetectColorbar[0] = Int32.Parse(subs[2]);
                DetectColorbar[1] = Int32.Parse(subs[3]);
                DetectBlack[0] = Int32.Parse(subs[5]);
                DetectBlack[1] = Int32.Parse(subs[6]);
                DetectUnknown[0] = Int32.Parse(subs[8]);
                DetectUnknown[1] = Int32.Parse(subs[9]);

                // StrLog =  String.Format(" c{0,3}{1,3}", DetectColorbar[0], DetectColorbar[1]);
                // StrLog += String.Format(" b{0,3}{1,3}", DetectBlack[0], DetectBlack[1]);
                // StrLog += String.Format(" u{0,3}{1,3}", DetectUnknown[0], DetectUnknown[1]);
                StrLog =  String.Format(" c{0,3}", DetectColorbar[0]);
                StrLog += String.Format(" b{0,3}", DetectBlack[0]);
                StrLog += String.Format(" u{0,3}", DetectUnknown[0]);
                m_oHostOptions.AppendToDebug(StrLog, false, false);
                StrLogLine = StrLogLine + StrLog;
            }
            else
            {
                StrLog = " shoot --  ";
                m_oHostOptions.AppendToDebug(StrLog, false, false);
                StrLogLine = StrLogLine + StrLog;

                StrLog = " Recog c -- b -- u --";
                m_oHostOptions.AppendToDebug(StrLog, false, false);
                StrLogLine = StrLogLine + StrLog;
            }

            // Verify the DUT is responding to a ping
            bool bPingResultPass = false;
            if (bPing == true)
            {
                string sPingOutput = PingDut(IP_Address);
                // m_oHostOptions.AppendToDebug(sPingOutput,false,false);
                // StrLogLine = StrLogLine + sPingOutput;
                bPingResultPass = AnalysePingResult(sPingOutput);
                if (bPingResultPass == true)
                {
                    StrLog = " Ping OK  ";
                }
                else
                {
                    StrLog = " Ping FAIL";
                }
                m_oHostOptions.AppendToDebug(StrLog, false, false);
                StrLogLine = StrLogLine + StrLog;
            }

            // Final diagnostic
            const int iThreshold = 60;
            int iValue = (eDefectType == EDefectType.BLACK) ? DetectBlack[0] : DetectColorbar[0];

            if ( ( bPing       && !bPingResultPass      ) ||
                 ( bCheckImage && (iValue < iThreshold) ) )
            {
                iFailCnt += 1;
                iMonitorErrCount += 1;
                StrLog = " FAIL (Pass " + iPassCnt + ", NoCheck " + iNoCheckCnt + ", Fail " + iFailCnt + ", total " + (iLoopIdx+1) + ")";
                // Stop on error (if enabled)
                if (iStopAfterErrorCount > 0 && iFailCnt >= iStopAfterErrorCount)
                {
                    bTerminateComplete = true;
                    m_oHostOptions.AppendToDebug("\r\nTermination after " + iFailCnt + "errors... \r\n");
                }
            }
            else if (!bCheckImage)
            {
                iNoCheckCnt += 1;
                StrLog = " ---- (Pass " + iPassCnt + ", NoCheck " + iNoCheckCnt + ", Fail " + iFailCnt + ", total " + (iLoopIdx+1) + ")";
            }
            else
            {
                iPassCnt += 1;
                StrLog = " PASS (Pass " + iPassCnt + ", NoCheck " + iNoCheckCnt + ", Fail " + iFailCnt + ", total " + (iLoopIdx+1) + ")";
            }

            m_oHostOptions.AppendToDebug(StrLog, false, false);
            StrLogLine = StrLogLine + StrLog;

            // =====================================================
            // Determine next Source change
            // =====================================================
            eSourceRes = (ESourceRes)GetNextValue(
                eSourceChangeAlgo,
                (int)eSourceRes,
                (int)ESourceRes.COUNT,
                iSourceChangeMask
            );

            // =====================================================
            // Determine next defect on transition
            // =====================================================
            eDefectType = (EDefectType)GetNextValue(
                eDefectChangeAlgo,
                (int)eDefectType,
                (int)EDefectType.COUNT,
                iDefectChangeMask
            );

            // =====================================================
            // Check for terminate request
            // =====================================================
            lock (m_oSync)
            {
                if (bTerminateRequest)
                {
                    bTerminateComplete = true;
                    m_oHostOptions.AppendToDebug("\r\nTermination requested by user... \r\n");
                    break;
                }
            }

            // m_oHostOptions.AppendToDebug("End of passs...", false, true);
            m_oHostOptions.AppendToDebug("\r\n", false, false); // Only print CR LF
            WriteLogFile2(LogFilename, StrLogLine); // Writing to log file automatically add CR LF

            iLoopIdx++;
        } while (!bTerminateComplete);

        m_oHostOptions.AppendToDebug("", false);
        m_oHostOptions.AppendToDebug("TEST COMPLETE", false);
        WriteLogFile2(LogFilename, "\r\nTEST COMPLETE");
        m_oHostOptions.AppendToDebug("Test duration: " + (DateTime.Now-TestStartTime).ToString(@"d\.hh\:mm\:ss") + "\r\n\r\n",false);
        WriteLogFile2(LogFilename, "Test duration: " + (DateTime.Now-TestStartTime).ToString(@"d\.hh\:mm\:ss") + "\r\n\r\n");
    }
}
