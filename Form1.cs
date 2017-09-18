using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using HWPCONTROLLib;

namespace removename
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int countprog = 0, tmp =0;
            int mode = 1;
            //0 : 모든 문제파일 _p붙이기
            //1 : 글꼴 일괄변경
            //2 : 특정 유형 파일 삭제
            //3 : 수정된 날짜 기준 삭제
            //4 : 변경

            //5 : GEQ 수정

            //6 : 페이지 마다 자르기

            if (mode == 6)
            {

                int EXcount = 1;
                FileInfo[] files = null;
                string pathToDir = @"D:\proj\170907_you_Graphproblem\Graph";
                //@"D:\proj\170907_convert\List\hwp";
                //@"C:\Users\user1\Documents\equationTest";          
                // 
                //@"D:\proj\170906_convert\List";
                DirectoryInfo dirinfo = new DirectoryInfo(pathToDir);
                string path = dirinfo.FullName;
                files = dirinfo.GetFiles("*.*", SearchOption.AllDirectories);
                string[] fileList = Directory.GetFiles(pathToDir, "*", SearchOption.AllDirectories);

                object TypMissing = Type.Missing;
                Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();

                Excel.Workbook _workbook = null;

                // 파일이 존재 하면 열고, 없으면 새로 만듬.
                string pathexcel = @"D:\20170915_Graph_4.xls";
            try
                {
                    if (File.Exists(pathexcel))

                    {

                        _workbook = ExcelApp.Workbooks.Open(pathexcel, TypMissing, TypMissing, TypMissing, TypMissing,

                            TypMissing, TypMissing,

                            TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing, TypMissing);

                    }

                    else

                    {

                        // Add Work Book

                        _workbook = ExcelApp.Workbooks.Add(Type.Missing);

                        // Save Excel File

                        _workbook.SaveAs(pathexcel, Excel.XlFileFormat.xlWorkbookNormal, TypMissing, TypMissing,

                       TypMissing, TypMissing, Excel.XlSaveAsAccessMode.xlNoChange,

                       TypMissing, TypMissing, TypMissing, TypMissing, TypMissing);

                    }

                    Excel.Worksheet Sheet = (Excel.Worksheet)_workbook.Worksheets.get_Item("Sheet1");

                    Excel.Range Range_ = Sheet.get_Range("A1", Type.Missing);

                    /////////////////////////////////////////////////////////////
                    //////////////////// 엑셀 기본 틀////////////////////////////
                    /////////////////////////////////////////////////////////////
                    ;


                    for (int i = 0; i < fileList.Length; i++)
                    {
                        if (fileList[i].IndexOf("P") >= 0)
                        {
                           
                            HWPCONTROLLib.DHwpAction Aact;
                            HWPCONTROLLib.DHwpParameterSet Aset;

                            axHwpCtrl1.Open(fileList[i]);
                            Console.WriteLine(axHwpCtrl1.PageCount);

                            axHwpCtrl2.Open(fileList[i].Replace("_P", "_A"));
                            axHwpCtrl3.Open(fileList[i].Replace("_P", "_S"));

                            int prob_count = axHwpCtrl1.PageCount;
                            int solu_count = axHwpCtrl2.PageCount;
                            int answ_count = axHwpCtrl3.PageCount;

                            if (prob_count == solu_count && prob_count == answ_count && answ_count == solu_count)
                            {
                                string You = fileList[i].Substring(fileList[i].LastIndexOf(@"\") + 1, 9);

                               
                                for (int j = 0; j < prob_count; j++)
                                {
                                    
                                    string chars = "abcdefghijklmnopqrstuvwxyz0123456789";
                                    char[] charArray = chars.ToCharArray();

                                    string chars2 = "abcdefghijklmnopqrstuvwxyz0123456789";
                                    char[] charArray2 = chars2.ToCharArray();

                                    int numPasswd = 5;       // 5자 출력
                                    string newPasswd = string.Empty;
                                    int seed = Environment.TickCount;
                                    Random rd = new Random(seed);
                                    int tempNum = 0;

                                    for (int k = 0; k < numPasswd; k++)
                                    {
                                        tempNum = rd.Next(0, charArray.Length - 1);
                                        newPasswd += charArray[tempNum];
                                    }

                                    int numPasswd2 = 3;       // 3자 출력
                                    string newPasswd2 = string.Empty;
                                    int seed2 = Environment.TickCount;
                                    Random rd2 = new Random(seed2);
                                    int tempNum2 = 0;
                                    for (int k = 0; k < numPasswd2; k++)
                                    {
                                        tempNum2 = rd.Next(0, charArray2.Length - 1);
                                        newPasswd2 += charArray2[tempNum2];
                                    }


                                    int tmptest = fileList[i].IndexOf(".");
                                    string savefold = fileList[i].Substring(0, tmptest);

                                    string youfold = fileList[i].Substring(41, 6);

                                    string PPAP = pathToDir + @"\HWP1\9\m\1\1\" + youfold + @"\";
                                    DirectoryInfo dihwp = new DirectoryInfo(PPAP);
                                    if (dihwp.Exists == false)
                                    {
                                        dihwp.Create();
                                    }
                                    DirectoryInfo diPNG = new DirectoryInfo(PPAP.Replace("HWP1", "ng"));
                                    if (diPNG.Exists == false)
                                    {
                                        diPNG.Create();
                                    }
                                    DirectoryInfo diPNG_300 = new DirectoryInfo(PPAP.Replace("HWP1", "d"));
                                    if (diPNG_300.Exists == false)
                                    {
                                        diPNG_300.Create();
                                    }
                                    //pro
                                    if(j < prob_count-1)
                                    {
                                        axHwpCtrl1.Run("MoveDocBegin");
                                        Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("SaveBlockAction");
                                        Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();
                                        Aact.GetDefault(Aset);
                                        axHwpCtrl1.Run("MoveSelPageDown");

                                        string fullYou = "15_211" + youfold; //15_풀 유형 코드

                                        int tmpn = j + 1;

                                        Aset.SetItem("FileName", PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_p.hwp");
                                        Aset.SetItem("Format", "HWP");
                                        Aact.Execute(Aset);
                                        axHwpCtrl1.Run("Delete");
                                        Console.WriteLine(PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_p.hwp");
                                        Sheet.Cells[EXcount, 1] = youfold;
                                        Sheet.Cells[EXcount, 2] = @"/math_problems/9/m/1/1/" + youfold.Replace(@"\", "/") + @"/" + fullYou + "_" + newPasswd + "_" + newPasswd2;
                                        

                                        EXcount += 1;

                                        //ans
                                        axHwpCtrl2.Run("MoveDocBegin");
                                        Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl2.CreateAction("SaveBlockAction");
                                        Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();
                                        Aact.GetDefault(Aset);
                                        axHwpCtrl2.Run("MoveSelPageDown");

                                        Aset.SetItem("FileName", PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_a.hwp");
                                        Console.WriteLine(PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_a.hwp");
                                        Aset.SetItem("Format", "HWP");
                                        Aact.Execute(Aset);
                                        axHwpCtrl2.Run("Delete");



                                        //sol
                                        axHwpCtrl3.Run("MoveDocBegin");
                                        Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl3.CreateAction("SaveBlockAction");
                                        Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();
                                        Aact.GetDefault(Aset);
                                        axHwpCtrl3.Run("MoveSelPageDown");


                                        Aset.SetItem("FileName", PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_s.hwp");
                                        Console.WriteLine(PPAP + PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_s.hwp");
                                        Aset.SetItem("Format", "HWP");
                                        Aact.Execute(Aset);
                                        axHwpCtrl3.Run("Delete");
                                    }
                                    else
                                    {




                                        Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("SaveBlockAction");
                                        Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();
                                        Aact.GetDefault(Aset);

                                        string fullYou = "15_211" + youfold; //15_풀 유형 코드

                                        int tmpn = j + 1;

                                        axHwpCtrl1.Run("MoveDocBegin");
                                        axHwpCtrl1.Run("SelectAll");
                                        Aset.SetItem("FileName", PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_p.hwp");
                                        Aset.SetItem("Format", "HWP");
                                        Aact.Execute(Aset);
                                        axHwpCtrl1.Run("Delete");
                                        Sheet.Cells[EXcount, 1] = youfold;
                                        Sheet.Cells[EXcount, 2] = @"/math_problems/9/m/1/1/" + youfold.Replace(@"\", "/") + @"/" + fullYou + "_" + newPasswd + "_" + newPasswd2;
                                        Console.WriteLine("마지막 장"+  PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_p.hwp");

                                        EXcount += 1;

                                        //ans
                                        axHwpCtrl2.Run("MoveDocBegin");
                                        Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl2.CreateAction("SaveBlockAction");
                                        Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();
                                        Aact.GetDefault(Aset);
                                        axHwpCtrl2.Run("SelectAll");

                                        Aset.SetItem("FileName", PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_a.hwp");
                                        Aset.SetItem("Format", "HWP");
                                        Aact.Execute(Aset);
                                        axHwpCtrl2.Run("Delete");



                                        //sol
                                        axHwpCtrl3.Run("MoveDocBegin");
                                        Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl3.CreateAction("SaveBlockAction");
                                        Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();
                                        Aact.GetDefault(Aset);
                                        axHwpCtrl3.Run("SelectAll");


                                        Aset.SetItem("FileName", PPAP + fullYou + "_" + newPasswd + "_" + newPasswd2 + "_s.hwp");
                                        Aset.SetItem("Format", "HWP");
                                        Aact.Execute(Aset);
                                        axHwpCtrl3.Run("Delete");
                                    }



                                }
                            }
                            else Console.WriteLine("요놈 요놈 "+fileList[i]);


                        }




                    }
                }
                catch(System.IndexOutOfRangeException)
                {

                }
                _workbook.Save();
                ExcelApp.Workbooks.Close();

                ExcelApp.Quit();



            }


            if (mode == 5)
            {
                FileInfo[] files = null;
                string pathToDir = @"D:\proj\170907_convert\List";
                //@"D:\proj\170907_convert\List\hwp";
                //@"C:\Users\user1\Documents\equationTest";          
                // 
                //@"D:\proj\170906_convert\List";
                DirectoryInfo dirinfo = new DirectoryInfo(pathToDir);
                string path = dirinfo.FullName;
                files = dirinfo.GetFiles("*.*", SearchOption.AllDirectories);
                string[] fileList = Directory.GetFiles(pathToDir, "*", SearchOption.AllDirectories);



                for (int i = 48700; i < fileList.Length - 1; i++)
                {
                    HWPCONTROLLib.DHwpAction Aact; //수식
                    HWPCONTROLLib.DHwpParameterSet Aset;

                    HWPCONTROLLib.DHwpAction Bact;
                    HWPCONTROLLib.DHwpParameterSet Bset;


                    //수식 ( x->0 수정)
                    axHwpCtrl1.Open(fileList[i]);


                    axHwpCtrl1.Run("MoveDocBegin");

                    axHwpCtrl1.Run("SelectCtrlFront");




                    Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("EquationModify");
                    Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();

                    Aact.GetDefault(Aset);

                    Boolean CC = Aset.ItemExist("String");

                    String KK = (String)Aset.Item("String");

                    if (CC && KK.IndexOf("LEQ") >= 0)
                    {
                        Console.WriteLine("파일명 : " + fileList[i] + "     문제내용   :      " + KK);
                        Aset.SetItem("String", KK.Replace("LEQ", " LEQ "));
                    }
                    else if (CC && KK.IndexOf(">=") >= 0)
                    {
                        Console.WriteLine("파일명 : " + fileList[i] + "     문제내용   :      " + KK);
                        Aset.SetItem("String", KK.Replace("GEQ", " GEQ "));
                    }

                    Aact.Execute(Aset);

                    if (i % 100 == 0)
                        Console.WriteLine(i);
                }
            }

            else if (mode == 4)
            {
                FileInfo[] files = null;
                string pathToDir = @"D:\proj\170907_convert\List";
                //@"D:\proj\170907_convert\List\hwp";
                //@"C:\Users\user1\Documents\equationTest";          
                // 
                //@"D:\proj\170906_convert\List";
                DirectoryInfo dirinfo = new DirectoryInfo(pathToDir);
                string path = dirinfo.FullName;
                files = dirinfo.GetFiles("*.*", SearchOption.AllDirectories);
                string[] fileList = Directory.GetFiles(pathToDir, "*", SearchOption.AllDirectories);



                for (int i = 0; i < fileList.Length - 1; i++)
                {
                    HWPCONTROLLib.DHwpAction Aact; //수식
                    HWPCONTROLLib.DHwpParameterSet Aset;

                    HWPCONTROLLib.DHwpAction Bact;
                    HWPCONTROLLib.DHwpParameterSet Bset;


                    //수식 ( x->0 수정)
                    axHwpCtrl1.Open(fileList[i]);


                    axHwpCtrl1.Run("MoveDocBegin");

                    axHwpCtrl1.Run("SelectCtrlFront");




                    Aact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("EquationModify");
                    Aset = (HWPCONTROLLib.DHwpParameterSet)Aact.CreateSet();

                    Aact.GetDefault(Aset);

                    Boolean CC = Aset.ItemExist("String");

                    String KK = (String)Aset.Item("String");


                    if (CC && KK.IndexOf("im") >= 0)
                    {
                        if (KK.IndexOf("+0}") > 0)
                        {
                            Console.WriteLine("파일명 : " + fileList[i] + "     문제내용   :      " + KK);
                            Aset.SetItem("String", KK.Replace("+0}", "+}"));


                        }

                        else if (KK.IndexOf("-0}") >= 0)
                        {
                            Console.WriteLine("파일명 : " + fileList[i] + "     문제내용   :      " + KK);
                            Aset.SetItem("String", KK.Replace("-0}", "-}"));
                        }
                    }
                    else if (CC && KK.IndexOf("<=") >= 0)
                    {
                        Console.WriteLine("파일명 : " + fileList[i] + "     문제내용   :      " + KK);
                        Aset.SetItem("String", KK.Replace("<=", "LEQ"));
                    }
                    else if (CC && KK.IndexOf(">=") >= 0)
                    {
                        Console.WriteLine("파일명 : " + fileList[i] + "     문제내용   :      " + KK);
                        Aset.SetItem("String", KK.Replace(">=", "GEQ"));
                    }

                    Aact.Execute(Aset);


                    axHwpCtrl1.InitScan();
                    string bodyStr = axHwpCtrl1.GetPageText(0);
                    axHwpCtrl1.ReleaseScan();

                    if (bodyStr.IndexOf("실수부") > 0 || bodyStr.IndexOf("허수부") > 0 || bodyStr.IndexOf("절대값") > 0 || bodyStr.IndexOf("극소값") > 0 ||
                        bodyStr.IndexOf("극대값") > 0 || bodyStr.IndexOf("함수값") > 0 || bodyStr.IndexOf("중간값") > 0 || bodyStr.IndexOf("꼭지점") > 0 ||
                        bodyStr.IndexOf("개구간") > 0 || bodyStr.IndexOf("폐구간") > 0 || bodyStr.IndexOf("근사값") > 0 || bodyStr.IndexOf("지표") > 0 ||
                        bodyStr.IndexOf("가수") > 0 || bodyStr.IndexOf("≦") > 0 || bodyStr.IndexOf("≧") > 0 ||
                        bodyStr.IndexOf("최대값") > 0 || bodyStr.IndexOf("개수") > 0 || bodyStr.IndexOf("최소값") > 0)
                    {
                        Console.WriteLine("파일명 : " + fileList[i] + "     문제내용   :      " + bodyStr);
                    }

                    // " 실수부 -> 실수부분 "
                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "실수부");

                    Bset.SetItem("ReplaceString", "실수부분");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");
                    string on = "on";
                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");


                    Bact.Execute(Bset);
                    // 오류 체크 (실수부분분)
                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "실수부분분");
                    Bset.SetItem("ReplaceString", "실수부분");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");


                    Bact.Execute(Bset);


                    // " 허수부 -> 허수부분 "
                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "허수부");
                    Bset.SetItem("ReplaceString", "허수부분");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");


                    Bact.Execute(Bset);


                    //오류체크 허수부분분

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "허수부분분");
                    Bset.SetItem("ReplaceString", "허수부분");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");


                    Bact.Execute(Bset);


                    //절대값

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "절대값");
                    Bset.SetItem("ReplaceString", "절댓값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");


                    Bact.Execute(Bset);

                    //함수값


                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "함수값");
                    Bset.SetItem("ReplaceString", "함숫값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");


                    Bact.Execute(Bset);
                    //대표값 -> 대푯값
                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "대표값");
                    Bset.SetItem("ReplaceString", "대푯값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");


                    Bact.Execute(Bset);

                    //극대값 - > 극 댓 값

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "극대값");
                    Bset.SetItem("ReplaceString", "극댓값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);
                    // 극소값 -> 극솟값

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "극소값");
                    Bset.SetItem("ReplaceString", "극솟값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);


                    // 최소값 -> 최솟값

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "최소값");
                    Bset.SetItem("ReplaceString", "최솟값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);
                    // 기대값 -> 기댓값

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "기대값");
                    Bset.SetItem("ReplaceString", "기댓값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    //꼭지점 -> 꼭짓점


                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "꼭지점");
                    Bset.SetItem("ReplaceString", "꼭짓점");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    //최대값 -> 최댓값


                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "최대값");
                    Bset.SetItem("ReplaceString", "최댓값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    // 근사값 -> 근삿값


                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "근사값");
                    Bset.SetItem("ReplaceString", "근삿값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    //지표 -> 정수 부분

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "지표");
                    Bset.SetItem("ReplaceString", "정수 부분");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    //가수 -> 소수 부분


                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "가수");
                    Bset.SetItem("ReplaceString", "소수 부분");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    //중간값 -> 사이값
                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "중간값");
                    Bset.SetItem("ReplaceString", "사이값");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    // 개구간 -> 열린 구간


                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "개구간");
                    Bset.SetItem("ReplaceString", "열린 구간");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);
                    //개수 -> 갯수

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "개수");
                    Bset.SetItem("ReplaceString", "갯수");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    //폐구간 -> 닫힌 구간


                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "폐구간");
                    Bset.SetItem("ReplaceString", "닫힌 구간");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    // 기호 1

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "≦");
                    Bset.SetItem("ReplaceString", "≤");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);

                    //기호 2

                    axHwpCtrl1.Run("MoveDocBegin");

                    Bact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("AllReplace");
                    Bset = (HWPCONTROLLib.DHwpParameterSet)Bact.CreateSet();
                    Bact.GetDefault(Bset);

                    Bset.SetItem("FindString", "≧");
                    Bset.SetItem("ReplaceString", "≥");
                    Bset.SetItem("Direction", 2);
                    Bset.SetItem("AllWordForms", "on");
                    Bset.SetItem("SeveralWords", "on");
                    Bset.SetItem("UseWildCards", "off");

                    Bset.SetItem("WholeWordOnly", "off");
                    Bset.SetItem("AutoSpell", "off");
                    Bset.SetItem("ReplaceMode", "on");
                    Bset.SetItem("IgnoreFindString", "off");
                    Bset.SetItem("IgnoreReplaceString", "off");
                    //Bset.SetItem("FindCharShape","");
                    //Bset.SetItem("FindParaShape","" );
                    //Bset.SetItem("ReplaceCharShape", "");
                    //Bset.SetItem("ReplaceParaShape", "");
                    //Bset.SetItem("FindStyle", "");
                    //Bset.SetItem("ReplaceStyle", "");

                    Bset.SetItem("IgnoreMessage", 1);
                    Bset.SetItem("HanjaFromHangul", "off");
                    Bset.SetItem("FindJaso", "off");
                    Bset.SetItem("FindRegExp", "off");
                    Bset.SetItem("FindType", "True");

                    Bact.Execute(Bset);


                    axHwpCtrl1.Run("HwpCtrlFileSave");

                    if (i % 100 == 0)
                        Console.WriteLine(i);

                }

            }
            if (mode == 3)
            {
                int k = 0;
                FileInfo[] files = null;
                string pathToDir = "D:\\mo_last\\PNG_300_Complete2";
                DirectoryInfo dirinfo = new DirectoryInfo(pathToDir);
                string path = dirinfo.FullName;
                files = dirinfo.GetFiles("*.*", SearchOption.AllDirectories);
                
                string[] fileList = Directory.GetFiles(pathToDir, "*", SearchOption.AllDirectories);
                
               
                    foreach(DirectoryInfo dir2 in dirinfo.GetDirectories())
                    {
                        foreach(DirectoryInfo dir3 in dir2.GetDirectories())
                        {
                            foreach(DirectoryInfo dir4 in dir3.GetDirectories())
                            {
                                foreach (var item in dir4.GetFiles())
                                {
                                    if(item.CreationTime < DateTime.Today.AddDays(-2))
                                    {
                                        Console.WriteLine(item.CreationTime);
                                        item.Delete();
                                    }

                                }
                            }

                        }

                    }

                }

            
            else if (mode == 0)
            {
                FileInfo[] files = null;
                string pathToDir = "D:\\mo_last\\HWP_Complete2\\MO_201609";
                DirectoryInfo dirinfo = new DirectoryInfo(pathToDir);
                string path = dirinfo.FullName;
                files = dirinfo.GetFiles("*.*", SearchOption.AllDirectories);

                string[] fileList = Directory.GetFiles(pathToDir, "*", SearchOption.AllDirectories);

                int k = 0;
                for (int i = 0; i < fileList.Length - 1; i++)
                {
                    if (fileList[i].IndexOf("_a") == -1 && fileList[i].IndexOf("_s") == -1)
                    {

                        int lastindex = fileList[i].LastIndexOf("\\");

                        System.IO.File.Move(fileList[i], fileList[i].Replace(".", "p."));
                    }

                }
            }
            else if(mode == 2)
            {
                FileInfo[] files = null;
                string pathToDir = "D:\\mo_last\\HWP_Complete2\\";
                DirectoryInfo dirinfo = new DirectoryInfo(pathToDir);
                string path = dirinfo.FullName;
                files = dirinfo.GetFiles("*.*", SearchOption.AllDirectories);

                string[] fileList = Directory.GetFiles(pathToDir, "*", SearchOption.AllDirectories);

                int k = 0;
                for (int i = 0; i < fileList.Length - 1; i++)
                {
                    if (fileList[i].IndexOf(".png") > 0 )
                    {
                        System.IO.File.Delete(fileList[i]);
                        Console.WriteLine(fileList[i]);
                    }

                }
            }
            else if (mode == 1)
            {
                int k = 0;
                FileInfo[] files = null;
                string pathToDir = @"D:\proj\170914_Convert\List\imhw";
                DirectoryInfo dirinfo = new DirectoryInfo(pathToDir);
                string path = dirinfo.FullName;
                files = dirinfo.GetFiles("*.hwp*", SearchOption.AllDirectories);
                string[] fileList = Directory.GetFiles(pathToDir, "*", SearchOption.AllDirectories);

                for (int i = 1846; i < fileList.Length - 1; i++)
                {
                    if (fileList[i].IndexOf(".hwp") > 0)// && fileList[i].IndexOf("201006") > 0 &&fileList[i].IndexOf("Ha_B")>0&& fileList[i].IndexOf("_a") > 0)
                    {
                     //   axHwpCtrl1.Open(fileList[i]);

                        /*

                        axHwpCtrl1.Run("SelectAll");
                        HWPCONTROLLib.DHwpAction pact = (HWPCONTROLLib.DHwpAction)axHwpCtrl1.CreateAction("CharShape");
                        HWPCONTROLLib.DHwpParameterSet pset = (HWPCONTROLLib.DHwpParameterSet)pact.CreateSet();
                        pact.GetDefault(pset);
                        string font = "나눔고딕";
                        pset.SetItem("FaceNameHangul", font);
                        pset.SetItem("FaceNameLatin", font);
                        pset.SetItem("FaceNameHanja", font);
                        pset.SetItem("FaceNameJapanese", font);
                        pset.SetItem("FaceNameOther", font);
                        pset.SetItem("FaceNameSymbol", font);
                        pset.SetItem("FaceNameUser", font);
                        pset.SetItem("Height", 1000);
                        pset.SetItem("TextColor", 0x0333333);
                        pset.SetItem("Bold", 1);
                        pact.Execute(pset);
                        // 
                        axHwpCtrl1.Run("HwpCtrlFileSave");
                        


                        ////
                        */
                       // axHwpCtrl1.CreatePageImage(fileList[i], 0, 558, 24, "gif");
                        
                        string path_Name = Path.GetFileNameWithoutExtension(fileList[i]);
                        string HwpFolderP, PngFolderP;
                        HwpFolderP = Path.GetDirectoryName(fileList[i]);
                        PngFolderP = HwpFolderP.Replace("imhw", "PNG");
                        DirectoryInfo PngFolder = new DirectoryInfo(PngFolderP);
                        string replacepathGIF = PngFolder + "\\" + path_Name + ".png";

                        Bitmap newimage = new Bitmap(fileList[i].Replace(".hwp",".gif"));

                        int pix_x = newimage.Width;
                        int pix_y = newimage.Height;
                        int memory_Left = 0;
                        int memory_Up = 0;
                        int memory_Down = 2;
                        int memory_Right = 2;



                        Boolean start = false;
                        Boolean finish = false;
                        System.Drawing.Graphics g = null;

                        for (int x = pix_x - 1; x >= 0; x--)
                        {
                            if (!start)
                            {
                                start = true;

                            }

                            Boolean exit = false;


                            for (int y = 0; y < pix_y; y++)
                            {

                                Color _pixel = newimage.GetPixel(x, y);

                                int clr = _pixel.ToArgb();
                                int red = (clr & 0x00ff0000) >> 16;
                                int green = (clr & 0x0000ff00) >> 8;
                                int blue = clr & 0x000000ff;



                                if (!((red == 255 || red == 254) && (green == 255 || green == 254) && (blue == 255 || blue == 254)))
                                {
                                 
                                    exit = true;


                                    memory_Right = x + 1;
                                    break;


                                }


                            }

                            if (exit)
                            {
                                break;
                            }

                        }
                        

                        //왼쪽
                        for (int x = 0; x < memory_Right; x++)
                        {
                            Boolean exit = false;


                            for (int y = 0; y < pix_y; y++)
                            {

                                Color _pixel = newimage.GetPixel(x, y);

                                int clr = _pixel.ToArgb();
                                int red = (clr & 0x00ff0000) >> 16;
                                int green = (clr & 0x0000ff00) >> 8;
                                int blue = clr & 0x000000ff;

                                if (!((red == 255 || red == 254) && (green == 255 || green == 254) && (blue == 255 || blue == 254)))
                                {
                                    exit = true;
                                    memory_Left = x - 1;
                                    break;
                                }


                            }

                            if (exit)
                            {
                                break;
                            }

                        }
                       



                        //아래
                        for (int y = pix_y - 1; y >= 0; y--)
                        {
                            Boolean exit = false;

                            for (int x = memory_Left; x < memory_Right; x++)
                            {

                                Color _pixel = newimage.GetPixel(x, y);

                                int clr = _pixel.ToArgb();
                                int red = (clr & 0x00ff0000) >> 16;
                                int green = (clr & 0x0000ff00) >> 8;
                                int blue = clr & 0x000000ff;

                                if (!((red == 255 || red == 254) && (green == 255 || green == 254) && (blue == 255 || blue == 254)))
                                {

                                    exit = true;
                                    memory_Down = y + 1;
                                    break;

                                }


                            }

                            if (exit)
                            {
                                break;
                            }

                        }


                        for (int y = 0; y < memory_Down; y++)
                        {
                            Boolean exit = false;

                            for (int x = memory_Left; x < memory_Right; x++)
                            {
                                Color _pixel = newimage.GetPixel(x, y);

                                int clr = _pixel.ToArgb();
                                int red = (clr & 0x00ff0000) >> 16;
                                int green = (clr & 0x0000ff00) >> 8;
                                int blue = clr & 0x000000ff;

                                if (!((red == 255 || red == 254) && (green == 255 || green == 254) && (blue == 255 || blue == 254)))
                                {
                                    finish = true;
                                    exit = true;
                                    memory_Up = y - 1;
                                    break;

                                }


                            }


                            if (exit)
                            {
                                break;
                            }

                        }
                        if (finish)
                        {


                            Bitmap CroppedImage = newimage.Clone(new System.Drawing.Rectangle(memory_Left, memory_Up, (memory_Right - memory_Left), (memory_Down - memory_Up)), newimage.PixelFormat);
                            int width = CroppedImage.Width / 2 * 103 / 100;
                            int height = CroppedImage.Height / 2 * 103 / 100;
                            Size resize = new Size(width, height);
                            Bitmap resizeImage = new Bitmap(CroppedImage, resize);
                            resizeImage.Save(replacepathGIF);
                            resizeImage.Dispose();
                            CroppedImage.Dispose();

                        }
                        
    

                        


                    }
                   
                        Console.WriteLine(i);
                    Console.WriteLine(fileList[i]);
                }
            }


        }

    }
}
