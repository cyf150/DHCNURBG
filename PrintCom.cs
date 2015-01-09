using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Xml;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using System.Drawing;
namespace PrintComm
{
    public class PrintCom
    {
        //常用入参20140815
        public string WebUrl = "";    //必须 webservice地址"http://10.160.16.30/dthealth/web/DWR.DoctorRound.cls";
        public string ItmName = "";   //必须 打印模板关键字
        public String TitleStr = "";  //表格必须 病人基本信息：web.DHCMGNurComm：Patinfo
        public string tranflag = "0"; //是否启用转科开关，默认不启用：1 启用；0 不启用
        public string PrnLoc = "";    //表格必须 打印时当前病人科室
        public string PrnBed = "";    //转科必须  打印时传入的病人当前床号
        public string SplitPage = "1";//转科打印时是否换页立即换页，默认换页 ---1 换页  0 不换页
        public int dxflag = 1;        //打印内容线条类型，默认是一条记录打印一直线：1 一条记录一条线；0 一行文字打印一条线 ;青岛是0
        public string xuflag = "0";   //续打开关，默认不许打-- 0：不续打；1:启用续打功能
        public string previewPrint = "0";             //预览打印 设置起始页和打印机 0 不弹出 1 弹出
        public int curPages = 0;      //打印第一页的页码
        public string AllLine = "Y";  //外框打印，Y--打印所有，其他--按内容行高打印
        public string ShowLocTran = "N";              //转科时科室名称是否显示转科记录：前一科->后一科  Y:是，N 否  2014.10.23
        public string BlueString = "24h出入液量统计"; //根据文字匹配下划线蓝色加粗20141121
        public string UserPrintDown = "Y"; //签名靠下打印:Y ; 靠上打印：N  //20150106
        public string PrintCareDateLine = "Y";  //日期时间列行数据为空时是否打印横线,Y-打印，N-不打印
       

        //生成图片配置
        public string StartMakePic = "N";   //开始
        public string MakeTemp = "N";       //是否生成图片
        public string IfUpload = "Y";       //生成的图片是否上传ftp
        public string filepath = "";        //生成图片配置xml地址
        public string MakeAllPages = "";    //是否生成所有页 ，默认是否(从原来最后一页开始生成)；"Y"--从第一页开始生成
        public string NurRecId = "";        //多次填写评估单id
        private string ftppath = "";        //服务器地址
        private string ftpuer = "";         //用户
        private string ftppwd = "";         //密码   
        private string ftpport = "";        //ftp端口
        private string ftpdealyTim = "";    //时间


        //CA打印相关
        public string CAStart = "0";       //默认不启用CA打印 1：启用；0：不启用
        public string IsVerifyCALoc = "0"; //科室是否启用CA--0 不启用；1 启用
        public int qmwildth = 60;          //电子签名图片宽
        public int qmheight = 12;          //电子签名图片高
        public int qmleft = 2;             //电子签名图片左边距
        public int qmtop = 2;              //电子签名图片上边距
        public int qmhori = 10;            //电子签名图片水平打印时两个图片间隔
        public int qmport = 2;             //电子签名图片垂直打印时图片间隔
        public int qmprnorientation = 1;   //多个电子签名图片打印方向1 横向 0 纵向
        public string blackflag = "N";     //是否打印黑色签名:"Y":打印黑色签名 ;青岛是N

        private int CareDateWidth = 0; //日期列宽度
        private int CareTimeWidth = 0; //时间列宽度

        public String RHeadCaption = "";   //RHead标签显示的内容
        public String LHeadCaption = "";   //LHead标签显示的内容
        public String RFootCaption = "";   //RFoot标签显示的内容        
        public String LFootCaption = "";   //LFoot标签显示的内容
        public String PreView = "1";
        public String EmrCode = "";
        public int cx = 0;
        public int cy = 0;
        public string ItmText = "";
        private Hashtable HashWinItm = new Hashtable(); // ///// 界面 Itm 
        public static int selShapIndex = -1;
        public string DesignFlag = "0";
        private int numprnflag = 0; //统计打印模板中RecLoc,RecBed,NextPageFlag数
        private Hashtable nextpagehastable = new Hashtable();
        public String printername = "";
        private Hashtable tcoldata = new Hashtable();
        private Hashtable tcolbakdata = new Hashtable();
        private Hashtable dzqmprnpgd = new Hashtable(); //电子签名评估单签名打印      
        public string isputong = "1"; //打印数据是从表Nur.DHCNurseRecSub取的：1 ；否：0
        //private String[] HFCaption=new String[10] ;
        private Hashtable HFCaption = new Hashtable();
        private int ItmCount = 2; //打印计数 
        private int PrnCount = 2;
        private int PagOffY = 0;
        public int stPage = 0;
        public int stRow = 0;  //起始行
        private int YLocation = 0;
        public int Pages = 0;  //页数
        public int stPrintPos = 0; //起始打印位置
        public string CareDateTim = "";//记录日期时间
        public int PrnFlag = 0;
        private Hashtable prnData = new Hashtable(); //打印数据
        private Hashtable DataTyp = new Hashtable();
        public XmlDocument xmlprndoc = new XmlDocument();
        public PrintDialog PrnDiaglog;
        public string ID = ""; //元素id值
        public string MultID = "";
        public string EpisodeID = "";
        public string LogonLoc = "";
        public string LogonUser = "";
        public string curhead = "";  //表头变更当前表头id --20141201
        public string Firsthead = "";  //表头变更打印时传入的第一个表头
        public PrintDocument pDoc;
        public PrintPreviewDialog PrnPreView;
        public string DataType = ""; //xml  qt //数据来源
        public string SourceFlag = ""; //Record，Method
        public string DataCls = "", DataMth = "", Parrm = "";
        public string[] GParrm = new string[16];
        public Hashtable DataHash = new Hashtable(); //存数据
        public Hashtable DataLable = new Hashtable(); //存数据  //空白栏
        public string ShapePrn = ""; //是否套打
        public int tabx = 0, taby = 0;
        public int tabW = 0, tabH = 0;
        public int Row = 0;
        public int PageRows = 0;  //每页打印的行数
        public string MthArr = ""; //方法组
        public DataTable table;
        public DataTable BakTable = new DataTable(); //备份数据
        public string[] head = null;
        private string ChangePageDiag = ""; //换页诊断
        private string ConnectStr = "";
        public string LabHead = "";  //表头变更打印标题
        private int StP = 0, EdP = 0; //起始页
        private bool showPrintDialog = false;
        public bool SelPrintFlag = false;
        //每页记录数 
        private int printPagesize = 20;
        private int PrnSumRows = 0; //总行数
        private int printIndex = 0;
        private Hashtable HPagRow = new Hashtable(); //每页行数记录
        private Hashtable HPagRow1 = new Hashtable(); //每页行数记录
        private DataTable tableBak; //备份数据
        private Hashtable HLhead = new Hashtable();  //记录标题
        private Hashtable HLasthead = new Hashtable();  //shenyang
        private Hashtable BakRowH = new Hashtable(); //记录行高
        //打印总页数 
        private int printpagecount = 0;
        private Hashtable PageLine = new Hashtable();   //记录统计总量线
        private XmlNode PageDiagNod = null; //诊断需要变换
        private XmlNode PageLocNod = null; //科室需要变换
        private XmlNode PageBedNod = null; //床号需要变换
        private Hashtable PageDiagH = new Hashtable();
        private Hashtable PageLocH = new Hashtable();
        private Hashtable PageBedH = new Hashtable();
        private PrintCommCache.DHCTranStream datastream = new PrintComm.PrintCommCache.DHCTranStream();  //流对象
        private DataSet CommData = new DataSet();
        /// 页边 //anhui20120515
        private int MargLeft = 0;  //左边
        private int MargTop = 0;   //右边

        public PageSetupDialog pageSetupDialog;
        private Boolean PageSet = false;
        private PageSettings oldPageSetting;
        private int MargRight = 0;   //右边
        private int MargBottom = 0;   //下边距
        private int Pageheight = 0;//页面高

        private int SetLeft = 0;  //设置左边距
        private int SetRight = 0; //设置右边距
        private int SetTop = 0;   //设置上边距
        private int SetBottom = 0; //设置下边距
        private int SetHeight = 0;//设置纸张高度
        private bool SetLandscape = false;
        private PaperSize SetPageType = null; //设置纸张名称

        private Hashtable BottomArray = new Hashtable(); //保存模板中高度在下边距内的元素
        private int InitTop = 0;
        private int InitBottom = 0;
        private int InitHeight = 0;
        private bool InitLandscape = false;


        public Hashtable rowidha = new Hashtable(); //续打 ，记录每行rowid
        private string printinfo = ""; //续打 续打信息
        public string lastprninfo = "";//续打，上次续打信息
        public int Startrow = 0; //续打
        public int Startpage = 0; //续打
        private ArrayList PgdPrintedArray = new ArrayList(); //评估单已打印元素
        private string LinkLoc = ""; //每天记录绑定的科室id
        public string NurseLocHuanYe = "Y"; //按绑定的护士登陆科室换页
        public string Patcurloc = ""; //病人当前科室

        private string rowprintinfo = ""; //记录每条记录的打印信息 生成图片用
        private void clearvalue()
        {
            head = null;
            HPagRow.Clear();
            HPagRow1.Clear();
            tableBak.Clear();
            HLhead.Clear();
            BakRowH.Clear();
            printpagecount = 0;
            HLasthead.Clear();  //沈阳医大诊断换页
            // BottomArray.Clear();
        }
        /// </summary>
        /// <param name="connectstr"></param>
        public void SetConnectStr(string connectstr)
        {
            ConnectStr = connectstr;
            Comm.connectstr = connectstr;
            // MessageBox.Show(connectstr); 
            //Comm.connect(); 
        }
        public void SetParrm(string parm)
        {
            //MessageBox.Show(parm);
            Parrm = parm;
            string[] tem = parm.Split('^');
            for (int i = 0; i < tem.Length; i++)
            {
                if (tem[i] == "") continue;
                GParrm[i] = tem[i];

            }
            string[] tempar = parm.Split('!');
            EpisodeID = tempar[0];
            if (tempar.Length > 6)
            {
                curhead = tempar[6]; //表头变更当前表头id
                Firsthead = tempar[6]; //病区打印时前台传入的当前表头
            }
            if (tempar.Length > 5)
            {
                EmrCode = tempar[5];
            }
         
        }
        public void SetHead(string headstr)
        {  //20110116 qse
            // MessageBox.Show("dd");
            string[] tem = headstr.Split('^');
            Hashtable titl = new Hashtable();
            for (int i = 0; i < tem.Length; i++)
            {
                if (tem[i] == "") continue;
                string[] itm = tem[i].Split('@');
                titl[itm[0]] = itm[1];
            }
            if (xmlprndoc["Root"]["PageHeadFoot"] != null)
            {
                foreach (XmlNode xn in xmlprndoc["Root"]["PageHeadFoot"].ChildNodes)
                {
                    if (xmlprndoc["Root"]["InstanceData"][xn.Name] != null)
                    {
                        XmlNode xx = xmlprndoc["Root"]["InstanceData"][xn.Name];
                        if (titl.Contains(xx.Attributes["text"].Value))
                        {
                            if (tranflag == "1")
                            {
                                if (xx.Attributes["text"].Value == "LOC") PageLocNod = xx; //科室变换 //转科
                                if (xx.Attributes["text"].Value == "BEDCODE") PageBedNod = xx; //床号变换 //转科
                            }
                            if ((((xx.Attributes["text"].Value == "LOC") & (PrnLoc != "")) || ((PrnBed != "") & (xx.Attributes["text"].Value == "BEDCODE"))) && (tranflag == "1")) //转科
                            {
                            }
                            else
                            {
                                if (xx.Attributes["text"].Value == "DIAG") PageDiagNod = xx; //诊断变换
                                xx.Attributes["text"].Value = titl[xx.Attributes["text"].Value].ToString();
                                if (xx.Attributes["text"].Value == "") xx.Attributes["text"].Value = " ";
                            }
                        }

                    }
                }
            }

        }
        public void SetMthParm(string parm)
        {
            Parrm = parm;
        }

        private string GetTreeData(string treename, string cls)
        {
            //ClsNurMg.ClsNurMg NurEmr = new ClsNurMg.ClsNurMg();
            //NurEmr.ConnectString = ConnectStr ;
            //NurEmr.Connect();
            //string xmlstr = NurEmr.getSubData(treename, cls);
            datastream = null;
            datastream = Comm.DocServComm.GetEmrData("NurEmr.NurEmrSub:GetStream", "parr:" + treename + "!", "!");

            return datastream.CommString;
        }
        public void SetPreView(string flag)
        {
            PreView = flag;

            // previewPrint = PreView;
        }
        private void SettingPrinter(XmlNode ps)
        {

            pDoc.DefaultPageSettings.Landscape = bool.Parse(ps.Attributes["LandFlag"].Value);
            //页边距   
            //anhui20120515

            if (SetLeft != 0)
            {
                pDoc.DefaultPageSettings.Margins.Left = SetLeft;
                MargLeft = SetLeft;
                pDoc.DefaultPageSettings.Landscape = SetLandscape;
            }
            else
            {
                pDoc.DefaultPageSettings.Margins.Left = int.Parse(ps.Attributes["left"].Value) * 39 / 10;
                MargLeft = int.Parse(ps.Attributes["left"].Value) * 39 / 10;
                InitLandscape = bool.Parse(ps.Attributes["LandFlag"].Value);


            }
            if (SetRight != 0)
            {
                pDoc.DefaultPageSettings.Margins.Right = SetRight;
            }
            else
            {
                pDoc.DefaultPageSettings.Margins.Right = int.Parse(ps.Attributes["right"].Value) * 39 / 10;
            }
            if (SetTop != 0)
            {
                pDoc.DefaultPageSettings.Margins.Top = SetTop;
                MargTop = SetTop;

            }
            else
            {
                pDoc.DefaultPageSettings.Margins.Top = int.Parse(ps.Attributes["top"].Value) * 39 / 10;
                MargTop = int.Parse(ps.Attributes["top"].Value) * 39 / 10;

            }
            if (SetBottom != 0)
            {
                pDoc.DefaultPageSettings.Margins.Bottom = SetBottom;
            }
            else
            {
                pDoc.DefaultPageSettings.Margins.Bottom = int.Parse(ps.Attributes["bottom"].Value) * 39 / 10;
                InitBottom = int.Parse(ps.Attributes["bottom"].Value) * 39 / 10;

            }
            //MessageBox.Show(this.printDocument1.DefaultPageSettings.PaperSize.ToString()   +   "   "   +   this.printDocument1.PrinterSettings.DefaultPageSettings.PaperSize.ToString()   +   "   "   +   this.printDocument1.PrinterSettings.DefaultPageSettings.Margins.ToString());   
            //纸张类型   
            String paperName = ps.Attributes["PaperType"].Value;
            bool fitPaper = false;
            if (SetPageType != null)
            {
                pDoc.DefaultPageSettings.PaperSize = SetPageType;
                fitPaper = true;
            }
            else
            {
                //获取打印机支持的所有纸张类型   
                foreach (PaperSize size in pDoc.PrinterSettings.PaperSizes)
                {
                    //看该打印机是否有我们需要的纸张类型   
                    if (size.PaperName.IndexOf(paperName) != -1)
                    {
                        pDoc.DefaultPageSettings.PaperSize = size;
                        fitPaper = true;
                        InitHeight = size.Height;
                        break;
                    }

                }
            }
            //如果没有我们需要的标准类型，则使用自定义的尺寸   
            if (!fitPaper)
            {
                foreach (PaperSize size in pDoc.PrinterSettings.PaperSizes)
                {

                    if (size.PaperName.Contains("A4"))
                    {
                        pDoc.DefaultPageSettings.PaperSize = size;
                        fitPaper = true;
                        break;
                    }
                }
            }
            if (!fitPaper)
            {

                pDoc.DefaultPageSettings.PaperSize = new PaperSize("Custom", int.Parse(ps.Attributes["Width"].Value), int.Parse(ps.Attributes["Height"].Value));
                //this.printDocument1.PrinterSettings.DefaultPageSettings.PaperSize   =   this.printDocument1.DefaultPageSettings.PaperSize;   
                //this.printDocument1.PrinterSettings.DefaultPageSettings.Margins   =   this.printDocument1.DefaultPageSettings.Margins;   
            }
        }
        protected virtual PageSettings ShowPageSetupDialog(PrintDocument printDocument)
        {  //检查printDocument是否为空，空的话抛出异常  
            //  ThrowPrintDocumentNullException(printDocument);   //声明返回值的PageSettings  
            PageSettings ps = new PageSettings();   //申明并实例化PageSetupDialog  
            PageSetupDialog psDlg = new PageSetupDialog();
            ps = printDocument.DefaultPageSettings;
            try
            {  //相关文档及文档页面默认设置 
                psDlg.Document = printDocument;
                Margins mg = printDocument.DefaultPageSettings.Margins;
                if (System.Globalization.RegionInfo.CurrentRegion.IsMetric)
                {
                    mg = PrinterUnitConvert.Convert(mg, PrinterUnit.Display, PrinterUnit.TenthsOfAMillimeter);
                }   //备份打印文档的DefaultPageSettings，  //因为转换后会改变，  
                //而设置对话框单击取消按钮后不还原就不能正确显示原来的值 
                PageSettings psPrintDocumentBack = (PageSettings)(printDocument.DefaultPageSettings.Clone());
                psDlg.PageSettings = psPrintDocumentBack;  //printDocument.DefaultPageSettings;   
                //用printDocument的时取消了对话框就要还原  
                psDlg.PageSettings.Margins = mg;
                //显示对话框  
                DialogResult result = psDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    ps = psDlg.PageSettings;
                    oldPageSetting = ps;
                    printDocument.DefaultPageSettings = psDlg.PageSettings;
                    SetLeft = psDlg.PageSettings.Margins.Left;
                    SetRight = psDlg.PageSettings.Margins.Right;
                    SetTop = psDlg.PageSettings.Margins.Top;
                    SetBottom = psDlg.PageSettings.Margins.Bottom;
                    SetPageType = psDlg.PageSettings.PaperSize;
                    SetHeight = psDlg.PageSettings.PaperSize.Height;
                    SetLandscape = psDlg.PageSettings.Landscape;
                    //pDoc.DefaultPageSettings.Landscape
                    PageSet = true;
                }
                else { }
            }
            catch (System.Drawing.Printing.InvalidPrinterException e)
            {
                // ShowInvalidPrinterException(e); 
            }
            catch (Exception ex)
            {
                // ShowPrinterException(ex); 
            }
            finally
            {
                psDlg.Dispose();
                PrnPreView.Dispose();
                clearvalue();
                clearxuprint();//续打
                PreView = "1";
                PrintOut();
                psDlg = null;

            }
            return ps;
        }


        protected void PageSet_Click(object sender, EventArgs e)
        {
            ShowPageSetupDialog(pDoc);

            return;
            pageSetupDialog = new PageSetupDialog();
            pageSetupDialog.Document = pDoc;

            pageSetupDialog.PrinterSettings = pDoc.PrinterSettings;

            if (pageSetupDialog.ShowDialog() == DialogResult.OK)
            {
                // PrintPriview.Document = null;
                //  PrintPriview.;
                // PrintPriview.Document = PrintDocument1;
                //  PrintPriview.Document.Print();
                PageSet = true;
                MargLeft = pDoc.DefaultPageSettings.Margins.Left;
                MargTop = pDoc.DefaultPageSettings.Margins.Top;
                MargRight = pDoc.DefaultPageSettings.Margins.Right; //会诊打印更新
                MargBottom = pDoc.DefaultPageSettings.Margins.Bottom; //会诊打印更新

                MargLeft = pageSetupDialog.PageSettings.Margins.Left;
                MargTop = pageSetupDialog.PageSettings.Margins.Top;
                MargRight = pageSetupDialog.PageSettings.Margins.Right;
                MargBottom = pageSetupDialog.PageSettings.Margins.Bottom;
                oldPageSetting = pDoc.DefaultPageSettings;
                // oldPageSetting.Margins = pageSetupDialog.PageSettings.Margins;

            }
        }
        //shezhibiaotou 
        private void setlablehead(string LabHead)
        {
            string[] lllab = LabHead.Split('&');
            for (int jj = 0; jj < lllab.Length; jj++)
            {
                if (lllab[jj] == "") continue;
                string[] itt = lllab[jj].Split('|');
                if (DataLable.Contains(itt[0]))
                {
                    DataLable[itt[0]] = itt[1];
                }
                else
                {
                    DataLable.Add(itt[0], itt[1]);
                }
            }
        }
        public void PrintOut()
        {
            //System.Diagnostics.Debugger.Launch();
            if (StartMakePic == "Y") return; //开始生成图片
            PrnFlag = 0;
            Comm.DocServComm.Url = WebUrl;

            PgdPrintedArray.Clear();
            //生成图片
            if (MakeTemp == "Y")
            {
                try
                {
                    XmlDocument pathdoc = new XmlDocument();
                    pathdoc.Load(filepath);
                    XmlNode ftpnod = pathdoc["FTP"];
                    ftppath = ftpnod.Attributes["server"].Value;
                    ftpuer = ftpnod.Attributes["user"].Value;
                    ftppwd = ftpnod.Attributes["pwd"].Value;
                    ftpport = ftpnod.Attributes["port"].Value;
                    ftpdealyTim = ftpnod.Attributes["dealyTim"].Value;
                    if (EpisodeID == "")
                    {
                        MessageBox.Show("EpisodeID不能为空！");
                        return;
                    }
                    if (EmrCode == "")
                    {
                        MessageBox.Show("EmrCode不能为空！");
                        return;
                    }
                    if (curPages == 0)
                    {
                        MakeAllPages = "Y";
                    }


                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }

            }
            if (xuflag == "1") //续打
            {
                string parr = EpisodeID + "&" + ItmName;
                lastprninfo = Comm.DocServComm.GetData("Nur.DHCNurRecPrintStPos:getval", "par:" + parr + "^");
                if ((lastprninfo == "") || (lastprninfo == null))
                {
                    if (MessageBox.Show("没有续打记录,设置起始行？\n确定：设置起始行\n取消：全部打印", "再确认一下！", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    {
                        return;
                    }
                    else
                    {
                        //xuflag = "0";

                    }
                }
                else
                {

                    string[] arrr = lastprninfo.Split('^');
                    Parrm = Parrm + "!" + arrr[0];
                    Startrow = Int32.Parse(arrr[2]);
                    Startpage = Int32.Parse(arrr[1]);
                    string prnflag = arrr[3];
                    curPages = Startpage;
                    string lastrowinfo = "";
                    try
                    {
                        string inpar = arrr[0].Replace("||", "&");
                        try
                        {
                            string addinfo = Comm.DocServComm.GetData("Nur.DHCNurRecPrint:ifhavealert", "par:" + inpar + "^");
                            if ((addinfo != "") && (addinfo != null))
                            {
                                if (MessageBox.Show("续打位置之前有记录增加或修改,详细信息：\n" + addinfo + "\n是否重新设置起始位置？\n确定：取消续打，重新设置\n取消 : 继续打印", "再确认一下！", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                                {
                                    return;
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Nur.DHCNurRecPrint:ifhavealert是否存在？" + inpar + ex.Message);
                        }
                        lastrowinfo = Comm.DocServComm.GetData("Nur.DHCNurRecPrint:Getinfobyid", "par:" + inpar + "^");
                        int stp = (Startpage + 1);
                        if (prnflag == "1")
                        {
                            if (MessageBox.Show("上次打印到第" + stp + "页,第" + Startrow + "行(" + arrr[5] + ")？\n记录明细:" + lastrowinfo + "\n" + arrr[0] + "\n是否放入正确的纸张？\n", "再确认一下！", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                            {

                            }
                            else
                            {
                                clearxuprint();
                                return;
                            }
                        }
                        else
                        {
                            string tex = "是否放入正确的纸张?";
                            if (Startrow == 0) tex = "是否放入空白纸张？";
                            if (MessageBox.Show("从第" + stp + "页,第" + (Startrow + 1) + "行开始打印(" + arrr[5] + ")？\n记录明细:" + lastrowinfo + "\n" + arrr[0] + "\n" + tex + "\n", "再确认一下！", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                            {

                            }
                            else
                            {
                                clearxuprint();
                                return;
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Nur.DHCNurRecPrint:Getinfobyid是否存在？params:" + arrr[0] + ex.Message);

                    }

                }


            }
            showPrintDialog = false;
            dzqmprnpgd.Clear(); //20140826
            nextpagehastable.Clear(); //NextPageFlag换页标志
            if (LabHead != "") //表头变更打印标题
            {  ///空白栏设定
                setlablehead(LabHead);
               
            }

            if ((PrnLoc != "") & (CAStart == "1"))
            {
                IsVerifyCALoc = Comm.DocServComm.GetData("web.DHCNurCASignVerify:GetIsVerifyCA", "par:" + PrnLoc + "^");
            }
            if ((TitleStr == "") && (PrnLoc != ""))  //续打201407
            {
                try
                {
                    TitleStr = Comm.DocServComm.GetData("web.DHCMGNurComm:PatInfo", "par:" + EpisodeID + "^");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("获取病人信息出错！");
                }
            }
            HFCaption["LHead"] = LHeadCaption;
            HFCaption["RHead"] = RHeadCaption;
            HFCaption["LFoot"] = LFootCaption;
            HFCaption["RFoot"] = RFootCaption;

            ItmCount = 0;
            PagOffY = 0;
            PrnCount = 0;
            if (curPages < 0) curPages = 0;
            Pages = curPages;
            pDoc = new PrintDocument();
            PrnPreView = new PrintPreviewDialog();

            ToolStrip MyToolStrip = (ToolStrip)PrnPreView.Controls["toolStrip1"];
            System.Drawing.Image tmpImage;  //可以设置一张图片的
            MyToolStrip.Items.Add("页面设置", null); //'增加一个页面设置按钮
            MyToolStrip.Items[MyToolStrip.Items.Count - 1].Click += new EventHandler(PageSet_Click);

            HPagRow.Clear();
            HPagRow1.Clear();

            HLhead.Clear();  //记录标题
            PageDiagH.Clear();
            PageLocH.Clear();
            PageBedH.Clear();
            BakRowH.Clear(); //记录行高
            PageLine.Clear();
            //  MessageBox.Show("sss"); 
            //总页数
            // printpagecount = obserDate.Length;
            ///数据
            prnData.Clear(); DataTyp.Clear();
            string xmlstr = "";
            if (MthArr != "")
            {
                xmlstr = getprintdata();
            }
            if (xmlstr != "")
            {
                XmlDocument xmdata = new XmlDocument();
                xmdata.LoadXml(xmlstr);
                foreach (XmlNode dx in xmdata["Root"].ChildNodes)
                {
                    prnData.Add(dx.Name, dx.InnerText);
                    DataTyp.Add(dx.Name, dx.Attributes["typ"].Value);
                }
            }
            ///
            xmlstr = GetTreeData(ItmName, "NurEmr.NurEmrSub");
            if (xmlstr != "")
            {
                xmlprndoc.LoadXml(xmlstr);
                if (xmlprndoc["Root"].Attributes["DataType"] != null)
                {
                    DataType = xmlprndoc["Root"].Attributes["DataType"].Value;
                }
                if (xmlprndoc["Root"].Attributes["ShapePrn"] != null)
                {
                    ShapePrn = xmlprndoc["Root"].Attributes["ShapePrn"].Value;
                }
                if (PrnLoc != "")
                {
                    XmlDocument xmltemp = new XmlDocument();
                    //ClsNurMg.ClsNurMg NurEmr = new ClsNurMg.ClsNurMg();
                    //NurEmr.ConnectString = ConnectStr ;
                    //NurEmr.Connect();
                    //string xmlstr1 = NurEmr.getPrnChHead (ItmName , PrnLoc );
                    datastream = null;
                    string xmlstr1 = "";
                    datastream = Comm.DocServComm.GetEmrData("Nur.NurDHCNurChangePrnHead:GetStream", "loc:" + PrnLoc + "!code:" + ItmName + "!", "!");
                    xmlstr1 = datastream.CommString;

                    if ((xmlstr1 != "") && (xmlstr1 != null))
                    {
                        xmltemp.LoadXml(xmlstr1);
                        if (xmlprndoc["Root"]["THEAD"] != null)
                        {
                            foreach (XmlNode xr in xmltemp["Root"].ChildNodes)
                            {
                                if (xmlprndoc["Root"]["THEAD"][xr.Name] != null)
                                {
                                    XmlNode xim = xmlprndoc.ImportNode(xr, true);
                                    xmlprndoc["Root"]["THEAD"].RemoveChild(xmlprndoc["Root"]["THEAD"][xr.Name]);
                                    xmlprndoc["Root"]["THEAD"].AppendChild(xim);
                                }
                            }
                        }
                    }

                }
                //10-10-20  15:45  dhcc  //替换成对应的数据
                foreach (XmlNode xx in xmlprndoc["Root"]["InstanceData"].ChildNodes)
                {
                    if ((prnData.Contains(xx.Name)) && (xx.Name.IndexOf("Label") != -1))
                    {
                        xx.Attributes["text"].Value = prnData[xx.Name].ToString(); ;
                        //替换对应的text
                    }

                    if (prnData.Contains(xx.Name))
                    {
                        string prnval = prnData[xx.Name].ToString();
                        if (CAStart == "0") continue;
                        if (prnval.IndexOf('+') > -1)
                        {
                            string[] array = prnval.Split('+');
                            for (int i = 0; i < array.Length; i++)
                            {
                                string userstr = array[i].ToString();
                                if (userstr.IndexOf('*') > -1)
                                {
                                    string[] array2 = userstr.Split('*');
                                    string userstr2 = array2[1].ToString();
                                    if (userstr2 != " ")
                                    {
                                        try
                                        {
                                            string uid = userstr2;
                                            string imageuser = Comm.DocServComm.GetData("web.DHCNurSignVerify:GetUserSignImage", "par:" + uid + "^");
                                            if (imageuser != null)
                                            {
                                                // xx.Attributes["text"].Value = imageuser;
                                                if (dzqmprnpgd.Contains(xx.Name))
                                                {
                                                    string imagestr = dzqmprnpgd[xx.Name].ToString();
                                                    string newimagetra = imageuser + "$" + imageuser;
                                                    dzqmprnpgd.Remove(xx.Name);
                                                    dzqmprnpgd.Add(xx.Name, newimagetra);
                                                }
                                                else
                                                {
                                                    dzqmprnpgd.Add(xx.Name, imageuser);
                                                }

                                            }
                                        }
                                        catch(Exception ex) { 
                                        
                                        }
                                    }
                                }
                            }

                        }
                        else
                        {
                            if (prnval.IndexOf('*') > -1)
                            {
                                if (prnval.IndexOf('|') > -1)
                                {
                                    string[] array2 = prnval.Split('|');
                                    prnval = array2[1].ToString();
                                }
                                string[] array = prnval.Split('*');
                                string userstr = array[1].ToString();

                                if (userstr != " ")
                                {
                                    try
                                    {
                                        string uid = userstr;
                                        string imageuser = Comm.DocServComm.GetData("web.DHCNurSignVerify:GetUserSignImage", "par:" + uid + "^");
                                        if (imageuser != null)
                                        {
                                            // xx.Attributes["text"].Value = imageuser;
                                            if (dzqmprnpgd.Contains(xx.Name))
                                            { }
                                            else
                                            {
                                                dzqmprnpgd.Add(xx.Name, imageuser);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    { }
                                }
                            }
                        }
                    }
                }
                SetHead(TitleStr);  //设置head
                if (xmlprndoc["Root"]["QtData"] != null)
                {
                    if (xmlprndoc["Root"]["QtData"].Attributes["typ"] != null) SourceFlag = xmlprndoc["Root"]["QtData"].Attributes["typ"].Value;
                    if (xmlprndoc["Root"]["QtData"].Attributes["Cls"] != null) DataCls = xmlprndoc["Root"]["QtData"].Attributes["Cls"].Value;
                    if (xmlprndoc["Root"]["QtData"].Attributes["mth"] != null) DataMth = xmlprndoc["Root"]["QtData"].Attributes["mth"].Value;
                    if (Parrm == "") if (xmlprndoc["Root"]["QtData"].Attributes["parrm"] != null) Parrm = xmlprndoc["Root"]["QtData"].Attributes["parrm"].Value;
                }
                if ((DataType == "qt") && (DesignFlag == "1"))
                {
                    frmParrm frm = new frmParrm();
                    Comm.formset(frm);
                    if (SourceFlag == "Record")
                    {
                        // string ret = Comm.GetCacheData("web.DHCMGPrintComm", "GetQueryHead", new object[] { DataCls + ":" + DataMth });
                        string ret = Comm.DocServComm.GetData("web.DHCMGPrintComm:GetQueryHead", "");
                        string[] tm = ret.Split((char)1);
                        frm.parrm = tm[1].ToString().Split('&');
                        frm.txtParrm.Visible = false;
                    }
                    frm.frmrp = this;
                    if (frm.ShowDialog() == DialogResult.OK)
                    {

                    }
                    else
                    {
                        return;
                    }

                }
                if (xmlprndoc["Root"]["SHAPES"] != null)
                {
                    foreach (XmlNode xx in xmlprndoc["Root"]["SHAPES"].ChildNodes)
                    {
                        if (xx.Attributes["typ"].Value == "NurEmrMaintain.RectangleShape")
                        {
                            string p1 = xx.Attributes["P1"].Value;
                            string p2 = xx.Attributes["P2"].Value;
                            string[] pt1 = p1.Split(',');
                            string[] pt2 = p2.Split(',');
                            tabx = int.Parse(pt1[0]);
                            taby = int.Parse(pt1[1]);
                            tabW = int.Parse(pt2[0]);
                            tabH = int.Parse(pt2[1]);
                        }
                    }

                }

            }
            XmlNode xHn = null;
            ////qse 20110216 add
            foreach (XmlNode xs in xmlprndoc["Root"]["InstanceData"].ChildNodes)
            {
                if (xs.Attributes["text"] != null)
                {
                    string a = xs.Attributes["text"].Value;
                    if (xs.Name.Length > 4)
                    {
                        if (xs.Name.Substring(0, 4) == "Page") continue;
                    }
                    if (a == "") continue;
                    if (a.Substring(0, 1) == "T") xHn = xs;
                }
            }
            if (xHn != null)
            {

                tabx = int.Parse(xHn.Attributes["left"].Value);
                taby = int.Parse(xHn.Attributes["top"].Value);

            }
            int HW = 0;

            SettingPrinter(xmlprndoc["Root"]["PrintCon"]);
            if (xmlprndoc["Root"]["PageHeadFoot"] != null)
            {
            }
            // xmlprndoc.Save("D:\\2.xml");
            if ((DataType == "xml") || (DataType == ""))
            {
                PrnDiaglog = new PrintDialog();
                PrnDiaglog.AllowSomePages = true;
                PrnDiaglog.AllowSelection = true;
                PrnDiaglog.AllowCurrentPage = true;
                PrnDiaglog.Document = pDoc;
                pDoc.PrintPage += new PrintPageEventHandler(pd_PrintPage);
                pDoc.BeginPrint += new PrintEventHandler(pDoc_BeginPrint);
                pDoc.EndPrint += new PrintEventHandler(pDoc_EndPrint);
            }
            else
            {
                if (SourceFlag == "Method")
                { //绑定方法
                    string ret = Comm.DocServComm.GetData("web.DHCMGPrintComm:GetData", Parrm);

                    if (ret != "")
                    {
                        string[] tm = ret.Split('^');
                        for (int i = 0; i < tm.Length; i++)
                        {
                            string[] itm = tm[i].Split((char)1);
                            DataHash.Add(itm[0], itm[1]);
                        }
                        pDoc.PrintPage += new PrintPageEventHandler(pd_PrintPage2);
                    }
                }
                if (SourceFlag == "Record")
                { //绑定记录集
                    //  Comm.ExeQuery(DataCls, DataMth, Parr );

                    try
                    {
                        if (MakeTemp == "Y")
                        {
                            Parrm = Parrm + "!!!!$" + EmrCode + "*" + MakeAllPages;
                        }
                        CommData = Comm.DocServComm.GetQueryDataX(DataCls + "." + DataMth, "parr:" + Parrm, "^");
                    }
                    catch (Exception exs)
                    {
                        MessageBox.Show(DataCls + "." + DataMth + "(parr:" + Parrm + ")" + exs.Message);
                        return;
                    }
                    Comm tab = new Comm();
                    int uu = 0;
                    ArrayList arr = new ArrayList();
                    if (xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"] != null)
                    { /////////////使用tabx
                        string[] hh = xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"].Value.Split('|');
                        for (int i = 0; i < hh.Length; i++)
                        {   //5-x,6-y,7-xname,8--key-relname 1-w,0-h
                            //hh + "^" + grid1.Columns[c].Width.ToString() + "^" + Num.ToString() + "^" + i.ToString() + "^" + c.ToString() + "^" + posx.ToString() + "^" + posy.ToString() + "^" + TitlArr[grid1[i, c].Value.ToString()].ToString();
                            if (hh[i] == "") continue;
                            string[] l = hh[i].Split('^');
                            if (l[8] == "") continue;
                            int lef = int.Parse(l[5]) + tabx;
                            string str = l[8] + "^" + lef.ToString() + "^" + l[1] + "^" + l[0];
                            if (int.Parse(l[1]) > 5) HW = HW + int.Parse(l[1]);
                            arr.Add(str);
                            uu++;

                        }
                    }
                    else
                    {
                        foreach (XmlNode xnod in xmlprndoc["Root"]["QtData"].ChildNodes)
                        {

                            if (xmlprndoc["Root"]["InstanceData"]["_" + xnod.Name] != null)
                            {
                                string str = xnod.Name + "^" + xmlprndoc["Root"]["InstanceData"]["_" + xnod.Name].Attributes["left"].Value + "^" + xmlprndoc["Root"]["InstanceData"]["_" + xnod.Name].Attributes["width"].Value + "^" + xmlprndoc["Root"]["InstanceData"]["_" + xnod.Name].Attributes["height"].Value;
                                arr.Add(str);
                                uu++;
                            }
                        }
                    }
                    //head = new string[uu];
                    int flag1 = 0; //转科是否有loc列
                    int flag2 = 0; //是否有bed列
                    int flag3 = 0; //换页标志
                    int flag4 = 0; //是否有par列  //续打
                    int flag5 = 0; //是否有rw列  //续打
                    int flag6 = 0; //是否RecSaveLoc,保存人科室
                    int flag7 = 0; //表头变换换页打印 --20141201
                    numprnflag = 7;
                    if (PrnLoc != "")
                    {
                        for (int j = 0; j < arr.Count; j++)
                        {
                            string[] tem = arr[j].ToString().Split('^');
                            if (tem[0] == "RecLoc")
                            {
                                flag1 = 1;
                                numprnflag = numprnflag - 1;
                            }
                            if (tem[0] == "RecBed")
                            {
                                flag2 = 1;
                                numprnflag = numprnflag - 1;
                            }

                            if (tem[0] == "NextPageFlag")
                            {
                                flag3 = 1;
                                numprnflag = numprnflag - 1;
                            }
                            if (tem[0] == "par")  //续打
                            {
                                flag4 = 1;
                                numprnflag = numprnflag - 1;
                            }
                            if (tem[0] == "rw") //续打
                            {
                                flag5 = 1;
                                numprnflag = numprnflag - 1;
                            }
                            if (tem[0] == "RecNurseLoc")
                            {
                                flag6 = 1;
                                numprnflag = numprnflag - 1;
                            }
                            if (tem[0] == "HeadDR")
                            {
                                flag7 = 1;
                                numprnflag = numprnflag - 1;
                            }
                        }
                        if ((flag1 == 0) & (flag2 == 0) & (flag3 == 0) & (flag4 == 0) & (flag5 == 0) & (flag6 == 0) & (flag7 == 0)) //续打
                        {
                            head = new string[uu + 7];
                        }
                        else
                        {
                            if ((flag1 == 1) & (flag2 == 1) & (flag3 == 1) & (flag4 == 1) & (flag5 == 1) & (flag6 == 1) & (flag7 == 0)) //续打
                            {
                                head = new string[uu];
                            }
                            else
                            {

                                head = new string[uu + numprnflag];
                            }

                        }
                    }
                    else
                    {
                        head = new string[uu];
                    }
                    if (HW != 0) tabW = HW + tabx;
                    do
                    {
                        int min = -1;
                        string minstr = "";
                        int index = 0;
                        ///////重新排序按left -x
                        for (int j = 0; j < arr.Count; j++)
                        {
                            string[] tem = arr[j].ToString().Split('^');
                            int l = int.Parse(tem[1]);
                            if (min == -1)
                            {
                                min = l;
                                minstr = arr[j].ToString();
                                index = j;
                            }
                            if (min > l)
                            {
                                min = l;
                                minstr = arr[j].ToString();
                                index = j;
                            }

                        }
                        arr.RemoveAt(index);
                        head[ItmCount] = minstr;
                        ItmCount++;
                    } while (arr.Count > 0);
                    if (head.Length > 1)
                    {
                        if (PrnLoc != "")
                        {
                            for (int k = 0; k < numprnflag; k++)
                            {

                                if (flag1 == 0)
                                {
                                    if (head[uu + k] != null) continue;
                                    string[] tmphead = head[uu + k - 1].Split('^');  //RecLoc 记录科室id
                                    int locx = int.Parse(tmphead[1]) + int.Parse(tmphead[2]);
                                    head[uu + k] = "RecLoc^" + locx + "^1^" + tmphead[3];
                                    flag1 = 1;
                                    continue;
                                }
                                if (flag2 == 0)
                                {
                                    if (head[uu + k] != null) continue;
                                    string[] tmphead2 = head[uu + k - 1].Split('^');  //RecBed 记录床号
                                    int locx2 = int.Parse(tmphead2[1]) + int.Parse(tmphead2[2]);
                                    head[uu + k] = "RecBed^" + locx2 + "^1^" + tmphead2[3];
                                    flag2 = 1;
                                    continue;
                                }
                                if (flag3 == 0)
                                {
                                    if (head[uu + k] != null) continue;
                                    string[] tmphead3 = head[uu + k - 1].Split('^');  //NextPageFlag 换页标志
                                    int locx3 = int.Parse(tmphead3[1]) + int.Parse(tmphead3[2]);
                                    head[uu + k] = "NextPageFlag^" + locx3 + "^1^" + tmphead3[3];
                                    flag3 = 1;
                                    continue;
                                }
                                if (flag4 == 0) //加上par列
                                {
                                    if (head[uu + k] != null) continue;
                                    string[] tmphead4 = head[uu + k - 1].Split('^');  //续打
                                    int locx4 = int.Parse(tmphead4[1]) + int.Parse(tmphead4[2]);
                                    head[uu + k] = "par^" + locx4 + "^1^" + tmphead4[3];
                                    flag4 = 1;
                                    continue;
                                }
                                if (flag5 == 0) //加上rw列
                                {
                                    if (head[uu + k] != null) continue;
                                    string[] tmphead5 = head[uu + k - 1].Split('^');  //续打
                                    int locx5 = int.Parse(tmphead5[1]) + int.Parse(tmphead5[2]);
                                    head[uu + k] = "rw^" + locx5 + "^1^" + tmphead5[3];
                                    flag5 = 1;

                                }
                                if (flag6 == 0) //RecSaveLoc
                                {
                                    if (head[uu + k] != null) continue;
                                    string[] tmphead6 = head[uu + k - 1].Split('^');  //续打
                                    int locx6 = int.Parse(tmphead6[1]) + int.Parse(tmphead6[2]);
                                    head[uu + k] = "RecNurseLoc^" + locx6 + "^1^" + tmphead6[3];
                                    flag6 = 1;

                                }
                                if (flag7 == 0) //RecSaveLoc
                                {
                                    if (head[uu + k] != null) continue;
                                    string[] tmphead7 = head[uu + k - 1].Split('^');  //
                                    int locx7 = int.Parse(tmphead7[1]) + int.Parse(tmphead7[2]);
                                    head[uu + k] = "HeadDR^" + locx7 + "^1^" + tmphead7[3];
                                    flag6 = 1;

                                }
                            }
                        }
                        if (CommData.Tables[0].Columns.Count > 1)
                        {
                            try
                            {
                                table = tab.GetTable3(CommData.Tables[0], head, "dt");
                            }
                            catch (Exception ex)
                            {
                                if (PrnLoc == "")
                                {
                                    MessageBox.Show("PrnLoc参数不能为空1");

                                }
                                else
                                {
                                    MessageBox.Show(ex.Message);
                                }
                                return;
                            }
                        }
                        else
                        {
                            rowidha.Clear(); //续打
                            table = tab.GetTable4(CommData.Tables[0], head, "dt", ref rowidha);
                        }

                        tableBak = tab.SetTabHead(head, "dbak");
                        //MessageBox.Show(tableBak.Rows.Count.ToString()); 
                        tableBak.Clear();
                        // BakTable = table;
                        if (table.Rows.Count == 0)
                        {
                            if (xuflag == "1")
                            {
                                clearxuprint();//续打
                                MessageBox.Show("没有可续打的数据!");
                            }
                            else
                            {
                                if (MakeTemp == "Y")
                                { }
                                else
                                {
                                    MessageBox.Show("没有可打印的数据!");
                                    CareDateTim = "";
                                }
                            }
                            return;
                        }
                        else
                        {
                            ///先设定诊断
                            string PageDiag = "";
                          
                            if (MakeTemp == "Y")
                            {
                                if (table.Columns.Contains("HeadDR"))       //表头变化模板
                                {
                                    curhead = table.Rows[Row]["HeadDR"].ToString();
                                }
                           }
                                                       
                            if (table.Columns.Contains("DiagnosDr"))
                            {
                                PageDiag = Comm.DocServComm.GetData("Nur.DHCNurCopyDiagnos:GetNurDiagnos", "par:" + table.Rows[0]["DiagnosDr"].ToString() + "^");

                                if (LHeadCaption != "")
                                {
                                    int sindex = LHeadCaption.IndexOf("诊断:");
                                    if (sindex > -1)
                                    {
                                        string oldcaption = LHeadCaption;
                                        LHeadCaption = LHeadCaption.Substring(0, sindex) + "诊断:" + PageDiag;
                                        HFCaption["LHead"] = LHeadCaption;
                                    }
                                }
                                if (PageDiagNod != null)
                                {
                                    PageDiagNod.Attributes["text"].Value = PageDiag;
                                }
                            }
                            if (PageLocNod != null) //转科
                            {

                                string recloc = table.Rows[Row]["RecLoc"].ToString();
                                if (NurseLocHuanYe == "Y") //按护士科室换页
                                {
                                    LinkLoc = table.Rows[Row]["RecNurseLoc"].ToString();

                                    if (LinkLoc == "")
                                    {
                                        string parr = EpisodeID + "!" + EmrCode;
                                        try
                                        {
                                            LinkLoc = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetFirstNulRecNurseloc", "par:" + parr + "^");
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("web.DHCNurRecPrint:GetFirstNulRecNurseloc是否存在？params:" + parr + ex.Message);
                                            return;
                                        }
                                    }
                                    if ((LinkLoc == "") || (LinkLoc == null)) LinkLoc = PrnLoc;
                                }
                                else
                                {
                                    LinkLoc = table.Rows[Row]["RecLoc"].ToString();

                                    if (LinkLoc == "")
                                    {
                                        string parr = EpisodeID + "!" + EmrCode;
                                        try
                                        {
                                            LinkLoc = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetFirstNulRecloc", "par:" + parr + "^");
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("web.DHCNurRecPrint:GetFirstNulRecloc是否存在？params:" + parr + ex.Message);
                                            return;
                                        }
                                    }
                                    if ((LinkLoc == "") || (LinkLoc == null)) LinkLoc = Patcurloc;
                                }
                                string par = table.Rows[Row]["par"].ToString();
                                string rw = table.Rows[Row]["rw"].ToString();
                                string PageLoc = "";
                                try
                                {
                                    PageLoc = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetNurloc", "par:" + LinkLoc + "^");
                                }
                                catch
                                { }
                                PageLocNod.Attributes["text"].Value = PageLoc;
                                // xmlprndoc["Root"]["InstanceData"][PageLocNod.Name].Attributes["text"].Value = PageLoc;
                            }
                            if (PageBedNod != null)
                            {
                                string recbed = table.Rows[0]["RecBed"].ToString();
                                if (recbed == "") recbed = PrnBed; //如果该页第一条记录没关联床号，默认取当前床号
                                PageBedNod.Attributes["text"].Value = recbed;
                            }

                        }

                    }
                    //else
                    // table = tab.GetGrid(Comm.Result, ref head, xmlprndoc);
                    PrnDiaglog = new PrintDialog();
                    PrnDiaglog.AllowSomePages = true;
                    PrnDiaglog.AllowSelection = true;
                    PrnDiaglog.AllowCurrentPage = true;
                    PrnDiaglog.Document = pDoc;
                    //BakTable = table;
                    pDoc.PrintPage += new PrintPageEventHandler(pd_PrintPage1);
                    pDoc.BeginPrint += new PrintEventHandler(pDoc_BeginPrint);
                    // pDoc.EndPrint +=new PrintEventHandler(pDoc_EndPrint); 
                    pDoc.EndPrint += new PrintEventHandler(pDoc_EndPrint);

                }
            }

            ItmCount = 0;
            PagOffY = 0;
            PrnCount = 0;
            // MessageBox.Show(PreView ); 
            // if (Pages==0)Pages =stPage ;
            if ((PreView == "0") && (MakeTemp!="Y"))  //shengchengtupian  make picture
            {

                pDoc.Print();
                pDoc.Dispose();
                return;
            }
            PrnPreView.Document = pDoc;

            //PrnPreView.ClientSize = new Size(this.Width, this.Height);
            PrnPreView.SetDesktopLocation(0, 0);
            PrnPreView.Left = 0;
            PrnPreView.Focus();
            try
            {
                if (MakeTemp == "Y")
                {

                }
                else
                {
                    PrnPreView.ShowDialog();
                    pDoc.Dispose();
                    PrnPreView.Dispose();
                    clearvalue();//续打
                    clearxuprint();//续打
                }
            }
            catch (Exception ex)
            {
                //this.textBox1.Text = ex.ToString();
                string mss = ex.Message.ToString();
                if (mss.IndexOf("成功") > -1)
                { }
                else
                {
                    // MessageBox.Show(mss+"请确定设定的打印机是否存在！");
                }
                //PrnPreView.Close();
            }

            // PrnPreView.ShowDialog();




        }
        private string getprintdata()
        {

            //ClsNurMg.ClsNurMg qu = new ClsNurMg.ClsNurMg();
            //qu.ConnectString = ConnectStr;  //s param="DHCNUR3MouldPrn^169||2^^"
            //qu.Connect();
            //string xmlstr = qu.GetPrintData("DHCNUR3MouldPrn^59||1^^");
            //= qu.GetPrintData(ItmName + "^" + ID + "^" + MultID + "^" + MthArr);
            datastream = null;
            datastream = Comm.DocServComm.GetEmrData("web.DHCNUREMR:GetPrintData", "parr:" + ItmName + "^" + ID + "^" + MultID + "^" + MthArr + "!", "!");
            return datastream.CommString;
        }
        private void DrawString1(ArrayList arrtxt, int x, int y, int width, int height, Graphics g, int RowH, XmlNode td, bool flag, XmlNode xalign)
        {

            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            float fontsize = float.Parse(td.Attributes["fontsize"].InnerText);
            Font font = new Font(td.Attributes["fontname"].InnerText,
            fontsize, style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();

            //设置对齐方式
            //switch (td.Attributes["align"].InnerText)
            //{
            //    case "center":
            //sf.Alignment = StringAlignment.Center;
            //        break;
            //    case "right":
            //sf.Alignment = StringAlignment.Center   ;
            //        break;
            //    default:
            //        sf.Alignment = StringAlignment.Far;
            //        break;
            //}
            int varr = 1;
            if (flag == true) varr = 2; //如果数据是一行 高度要调整为两行高
            sf.LineAlignment = StringAlignment.Center;
            if (xalign != null)
            {
                string align = xalign.Attributes["VAlign"].Value;
                switch (align)
                {
                    case "Center":
                        sf.LineAlignment = StringAlignment.Center;
                        break;
                    case "Right":
                        sf.LineAlignment = StringAlignment.Far;
                        break;
                    case "Left":
                        sf.LineAlignment = StringAlignment.Near;
                        break;
                    default:
                        break;

                }
                string halign = xalign.Attributes["HAlign"].Value;
                switch (halign)
                {
                    case "Center":
                        sf.Alignment = StringAlignment.Center;
                        break;
                    case "Right":
                        sf.Alignment = StringAlignment.Far;
                        break;
                    case "Left":
                        sf.Alignment = StringAlignment.Near;
                        break;
                    default:
                        break;

                }
                if ((flag == false) && (xalign.Name == "User"))
                {
                    int sum = height / (RowH);
                    int aa = sum - arrtxt.Count;
                    y = y + aa * RowH;
                }

            }
            // if (arrtxt == null) return;
            for (int i = 0; i < arrtxt.Count; i++)
            {

                int hhh = RowH * varr;
                Rectangle rect = new Rectangle(x, y,
                    width, hhh);
                if (arrtxt[i].ToString() == "") continue;
                g.DrawString(arrtxt[i].ToString(), font, brush, rect, sf);
                y = y + RowH * varr;
            }
            //  g.DrawRectangle(new Pen(Color.Black, 1), rect);
        }
        private void DrawString11(string arrtxt, int x, int y, int width, int height, Graphics g, int RowH, XmlNode td, bool flag, XmlNode xalign)
        {

            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            float fontsize = float.Parse(td.Attributes["fontsize"].InnerText);
            Font font = new Font(td.Attributes["fontname"].InnerText,
            fontsize, style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();

            //设置对齐方式
            //switch (td.Attributes["align"].InnerText)
            //{
            //    case "center":
            //sf.Alignment = StringAlignment.Center;
            //        break;
            //    case "right":
            //sf.Alignment = StringAlignment.Center   ;
            //        break;
            //    default:
            //        sf.Alignment = StringAlignment.Far;
            //        break;
            //}
            sf.LineAlignment = StringAlignment.Center;
            if (xalign != null)
            {
                string align = xalign.Attributes["VAlign"].Value;
                switch (align)
                {
                    case "Center":
                        sf.LineAlignment = StringAlignment.Center;
                        break;
                    case "Right":
                        sf.LineAlignment = StringAlignment.Far;
                        break;
                    case "Left":
                        sf.LineAlignment = StringAlignment.Near;
                        break;
                    default:
                        break;

                }
                string halign = xalign.Attributes["HAlign"].Value;
                switch (halign)
                {
                    case "Center":
                        sf.Alignment = StringAlignment.Center;
                        break;
                    case "Right":
                        sf.Alignment = StringAlignment.Far;
                        break;
                    case "Left":
                        sf.Alignment = StringAlignment.Near;
                        break;
                    default:
                        break;

                }
            }


            int varr = 1;
            if (flag == true) varr = 2;
            Rectangle rect = new Rectangle(x, y,
                width, RowH * varr);
            g.DrawString(arrtxt, font, brush, rect, sf);
            //  g.DrawRectangle(new Pen(Color.Black, 1), rect);
        }
        private void DrawString(string txt, int x, int y, int width, int height, Graphics g)
        {

            Brush brush = new SolidBrush(Color.FromName("Black"));
            // g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式

            Font font = new Font("宋体",
                10, style, GraphicsUnit.Point);
            StringFormat sf = new StringFormat();

            sf.LineAlignment = StringAlignment.Center;

            Rectangle rect = new Rectangle(x, y,
                width, height);
            g.DrawString(txt, font, brush, rect, sf);

            //  g.DrawRectangle(new Pen(Color.Black, 1), rect);
        }
        private void DrawLines(int x, int y, int w, int height, Graphics g, int RowH, string[] head)
        {
            int h = RowH;
            int rws = height / h;

            int ty = y;
            for (int i = 0; i < rws; i++)
            {
                ty = y + RowH * i;

            }
            g.DrawRectangle(new Pen(Color.Black), new Rectangle(x, y, w, RowH * rws));
            ty = ty + RowH;
            for (int c = 0; c < head.Length; c++)
            {
                string[] aa = head[c].Split('^');
                int cw = int.Parse(aa[4]);
                if (cw < 5) cw = 0;
                x = x + cw;
                g.DrawLine(new Pen(Color.Black), x, y, x, ty);
            }

        }
        private void DrawTxt(XmlNode td, Graphics g)
        {

            if (td == null) return;

            int x, y, width, height;
            x = int.Parse(td.Attributes["left"].Value) - cx;
            y = int.Parse(td.Attributes["top"].Value) - cy;
            if (td.ParentNode.Name != "PageHeadFoot")
                y = y - PagOffY;

            width = int.Parse(td.Attributes["width"].Value);
            height = int.Parse(td.Attributes["height"].Value);
            if (SetLeft != 0) //如果是设置后打印
            {

                if (BottomArray.Contains(td.Name))
                {
                    y = SetHeight - Convert.ToInt32(BottomArray[td.Name]);
                }
            }
            else
            {
                if ((y >= (InitHeight - InitBottom)) && (!BottomArray.Contains(td.Name)))
                {
                    BottomArray.Add(td.Name, InitHeight - y);
                }

            }
            if (dzqmprnpgd.Contains(td.Name))
            {

                Comm cmg = new Comm();
                string intext = dzqmprnpgd[td.Name].ToString(); // td.Attributes["text"].Value;
                if (intext.IndexOf('$') > -1)
                {
                    string[] imagespit = intext.Split('$');
                    for (int i = 0; i < imagespit.Length; i++)
                    {
                        string istr = imagespit[i].ToString();
                        Image img = cmg.StringToImage(istr);
                        int addh = 0;
                        int addw = 0;
                        if (qmprnorientation == 1)
                        {
                            addw = qmwildth + 4;
                        }
                        if (qmprnorientation == 0)
                        {
                            addh = qmheight + 2;
                        }
                        if (blackflag == "Y")
                        {
                            Bitmap bmp = changecolor(img, 4);
                            g.DrawImage(bmp, x + addw * i, y + i * addh, qmwildth, qmheight);
                        }
                        else
                        {
                            g.DrawImage(img, x + addw * i, y + i * addh, qmwildth, qmheight);
                        }
                    }
                    return;


                }
                else
                {
                    Image img = cmg.StringToImage(intext);
                    if (blackflag == "Y")
                    {
                        Bitmap bmp = changecolor(img, 4);
                        g.DrawImage(bmp, x, y, qmwildth, qmheight);
                    }
                    else
                    {
                        g.DrawImage(img, x, y, qmwildth, qmheight);
                    }
                    return;
                }
            }

            if (td.Name.Substring(0, 1) == "J")
            {
                Comm cmg = new Comm();

                Image img = cmg.StringToImage(td.InnerText);
                g.DrawImage(img, x, y, width, height);
                return;
            }
            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
            {
                // style |= FontStyle.Underline;
                int lx1 = x;
                int lx2 = lx1 + width;
                int ly1 = y + height;
                int ly2 = ly1;
                g.DrawLine(new Pen(Color.Black), lx1, ly1, lx2, ly2);
            }
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            float fontsize = float.Parse(td.Attributes["fontsize"].InnerText); ;
            if (xmlprndoc["Root"]["PageHeadFoot"] != null)
            {
                if (xmlprndoc["Root"]["PageHeadFoot"][td.Name] != null)
                {
                    fontsize = float.Parse(td.Attributes["fontsize"].InnerText);
                }
                else
                {
                    fontsize = float.Parse(td.Attributes["fontsize"].InnerText);
                }
            }
            Font font = new Font(td.Attributes["fontname"].InnerText,
                fontsize, style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();
            // 设置对齐方式
            if (td.Attributes["align"] != null)
            {
                switch (td.Attributes["align"].InnerText)
                {
                    case "center":
                        sf.LineAlignment = StringAlignment.Center;
                        break;
                    case "right":
                        sf.LineAlignment = StringAlignment.Center;
                        break;
                    default:
                        sf.LineAlignment = StringAlignment.Far;
                        break;
                }
            }
            else
            {
                sf.LineAlignment = StringAlignment.Near;
            }

            RectangleF rect = new RectangleF((float)x, (float)y,
                (float)width, (float)height);
            string txt = td.Attributes["text"].Value;

            float stwidth = g.MeasureString(txt, font).Width;
            if ((stwidth - rect.Width) > 0)
            {
                float cc = stwidth - rect.Width;
                if (cc < 3)
                {
                    rect.Width = rect.Width + 3;

                }
            }
            if (td.Name.IndexOf("Itm") != -1)
            {
                int end = td.Name.IndexOf("Itm");
                string itm = td.Name.Substring(0, end);
                int indx = int.Parse(td.Name.Substring(td.Name.Length - 1));

                if (prnData[itm] != null)
                {
                    string sutxt = prnData[itm].ToString();
                    string typ = DataTyp[itm].ToString().Substring(0, 1);
                    bool flag = false;
                    switch (typ)
                    {
                        case "O":
                            flag = getsel(sutxt, indx);
                            break;
                        case "M":
                            flag = getmulsel(sutxt, indx);

                            break;
                        default:
                            break;

                    }
                    if (flag == true)
                    {
                        if (typ == "O")
                        {
                            txt = "√ " + txt;
                        }
                        else
                        {
                            txt = txt + " √";
                        }
                    }
                }
            }
            bool flagprn = false;
            if (prnData[td.Name] != null)
            {
                txt = prnData[td.Name].ToString();
                string typ = DataTyp[td.Name].ToString().Substring(0, 1);
                switch (typ)
                {
                    case "I":
                        if (txt != "")
                        {
                            string[] it = txt.Split('|');
                            if (it.Length > 1)
                            {
                                txt = it[1];
                            }
                        }
                        break;
                    case "O":
                        txt = gettxt(txt);
                        break;
                    case "M":
                        txt = getmulseltxt(txt);
                        break;
                    case "T":

                        string[] aa = txt.Split('$');
                        string[] dd = aa[0].Split('|');
                        if (txt == "") break;
                        if (dd.Length > 3)
                        {
                            printTable2(td, g, txt);
                        }
                        else
                        {
                            PrintTable(td, g, txt);
                        }
                        flagprn = true;
                        break;
                    default:
                        txt = prnData[td.Name].ToString();
                        break;

                }

            }
            if (DataHash.Contains(td.Name))
            {
                g.DrawString(DataHash[td.Name].ToString(), font, brush, rect, sf);
                return;
            }
            if (td.Name.IndexOf("PageNo") != -1) txt = (Pages + 1).ToString();
            if (flagprn == false) g.DrawString(txt, font, brush, rect, sf);
            flagprn = false;
        }
        private void DrawTxt1(XmlNode td, Graphics g)
        {

            if (td == null) return;

            int x, y, width, height;
            x = int.Parse(td.Attributes["left"].Value) - cx;
            y = int.Parse(td.Attributes["top"].Value) - cy;
            if (td.ParentNode.Name != "PageHeadFoot")
                y = y - PagOffY;

            width = int.Parse(td.Attributes["width"].Value);
            height = int.Parse(td.Attributes["height"].Value);

            if (SetLeft != 0) //如果是设置后打印
            {

                if (BottomArray.Contains(td.Name))
                {
                    y = SetHeight - Convert.ToInt32(BottomArray[td.Name]);
                }
            }
            else
            {
                if ((y >= (InitHeight - InitBottom)) && (!BottomArray.Contains(td.Name)))
                {
                    BottomArray.Add(td.Name, InitHeight - y);
                }

            }

            if (td.Name.Substring(0, 1) == "J")
            {
                Comm cmg = new Comm();

                Image img = cmg.StringToImage(td.InnerText);
                g.DrawImage(img, x, y, width, height);
                return;
            }
            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            float fontsize = float.Parse(td.Attributes["fontsize"].InnerText); ;
            if (xmlprndoc["Root"]["PageHeadFoot"] != null)
            {
                if (xmlprndoc["Root"]["PageHeadFoot"][td.Name] != null)
                {
                    fontsize = float.Parse(td.Attributes["fontsize"].InnerText);
                }
                else
                {
                    fontsize = float.Parse(td.Attributes["fontsize"].InnerText);
                }
            }
            Font font = new Font(td.Attributes["fontname"].InnerText,
                fontsize, style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();
            // 设置对齐方式
            if (td.Attributes["align"] != null)
            {
                switch (td.Attributes["align"].InnerText)
                {
                    case "center":
                        sf.LineAlignment = StringAlignment.Center;
                        break;
                    case "right":
                        sf.LineAlignment = StringAlignment.Center;
                        break;
                    default:
                        sf.LineAlignment = StringAlignment.Far;
                        break;
                }
            }
            else
            {
                sf.LineAlignment = StringAlignment.Near;
            }

            RectangleF rect = new RectangleF((float)x, (float)y,
                (float)width, (float)height);
            string txt = td.Attributes["text"].Value;

            float stwidth = g.MeasureString(txt, font).Width;
            if ((stwidth - rect.Width) > 0)
            {
                float cc = stwidth - rect.Width;
                if (cc < 3)
                {
                    rect.Width = rect.Width + 3;

                }
            }
            if (td.Name.IndexOf("Itm") != -1)
            {
                int end = td.Name.IndexOf("Itm");
                string itm = td.Name.Substring(0, end);
                int indx = int.Parse(td.Name.Substring(td.Name.Length - 1));

                if (prnData[itm] != null)
                {
                    string sutxt = prnData[itm].ToString();
                    string typ = DataTyp[itm].ToString().Substring(0, 1);
                    bool flag = false;
                    switch (typ)
                    {
                        case "O":
                            flag = getsel(sutxt, indx);
                            break;
                        case "M":
                            flag = getmulsel(sutxt, indx);

                            break;
                        default:
                            break;

                    }
                    if (flag == true)
                    {
                        if (typ == "O")
                        {
                            txt = "√ " + txt;
                        }
                        else
                        {
                            txt = txt + " √";
                        }
                    }
                }
            }
            bool flagprn = false;
            if (prnData[td.Name] != null)
            {
                txt = prnData[td.Name].ToString();
                string typ = DataTyp[td.Name].ToString().Substring(0, 1);
                switch (typ)
                {
                    case "I":
                        if (txt != "")
                        {
                            string[] it = txt.Split('|');
                            txt = it[1];
                        }
                        break;
                    case "O":
                        txt = gettxt(txt);
                        break;
                    case "M":
                        txt = getmulseltxt(txt);
                        break;
                    case "T":

                        //string[] aa = txt.Split('$');
                        //string[] dd=aa[0].Split('|');
                        //if (dd.Length > 3)
                        //{
                        //    printTable2(td, g, txt);
                        //}
                        //else
                        //{
                        //    PrintTable(td, g, txt);
                        //}
                        //flagprn = true;
                        break;
                    default:
                        txt = prnData[td.Name].ToString();
                        break;

                }

            }
            if (DataHash.Contains(td.Name))
            {
                g.DrawString(DataHash[td.Name].ToString(), font, brush, rect, sf);
                return;
            }
            if (td.Name.IndexOf("PageNo") != -1) txt = (Pages + 1).ToString();
            if (flagprn == false) g.DrawString(txt, font, brush, rect, sf);
            flagprn = false;
        }
        private void PrintTable(XmlNode td, Graphics g, string txt)
        {
            int x, y, width, height;
            x = int.Parse(td.Attributes["left"].Value) - cx;
            y = int.Parse(td.Attributes["top"].Value) - cy;
            if (td.ParentNode.Name != "PageHeadFoot")
                y = y - PagOffY;

            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            // g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            Font font = new Font(td.Attributes["fontname"].InnerText,
                float.Parse(td.Attributes["fontsize"].InnerText), style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            if (txt == "")
            {
                return;
            }
            string[] tm = txt.Split('$');
            string[] ddata = tm[0].Split('\\');

            string[] titl = ddata[0].Split('^');
            string[] rowh = tm[1].Split('^');
            string[] colw = tm[2].Split('^');
            // string[] data = ddata[0].Split('^');
            string[] data = tm[3].Split('^');


            for (int i = 0; i < titl.Length; i++)
            {
                if ((titl[i] == "") || (titl[i] == "\\")) continue;
                Point location = new Point(x, y);
                int w = int.Parse(colw[i]);
                int h = 30;
                Size recsize = new Size(w, h);
                Rectangle rect = new Rectangle(location, recsize);
                g.DrawString(titl[i], font, brush, rect, sf);
                g.DrawRectangle(new Pen(Color.Black, 1), rect);
                x = x + w;
            }
            y = y + 30;
            x = int.Parse(td.Attributes["left"].Value) - cx;
            int dx = x;
            int dy = y;
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == "") continue;
                string[] coldata = data[i].Split('!');
                int h = int.Parse(rowh[i + 1]);

                for (int j = 0; j < coldata.Length; j++)
                {
                    if (colw[j] == "") continue;
                    Point location = new Point(dx, dy);
                    int w = int.Parse(colw[j]);
                    Size recsize = new Size(w, h);
                    Rectangle rect = new Rectangle(location, recsize);
                    g.DrawString(coldata[j], font, brush, rect, sf);
                    g.DrawRectangle(new Pen(Color.Black, 1), rect);
                    dx = dx + w;

                }
                dy = dy + h;
                dx = x;
            }

        }
        private void PrintShape(XmlNode shap, Graphics g, int pagheight)
        {
            foreach (XmlNode xd in shap.ChildNodes)
            {
                string typ = xd.Attributes["typ"].Value;

                string[] par1 = xd.Attributes["P1"].Value.Split(',');
                string[] par2 = xd.Attributes["P2"].Value.Split(',');
                string[] colorarr = xd.Attributes["pencolor"].Value.Split(',');
                string width = xd.Attributes["penwidth"].Value;
                Point p1, p2;
                p1 = new Point(Convert.ToInt16(par1[0]), Convert.ToInt16(par1[1]));
                p2 = new Point(Convert.ToInt16(par2[0]), Convert.ToInt16(par2[1]));
                Color pencolor = Color.FromArgb(Convert.ToByte(colorarr[0]), Convert.ToByte(colorarr[1]), Convert.ToByte(colorarr[2]));
                int penwidth = Convert.ToInt16(width);
                int x1, x2, y1, y2;
                x1 = p1.X - cx; x2 = p2.X - cx; y1 = p1.Y - cy; y2 = p2.Y - cy;

                int maxY = 0;
                if (y1 < y2) maxY = y1;
                if (y1 > y2) maxY = y2;
                if (y1 == y2) maxY = y1;
                if (maxY < PagOffY) continue;
                if (maxY > pagheight * (Pages + 1)) continue;
                y1 = y1 - PagOffY; y2 = y2 - PagOffY;

                if (x1 > x2)  //直线小误差修正
                {
                    int exval = x1 - x2;
                    if (exval < 5) x2 = x1;
                }
                else
                {
                    int exval = x2 - x1;
                    if (exval < 5) x1 = x2;
                }
                if (y1 > y2) //直线小误差修正
                {
                    int exval = y1 - y2;
                    if (exval < 5) y2 = y1;
                }
                else
                {
                    int exval = y2 - y1;
                    if (exval < 5) y1 = y2;
                }
                switch (typ)
                {
                    case "NurEmrMaintain.LineShape":
                        g.DrawLine(new Pen(pencolor, penwidth), new Point(x1, y1), new Point(x2, y2));
                        break;
                    case "NurEmrMaintain.RectangleShape":
                        g.DrawRectangle(new Pen(pencolor, penwidth), x1, y1, (x2 - x1), (y2 - y1));
                        break;
                    case "NurEmrMaintain.EllipseShape":
                        g.DrawEllipse(new Pen(pencolor, penwidth), x1, y1, (x2 - x1), (y2 - y1));
                        break;
                    case "NurEmrMaintain.CircleShape":
                        int r = (int)Math.Pow(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2), 0.5);
                        g.DrawEllipse(new Pen(pencolor, penwidth), x1 - r, y1 - r, 2 * r, 2 * r);
                        break;
                }
            }
        }
        private void pd_PrintPage3(object sender, PrintPageEventArgs ev)
        {
            //int rowcount = 0;
            //int rowsperpage = 0;

            bool HasMorePages = false;
            int yh = 0;
            //ev.Graphics.DrawRectangle(new Pen(Color.Black), ev.MarginBounds);
            Rectangle rect = new Rectangle(new Point(0, 0), ev.PageBounds.Size);
            PagOffY = (Pages - curPages) * (ev.PageBounds.Height);
            // ev.Graphics.DrawRectangle(new Pen(Color.Red ),rect);
            PrintShape(xmlprndoc["Root"]["SHAPES"], ev.Graphics, ev.PageBounds.Size.Height);
            if (xmlprndoc["Root"]["PageHeadFoot"] != null)
            {
                foreach (XmlNode xn in xmlprndoc["Root"]["PageHeadFoot"].ChildNodes)
                {
                    if (xn.Name.IndexOf("HFNOD", 0) == -1)
                    {
                        if (xmlprndoc["Root"]["InstanceData"][xn.Name] != null) DrawTxt(xmlprndoc["Root"]["InstanceData"][xn.Name], ev.Graphics);
                    }
                }
            }
            ItmCount = 0;
            do
            {

                XmlNode xnod = xmlprndoc["Root"]["InstanceData"].ChildNodes[ItmCount];
                ItmCount++;
                if (xnod == null) break;
                yh = int.Parse(xnod.Attributes["top"].Value);
                if (yh < PagOffY) continue;
                if (yh > ev.PageBounds.Height * (Pages - curPages + 1)) continue;
                if ((xnod.Attributes["type"].Value == "System.Windows.Forms.Label") || (xnod.Attributes["type"].Value == "System.Windows.Forms.PictureBox"))
                {
                    DrawTxt(xnod, ev.Graphics);
                    PrnCount++;
                }
                //  yh = yh - PagOffY;
                //if (yh > ev.PageBounds.Height) break;

            } while (ItmCount < xmlprndoc["Root"]["InstanceData"].ChildNodes.Count);
            if (Pages < EdP)
            {
                HasMorePages = true;
                PagOffY = PagOffY + ev.PageBounds.Height;

            }
            if (HasMorePages)
            {
                Pages++;
                if ((Pages == EdP) && (EdP != 0))
                {
                    //stPage = Pages;

                    HasMorePages = false;
                    // Pages = 0;
                    Row = 0;
                    stRow = 0;
                    EdP = 0;
                    ev.HasMorePages = HasMorePages;
                    PrnPreView.Dispose();
                    PrnDiaglog.Dispose();
                    pDoc.Dispose();
                    PrnFlag++;
                    return;
                }

            }
            else
            {
                printpagecount = Pages + 1;
                ItmCount = 0;
                PagOffY = 0;
                PrnCount = 0;

                Pages = curPages;

            }
            ev.HasMorePages = HasMorePages;

        }
        //评估单类打印逻辑
        private void makepgdpage(Graphics g, ref bool HasMorePages)
        {
            int yh = 0;

            Rectangle rect = new Rectangle(new Point(0, 0), pDoc.DefaultPageSettings.Bounds.Size);
            //PrintShape(xmlprndoc["Root"]["SHAPES"], g, pdoc.PageBounds.Size.Height);
            PrintShape(xmlprndoc["Root"]["SHAPES"], g, pDoc.DefaultPageSettings.Bounds.Size.Height);
            if (xmlprndoc["Root"]["PageHeadFoot"] != null)
            {
                foreach (XmlNode xn in xmlprndoc["Root"]["PageHeadFoot"].ChildNodes)
                {
                    if (xn.Name.IndexOf("HFNOD", 0) == -1)
                    {
                        if (xmlprndoc["Root"]["InstanceData"][xn.Name] != null) DrawTxt(xmlprndoc["Root"]["InstanceData"][xn.Name], g);
                    }
                }
            }
            ItmCount = 0;
            do
            {

                XmlNode xnod = xmlprndoc["Root"]["InstanceData"].ChildNodes[ItmCount];
                ItmCount++;
                if (PgdPrintedArray.Contains(xnod.Name)) continue; //已经打印过
                if (xnod == null) break;
                yh = int.Parse(xnod.Attributes["top"].Value);
                if (yh < PagOffY) continue;
                if (yh > pDoc.DefaultPageSettings.Bounds.Height * (Pages - curPages + 1)) continue;
                // if (yh > ev.PageBounds.Height * (Pages - curPages + 1)) continue;
                if ((xnod.Attributes["type"].Value == "System.Windows.Forms.Label") || (xnod.Attributes["type"].Value == "System.Windows.Forms.PictureBox"))
                {
                    DrawTxt(xnod, g);
                    PrnCount++;
                    if (!PgdPrintedArray.Contains(xnod.Name))
                    {
                        PgdPrintedArray.Add(xnod.Name);
                    }
                }
                //  yh = yh - PagOffY;
                //if (yh > ev.PageBounds.Height) break;

            } while (ItmCount < xmlprndoc["Root"]["InstanceData"].ChildNodes.Count);
            if (PrnCount < xmlprndoc["Root"]["InstanceData"].ChildNodes.Count)
            {
                HasMorePages = true;
                PagOffY = PagOffY + pDoc.DefaultPageSettings.Bounds.Height;
                //PagOffY = PagOffY + ev.PageBounds.Height;
            }
        }
        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            //int rowcount = 0;
            //int rowsperpage = 0;
            if (EdP > 0)
            {
                // MessageBox.Show(EdP.ToString());
                pd_PrintPage3(sender, ev);
                return;
            }
            bool HasMorePages = false;
            makepgdpage(ev.Graphics, ref HasMorePages);  //画图
            if (HasMorePages)
            {
                Pages++;
            }
            else
            {
                printpagecount = Pages + 1;
                ItmCount = 0;
                PagOffY = 0;
                PrnCount = 0;
                Pages = curPages;
                PgdPrintedArray.Clear(); //打印记录

            }
            ev.HasMorePages = HasMorePages;

        }
        private int PrintHead(XmlNode td, Graphics g, string txt)
        {
            int x, y, width, height;
            x = int.Parse(td.Attributes["left"].Value);
            y = int.Parse(td.Attributes["top"].Value);
            if (tabx == 0) tabx = x;
            if (taby == 0) taby = y;
            //anhui20120515
            if (Pages - curPages > 0)
            {
                //tabx = MargLeft;
                taby = MargTop;
            }
            if ((Pages==curPages)&&(curPages> 0)&&(MthArr!="")&&(Parrm!="")) //hunhemo
            {
                //tabx = MargLeft;
                taby = MargTop;
            }
            if (td.ParentNode.Name != "PageHeadFoot")
                y = y - PagOffY;

            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            // g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            Font font = new Font(td.Attributes["fontname"].InnerText,
                float.Parse(td.Attributes["fontsize"].InnerText), style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            sf.Alignment = StringAlignment.Center;
            string[] titl = txt.Split('|');
            int posx = 0, posy = 0;
            int yy = 0;

            if (curhead != "") //表头变更当前表头id
            {
                string parr = EmrCode + "!" +curhead;
                try
                {
                                         
                        string headstr = Comm.DocServComm.GetData("NurEmr.webheadchange:GetLableRecByID", "par:" + parr + "^");
                        if ((headstr != null) && (headstr != ""))
                        {
                            setlablehead(headstr);
                        }
                    
                }
                catch(Exception ex){
                    MessageBox.Show("NurEmr.webheadchange:GetLableRecByID("+parr+")"+ex.Message);
                }

               
            }
            for (int i = 0; i < titl.Length; i++)
            {
                if (titl[i] == "") continue;
                string[] hh = titl[i].Split('^');
                if (DataLable.ContainsKey(hh[8]))  //表头变更打印表头
                {
                    hh[9] = DataLable[hh[8]].ToString();
                }
                int sind = hh[9].IndexOf("_");
                if (sind != -1) hh[9] = hh[9].Replace("_", ""); //20130311替换打印模板的"_"
                //int sind = hh[9].IndexOf("_",(hh[9].Length-1) );
                // if (sind !=-1) hh[9] = hh[9].Substring(0, sind);
                posx = int.Parse(hh[5]) + tabx; posy = int.Parse(hh[6]) + taby;
                Point location = new Point(posx, posy);
                int w = int.Parse(hh[1]);
                int h = int.Parse(hh[0]);
                if (w < 5) w = 0;
                if (PrintCareDateLine == "N") //日期时间列数据为空不打印该列下面的横线
                {
                    if (hh[8] == "CareDate")
                    {
                        CareDateWidth = w;
                    }
                    if (hh[8] == "CareTime")
                    {
                        CareTimeWidth = w;
                    }
                }
                Size recsize = new Size(w, h);

                //char[] prntxt =hh[9].ToCharArray();
                //for (int c = 0; c < prntxt.Length; c++)
                //{
                //    string aa= prntxt[c].ToString();
                //}
                if ((Pages == curPages) && (xuflag == "1") && (Startrow > 0))
                { }
                else
                {
                    Rectangle rect = new Rectangle(location, recsize);
                    if (w != 0) g.DrawString(hh[9], font, brush, rect, sf);
                    g.DrawRectangle(new Pen(Color.Black, 1), rect);
                }
                yy = posy + h;
                //x = x + w;
            }
            //anhui20120515
            tabx = x;
            taby = y;
            return yy;
        }
        private int PrintHeadN(XmlNode td, Graphics g, string txt)
        {
            int x, y, width, height;
            x = int.Parse(td.Attributes["left"].Value) - cx;
            y = int.Parse(td.Attributes["top"].Value) - cy;
            if (tabx == 0) tabx = x;
            if (taby == 0) taby = y;
            //anhui20120515
            if (Pages > 0)
            {
                tabx = MargLeft;
                taby = MargTop;
            }

            if (td.ParentNode.Name != "PageHeadFoot")
                y = y - PagOffY;

            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            // g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            Font font = new Font(td.Attributes["fontname"].InnerText,
                float.Parse(td.Attributes["fontsize"].InnerText), style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            sf.Alignment = StringAlignment.Center;
            string[] titl = txt.Split('|');
            int posx = 0, posy = 0;
            int yy = 0;
            for (int i = 0; i < titl.Length; i++)
            {
                if (titl[i] == "") continue;
                string[] hh = titl[i].Split('^');
                if (DataLable.ContainsKey(hh[8]))  //表头变更打印表头
                {
                    hh[9] = DataLable[hh[8]].ToString();
                }
                int sind = hh[9].IndexOf("_", (hh[9].Length - 1));
                if (sind != -1) hh[9] = hh[9].Substring(0, sind);
                posx = int.Parse(hh[5]) + x; posy = int.Parse(hh[6]) + y;
                Point location = new Point(posx, posy);
                int w = int.Parse(hh[1]);
                int h = int.Parse(hh[0]);
                if (w < 5) w = 0;
                Size recsize = new Size(w, h);

                //char[] prntxt =hh[9].ToCharArray();
                //for (int c = 0; c < prntxt.Length; c++)
                //{
                //    string aa= prntxt[c].ToString();
                //}

                Rectangle rect = new Rectangle(location, recsize);
                if (w != 0) g.DrawString(hh[9], font, brush, rect, sf);
                g.DrawRectangle(new Pen(Color.Black, 1), rect);
                yy = posy + h;
                //x = x + w;
            }
            //anhui20120515
            tabx = x;
            taby = y;
            return yy;
        }
        private void printTable2(XmlNode td, Graphics g, string txt)
        {
            int x, y, width, height;
            x = int.Parse(td.Attributes["left"].Value) - cx;
            y = int.Parse(td.Attributes["top"].Value) - cy;
            if (td.ParentNode.Name != "PageHeadFoot")
                y = y - PagOffY;

            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            // g.FillRectangle(brush, x + 1, y + 1, width, height);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            Font font = new Font(td.Attributes["fontname"].InnerText,
                float.Parse(td.Attributes["fontsize"].InnerText), style, fUnit);
            brush = new SolidBrush(Color.FromName(td.Attributes["fontcolor"].InnerText));
            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            string[] tm = txt.Split('$');
            // string[] titl = tm[0].Split('^');
            string[] ddata = tm[0].Split('\\');
            string[] rowh = tm[1].Split('^');
            string[] colw = tm[2].Split('^');
            string[] data = tm[3].Split('^');
            //anhui20120515

            if (xmlprndoc["Root"].Attributes["DataType"].Value == "xml")
            {
                y = PrintHeadN(td, g, ddata[0]);
            }
            else
            {
                y = PrintHead(td, g, ddata[0]);
            }
            x = int.Parse(td.Attributes["left"].Value) - cx;
            int dx = x;
            int dy = y;
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == "") continue;
                string[] coldata = data[i].Split('!');
                int h = int.Parse(rowh[i + 1]);

                for (int j = 0; j < coldata.Length; j++)
                {
                    if (colw[j] == "") continue;
                    Point location = new Point(dx, dy);
                    int w = int.Parse(colw[j]);
                    Size recsize = new Size(w, h);
                    Rectangle rect = new Rectangle(location, recsize);
                    g.DrawString(coldata[j], font, brush, rect, sf);
                    g.DrawRectangle(new Pen(Color.Black, 1), rect);
                    dx = dx + w;

                }
                dy = dy + h;
                dx = x;
            }


        }
        public void SetTable(DataTable table, Hashtable rowlinedata, string[] headstr, int rows, bool flag)
        {

            DataRow row;
            if (flag == true) rows = 1;
            for (int r = 0; r < rows; r++)
            {
                row = table.NewRow();
                for (int j = 0; j < head.Length; j++)
                {
                    //if Res.Get
                    if (head[j] == null) continue;
                    if (head[j] == "") continue;
                    string[] th = head[j].Split('^');

                    ArrayList array = (ArrayList)rowlinedata[th[0]];
                    if (array == null)
                    {
                        array = new ArrayList();
                    }

                    if (array.Count <= r)
                    {
                        row[j] = "";
                    }
                    else
                    {
                        row[j] = array[r];
                    }

                }
                table.Rows.Add(row);
            }

        }
        private void pd_PrintPage11(object sender, PrintPageEventArgs ev)
        {
            bool HasMorePages = false;
            int yh = 0;
            //ev.Graphics.DrawRectangle(new Pen(Color.Black), ev.MarginBounds);
            //Rectangle rect = new Rectangle(new Point(0, 0), ev.PageBounds.Size);
            //MessageBox.Show(Pages.ToString());
            // ev.Graphics.DrawRectangle(new Pen(Color.Red ),rect);
            bool headflag = false;
            //if (Pages == 0)  //anhui20120515
            if (Pages == curPages)  //20130311集中打印修改
                PrintShape(xmlprndoc["Root"]["SHAPES"], ev.Graphics, ev.PageBounds.Size.Height);

            if (HLhead.Contains(Pages + 1))
            {
                HFCaption["LHead"] = HLhead[Pages + 1].ToString();
            }
            if (PageDiagNod != null)
            {
                if (PageDiagH.Contains(Pages + 1)) PageDiagNod.Attributes["text"].Value = PageDiagH[Pages + 1].ToString();
            }
            if (PageLocNod != null)
            {
                if (PageLocH.Contains(Pages + 1)) PageLocNod.Attributes["text"].Value = PageLocH[Pages + 1].ToString();
            }
            if (PageBedNod != null)
            {
                if (PageBedH.Contains(Pages + 1)) PageBedNod.Attributes["text"].Value = PageBedH[Pages + 1].ToString();
            }
            // if ((Pages < EdP) && (EdP != 0)) HasMorePages = true;
            int POSY = 0;
            if (xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"] != null)
            {
                headflag = true;
            }
            string cellDatanod = "";  //数据节点
            string xHn = "";   //表头节点
            foreach (XmlNode xs in xmlprndoc["Root"]["MetaData"].ChildNodes)
            {
                if (xs.Attributes["Rel"] != null)
                {
                    string[] a = xs.Attributes["Rel"].Value.Split('.');
                    if (a[1].Substring(0, 1) == "T") cellDatanod = xs.Name;
                }
            }
            foreach (XmlNode xs in xmlprndoc["Root"]["InstanceData"].ChildNodes)
            {
                if (xs.Attributes["text"] != null)
                {
                    string a = xs.Attributes["text"].Value;
                    if (a != "")
                    {
                        if (a.Substring(0, 1) == "T")
                        {
                            if (xs.Name.Substring(0, 1) == "B")
                            {
                                xHn = xs.Name;
                            }
                        }
                    }
                }
            }
            string TableText = "";
            if ((xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"] != null) && (stPrintPos == 0))
            {

                if (xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"].Value != "")
                {
                    if (xmlprndoc["Root"]["THEAD"] != null)
                    {
                        POSY = drawHead(ev.Graphics, xmlprndoc["Root"]["THEAD"]);
                    }
                    else
                    {
                        if (xHn != "")
                        {  //anhui20120515
                            TableText = xmlprndoc["Root"]["InstanceData"][xHn].Attributes["text"].Value;
                            POSY = PrintHead(xmlprndoc["Root"]["InstanceData"][xHn], ev.Graphics, xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"].Value);
                        }
                        else
                        {
                            POSY = PrintHead(xmlprndoc["Root"]["InstanceData"].ChildNodes[0], ev.Graphics, xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"].Value);
                        }
                    }
                    headflag = true;
                }

            }
            if (headflag == false)
            {
                PrintShape(xmlprndoc["Root"]["SHAPES"], ev.Graphics, ev.PageBounds.Size.Height);
            }
            if ((xmlprndoc["Root"]["PageHeadFoot"] != null) && (stPrintPos == 0))
            {
                //int hf = 0;
                /// int nodes = 0;
                foreach (XmlNode xn in xmlprndoc["Root"]["PageHeadFoot"].ChildNodes)
                {
                    if (xn.Name.IndexOf("HFNOD", 0) == -1)
                    {
                        if (xmlprndoc["Root"]["InstanceData"][xn.Name] != null)
                        {
                            string aa = xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value;
                            if (HFCaption.ContainsKey(aa))
                            {
                                xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value = HFCaption[aa].ToString();
                            }

                            DrawTxt(xmlprndoc["Root"]["InstanceData"][xn.Name], ev.Graphics);
                            xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value = aa;

                        }
                    }
                }
            }
            if (headflag == true)
            {
                //if (Pages == 0)  //anhui20120515
                if (Pages == curPages)  //20130311集中打印修改
                {
                    foreach (XmlNode xnod in xmlprndoc["Root"]["InstanceData"].ChildNodes)
                    {
                        if (xnod.Attributes["text"].Value.IndexOf(TableText) == -1)
                        {
                            string valueflag = xnod.Attributes["text"].Value;
                            if ((valueflag != "LHead") && (valueflag != "RHead") && (valueflag != "LFoot") && (valueflag != "RFoot"))
                            {
                                DrawTxt1(xnod, ev.Graphics);
                            }
                        }
                    }
                }
            }
            XmlNode cellnod;
            if (cellDatanod != "") cellnod = xmlprndoc["Root"]["InstanceData"][cellDatanod];
            else cellnod = xmlprndoc["Root"]["InstanceData"][xHn];
            float fontnum = float.Parse(cellnod.Attributes["fontsize"].Value); //字体尺寸
            float RowH = float.Parse(cellnod.Attributes["height"].Value); //行高
            int y1 = POSY;
            int x1 = tabx;
            if (stRow != 0) y1 = stPrintPos - (int)(stRow * RowH);
            //int PrnH =(int)RowH  * int.Parse(HPagRow[Pages+1].ToString());//(ev.MarginBounds.Y + ev.MarginBounds.Height - y1);  //
            int PrnH = ev.MarginBounds.Y + ev.MarginBounds.Height - y1; //最后一页的格子都打出来
            ItmCount = 0;
            if ((Row == 0) && (stPrintPos != 0)) y1 = stPrintPos;
            for (int i = 0; i < tableBak.Columns.Count; i++)
            {
                if (head[i] == "") continue;
                string[] th = head[i].Split('^');
                string[] afth = null;
                if (i < (tableBak.Columns.Count - 1)) afth = head[i + 1].Split('^');
                int tx = 0, tw = 0, afx = 0;
                if (afth != null) afx = int.Parse(afth[1]);
                else
                {
                    afx = tabW;
                }
                tx = int.Parse(th[1]); tw = afx - x1;
                head[i] = head[i] + "^" + tw.ToString();
                x1 = x1 + tw;
            }

            int singcount = 0;
            if (stRow != 0) singcount = stRow;
            int PH = ev.PageBounds.Size.Height - pDoc.DefaultPageSettings.Margins.Bottom;//页面可利用高度
            PageRows = PrnH / (int)RowH; ///
            printPagesize = PageRows;
            if (stPrintPos == 0)
            {
                DrawLines(tabx, y1, tabW - tabx, PrnH, ev.Graphics, (int)RowH, head);
            }
            if (dxflag == 0)
            {
                //不管多少行记录都画线
                for (int k = 0; k < PageRows; k++)
                {
                    int hh = (int)RowH;
                    ev.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(tabx, y1 + (k + 1) * hh), new Point(tabW, y1 + (k + 1) * hh));
                }
            }
            do
            {
                x1 = tabx;
                int hh = (int)RowH;
                if (tableBak.Rows.Count == 0) break;
                if (Row >= tableBak.Rows.Count) break;

                int htt = 0;
                float RH; //
                //int retrows = 0;
                singcount = singcount + 1; //单页行数
                int stx = 0, sty = 0;
                string CareDate = "", CareTime = "";
                bool rowflag = false;
                if (BakRowH.Contains(Row))
                {
                    rowflag = true;
                    hh = (int)(RowH * 2);
                }

                for (int i = 0; i < tableBak.Columns.Count; i++)
                {
                    if (head[i] == "") continue;
                    string[] th = head[i].Split('^');
                    string[] afth = null;
                    if (i < (tableBak.Columns.Count - 1)) afth = head[i + 1].Split('^');
                    int tx = 0, tw = 0, afx = 0;
                    if (afth != null) afx = int.Parse(afth[1]);
                    else
                    {
                        afx = tabW;
                    }
                    tx = int.Parse(th[1]); tw = int.Parse(th[4]);
                    if (tw < 5) tw = 0; //列宽
                    Graphics g = ev.Graphics;

                    if (th[0] == "CareDateTim") CareDateTim = tableBak.Rows[Row][th[0]].ToString();
                    if (th[0] == "CareDate") CareDate = tableBak.Rows[Row][th[0]].ToString();
                    if (th[0] == "CareTime") CareTime = tableBak.Rows[Row][th[0]].ToString();
                    if ((CareDate != "") && (CareTime != ""))
                    {
                        CareDateTim = CareDate + "/" + CareTime;
                    }
                    //ArrayList array = (ArrayList)tableBak.Rows[Row][th[0]];
                    // ArrayList array = (ArrayList)rowlinedata[th[0]];
                    if (tw != 0)
                    {
                        XmlNode xalign = null;
                        if (xmlprndoc["Root"]["THEAD"] != null)
                        {
                            xalign = xmlprndoc["Root"]["THEAD"][th[0]];
                        }
                        if ((th[0] == "User"))
                        {
                            if (IsVerifyCALoc == "1")
                            {

                                /*
                                 string uid = tableBak.Rows[Row][th[0]].ToString();
                                string imageuser = Comm.DocServComm.GetData("web.DHCNurSignVerify:GetUserSignImage", "par:" + uid + "^");
                                if (imageuser != null)
                                {
                                    Comm cmg = new Comm();
                                    Image img = cmg.StringToImage(imageuser);
                                    g.DrawImage(img, x1 + 2, y1 + 5, qmwildth, qmheight);

                                }
                                else
                                {

                                    DrawString11(tableBak.Rows[Row][th[0]].ToString(), x1, y1, tw, htt, ev.Graphics, (int)RowH, cellnod, rowflag, xalign);
                                }
                                */
                                string uidsss = tableBak.Rows[Row][th[0]].ToString();
                                string[] useridstr = uidsss.Split(' ');
                                for (int i2 = 0; i2 < useridstr.Length; i2++)
                                {
                                    string uid = useridstr[i2];
                                    string imageuser = null;
                                    try
                                    {
                                        imageuser = Comm.DocServComm.GetData("web.DHCNurSignVerify:GetUserSignImage", "par:" + uid + "^");
                                    }
                                    catch(Exception ex) { 
                                      
                                    }
                                    if (imageuser != null)
                                    {
                                        Comm cmg = new Comm();
                                        Image img = cmg.StringToImage(imageuser);

                                        int addh = 0;
                                        int addw = 0;
                                        if (qmprnorientation == 1)
                                        {
                                            addw = qmwildth + qmhori;
                                        }
                                        if (qmprnorientation == 0)
                                        {
                                            addh = qmheight + qmport;
                                        }
                                        if (blackflag == "Y")
                                        {
                                            Bitmap bmp = changecolor(img, 4);
                                            g.DrawImage(bmp, x1 + qmleft + addw * i2, y1 + qmtop + addh * i2, qmwildth, qmheight);
                                        }
                                        else
                                        {
                                            //g.DrawImage(img, x1 + 2, yimage + 5 + i2 * 10, qmwildth, qmheight);
                                            g.DrawImage(img, x1 + qmleft + addw * i2, y1 + qmtop + addh * i2, qmwildth, qmheight);
                                            //g.DrawImage(img, x + addw * i, y + i * addh, qmwildth, qmheight);
                                        }
                                    }
                                    else
                                    {
                                        //DrawString1(array, x1, y1, tw, htt, ev.Graphics, (int)RowH, cellnod, rowflag, xalign);
                                        DrawString11(uid.ToString(), x1, y1, tw, htt, ev.Graphics, (int)RowH, cellnod, rowflag, xalign);

                                    }


                                }


                            }
                            else
                            {

                                DrawString11(tableBak.Rows[Row][th[0]].ToString(), x1, y1, tw, htt, ev.Graphics, (int)RowH, cellnod, rowflag, xalign);
                            }

                        }
                        else
                        {
                            // string PageLoc = Comm.DocServComm.GetData("web.DHCNurSignVerify:GetUserDesc", "par:"  + "^");
                            //DrawString1(array, x1, y1, tw, htt, ev.Graphics, (int)RowH, cellnod, rowflag, xalign);
                            DrawString11(tableBak.Rows[Row][th[0]].ToString(), x1, y1, tw, htt, ev.Graphics, (int)RowH, cellnod, rowflag, xalign);
                        }



                    }


                    if (i == 0)
                    {
                        stx = x1; sty = y1;
                    }
                    x1 = x1 + tw;


                }
                bool linflag = false;
                if (Row == tableBak.Rows.Count - 1)
                {

                    ev.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(stx, sty + hh), new Point(x1, sty + hh));


                }
                if ((Row < (tableBak.Rows.Count - 1))) //( Row !=0)&&
                {
                    if (tableBak.Columns.Contains("CareDateTim"))
                    {
                        string aa = tableBak.Rows[Row + 1]["CareDateTim"].ToString();
                        if ((aa != "") && (aa.IndexOf("-") != -1))
                        {
                            linflag = true;
                        }
                    }
                    if (tableBak.Columns.Contains("CareDate"))
                    { //合肥
                        string aa = tableBak.Rows[Row + 1]["CareDate"].ToString();
                        if ((aa != ""))
                        {
                            linflag = true;
                        }
                        if (aa == "") //如果日期为空，日期和时间列下的横线不要
                        {
                            if (tableBak.Columns.Contains("User"))
                            {
                                if (UserPrintDown == "N")
                                {
                                    string aa1 = tableBak.Rows[Row]["User"].ToString();
                                    aa1 = aa1.Replace(" ", "");
                                    string aa2 = tableBak.Rows[Row + 1]["User"].ToString();
                                    aa2 = aa2.Replace(" ", "");
                                    if ((aa2 != "") && (aa1 == ""))
                                    {
                                        ev.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(stx + CareDateWidth + CareTimeWidth, sty + hh), new Point(x1, sty + hh));
                                    }
                                }
                                else
                                {
                                    string aa3 = tableBak.Rows[Row]["User"].ToString();
                                    aa3 = aa3.Replace(" ", "");
                                    string aa4 = tableBak.Rows[Row + 1]["User"].ToString();
                                    aa4 = aa4.Replace(" ", "");
                                    if ((aa3 != "") && (aa4 == ""))
                                    {
                                        ev.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(stx + CareDateWidth + CareTimeWidth, sty + hh), new Point(x1, sty + hh));
                                    }
                                }
                            }
                        }
                    }
                    if (tableBak.Columns.Contains("Item101"))
                    { //病室报告
                        string aa = tableBak.Rows[Row + 1]["Item101"].ToString();
                        if ((aa != ""))
                        {
                            linflag = true;
                        }
                    }
                }

                if (linflag == true)
                    ev.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(stx, sty + hh), new Point(x1, sty + hh));
                // ev.Graphics.DrawLine(new Pen(Color.Black, 1), new Point(stx, sty ), new Point(x1, sty ));
                y1 = y1 + hh;

                Row++;
                if (singcount == int.Parse(HPagRow1[Pages + 1].ToString()))
                {
                    HasMorePages = true;
                    break;
                }

            } while ((Row < tableBak.Rows.Count) && (y1 < (ev.PageBounds.Size.Height - pDoc.DefaultPageSettings.Margins.Bottom)));

            if (HasMorePages == false)
            {
                Row = 0;
                string parm = EpisodeID + "^" + CareDateTim + "^" + y1.ToString() + "^" + Pages.ToString();
                printpagecount = Pages + 1;
                DrawRecLine(Pages + 1, ev.Graphics);
                // stPrintPos =y1;
                if (PreView == "1")
                {
                    Pages = stPage;
                    // y1 = stPrintPos;
                }
                else
                {
                    stPage = Pages;
                    stPrintPos = y1;
                    stRow = singcount;
                    PrnPreView.Dispose();
                }
                //首次是预览 将PreView=0，第二次是打印
                //if (PreView == "1") PreView = "0";
                // if (PreView == "0") PreView = "1";
                PreView = "0";
                PrnFlag++;
                PrintCareDateLine = "Y";
                CareDateWidth = 0;
                CareTimeWidth = 0;
                // HFCaption["LHead"] = ChangePageDiag;

                if (ChangePageDiag == "") HFCaption["LHead"] = LHeadCaption;
                // MessageBox.Show(PrnFlag.ToString() ); 

            }
            else
            {
                stPrintPos = 0;
                stRow = 0;
                Pages++;
                DrawRecLine(Pages, ev.Graphics);   //画记录的线
                // HFCaption["LHead"] = HLhead[Pages]; 
                //  HPagRow.Add(Pages, singcount);
                if ((Pages == EdP) && (EdP != 0))
                {
                    //stPage = Pages;

                    HasMorePages = false;
                    // Pages = 0;
                    Row = 0;
                    stRow = 0;
                    EdP = 0;
                    ev.HasMorePages = HasMorePages;
                    PrnPreView.Dispose();
                    PrnDiaglog.Dispose();
                    pDoc.Dispose();
                    PrnFlag++;
                    return;
                }
            }
            ev.HasMorePages = HasMorePages;

        }
        public static Bitmap changecolor(Image SImage, int p)
        {
            int Height = SImage.Height;
            int Width = SImage.Width;
            Bitmap bitmap = new Bitmap(Width, Height);
            Bitmap MyBitmap = (Bitmap)SImage;
            Color pixel1, pixel2;
            for (int x = 0; x < Width - 1; x++)
            {
                for (int y = 0; y < Height - 1; y++)
                {
                    int r = 0, g = 0, b = 0;
                    pixel1 = MyBitmap.GetPixel(x, y);
                    pixel2 = MyBitmap.GetPixel(x + 1, y + 1);
                    r = pixel1.R;// -pixel2.R + p;
                    g = pixel1.G;// -pixel2.G + p;
                    b = pixel1.B;//- pixel2.B + p;
                    if ((r == 255) & (g == 255) & (b == 255))
                    {

                    }
                    else
                    {
                        r = 0;
                        b = 0;
                        g = 0;
                    }
                    if (r > 255)
                        r = 255;
                    if (r < 0)
                        r = 0;
                    if (g > 255)
                        g = 255;
                    if (g < 0)
                        g = 0;
                    if (b > 255)
                        b = 255;
                    if (b < 0)
                        b = 0;
                    bitmap.SetPixel(x, y, Color.FromArgb(r, g, b));
                }
            }
            return bitmap;
        }

        string PrevHeadDR = "", CurrHeadDR = "";   //表头变换模板，换表头后换页 --20141201
        private void makepage(Graphics g, ref bool HasMorePages)
        {


            // ev.Graphics = gimage;

            printinfo = "";
            rowprintinfo = ""; //每页每条记录打印信息
            HasMorePages = false;
            int yh = 0;

            //ev.Graphics.DrawRectangle(new Pen(Color.Black), ev.MarginBounds);
            //Rectangle rect = new Rectangle(new Point(0, 0), ev.PageBounds.Size);
            //MessageBox.Show(Pages.ToString());
            // ev.Graphics.DrawRectangle(new Pen(Color.Red ),rect);
            bool headflag = false;

            //if (Pages == 0)  //anhui20120515
            if (Pages == curPages)  //20130311集中打印修改
            {
                if ((xuflag == "1") && (Startrow != 0)) //续打
                {

                }
                else
                {
                    // PrintShape(xmlprndoc["Root"]["SHAPES"], g, ev.PageBounds.Size.Height);
                    PrintShape(xmlprndoc["Root"]["SHAPES"], g, pDoc.DefaultPageSettings.Bounds.Height);


                }
            }

            // if ((Pages < EdP) && (EdP != 0)) HasMorePages = true;
            int POSY = 0;
            string PrevDiaID = "", CurrDiaID = "";
            string PrevLocID = "", CurrLocID = "";
            string PrevBed = "", CurrBed = "";     //打印转床信息   2014.10.23
           
            string Prevpageflag = "", Currpageflag = "";
            if (xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"] != null)
            {
                headflag = true;
            }
            string cellDatanod = "";  //数据节点
            string xHn = "";   //表头节点
            foreach (XmlNode xs in xmlprndoc["Root"]["MetaData"].ChildNodes)
            {
                if (xs.Attributes["Rel"] != null)
                {
                    string[] a = xs.Attributes["Rel"].Value.Split('.');
                    if (a[1].Substring(0, 1) == "T") cellDatanod = xs.Name;
                }
            }
            foreach (XmlNode xs in xmlprndoc["Root"]["InstanceData"].ChildNodes)
            {
                if (xs.Attributes["text"] != null)
                {
                    string a = xs.Attributes["text"].Value;
                    if (a == "") continue;
                    if (a.Substring(0, 1) == "T")
                    {
                        if (xs.Name.Substring(0, 1) == "B") //续打
                        {
                            xHn = xs.Name;
                        }
                    }
                }
            }
            string TableText = "";

            if ((xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"] != null) && (stPrintPos == 0))
            {

                if (xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"].Value != "")
                {
                    if (xmlprndoc["Root"]["THEAD"] != null)
                    {

                        POSY = drawHead(g, xmlprndoc["Root"]["THEAD"]);

                    }
                    else
                    {
                        if (xHn != "")
                        {  ////anhui20120515
                            TableText = xmlprndoc["Root"]["InstanceData"][xHn].Attributes["text"].Value;


                            POSY = PrintHead(xmlprndoc["Root"]["InstanceData"][xHn], g, xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"].Value);



                        }
                        else
                        {

                            POSY = PrintHead(xmlprndoc["Root"]["InstanceData"].ChildNodes[0], g, xmlprndoc["Root"]["SHAPES"].Attributes["HeadPrn"].Value);

                        }
                    }
                    headflag = true;
                }

            }
            if (headflag == false)
            {
                if ((xuflag == "1") && (Startrow != 0) && (Startpage == Pages)) //续打
                {

                }
                else
                {

                    //PrintShape(xmlprndoc["Root"]["SHAPES"], g, ev.PageBounds.Size.Height);
                    PrintShape(xmlprndoc["Root"]["SHAPES"], g, pDoc.DefaultPageSettings.Bounds.Height);


                }
            }
            if ((xmlprndoc["Root"]["PageHeadFoot"] != null) && (stPrintPos == 0))
            {
                //int hf = 0;
                /// int nodes = 0;
                foreach (XmlNode xn in xmlprndoc["Root"]["PageHeadFoot"].ChildNodes)
                {
                    if (xn.Name.IndexOf("HFNOD", 0) == -1)
                    {
                        if (xmlprndoc["Root"]["InstanceData"][xn.Name] != null)
                        {
                            // nodes = nodes + 1;
                            //if ((nodes != 1) && (xn.Name.IndexOf("PageNo") == -1))
                            // {
                            //  if (HFCaption[hf]!=null) xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value = HFCaption[hf];
                            //   hf = hf + 1;
                            // }
                            string aa = xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value;
                            if (PageLocNod != null)
                            {
                                if ((xn.Name == PageLocNod.Name) && (aa != "")) //续打
                                {
                                    try
                                    {
                                        int locid = Convert.ToInt32(aa);
                                        aa = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetNurloc", "par:" + locid + "^");
                                    }
                                    catch
                                    {

                                    }

                                }
                            }
                            xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value = aa;
                            if (HFCaption.ContainsKey(aa))
                            {
                                xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value = HFCaption[aa].ToString();

                            }
                            if (HLasthead.Contains(Pages))  //沈阳医大诊断换页
                            {
                                if (aa.IndexOf("诊断") > -1)
                                {

                                    xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value = HLasthead[Pages].ToString();

                                }

                            }
                            if ((xuflag == "1") && (Startrow != 0) && (Startpage == Pages)) //续打
                            {

                            }
                            else
                            {

                                DrawTxt(xmlprndoc["Root"]["InstanceData"][xn.Name], g);

                            }


                        }
                    }
                }
            }
            if (headflag == true)  //anhui20120515
            {
                //if (Pages == 0)
                if (Pages == curPages)  //20130311集中打印修改
                {
                    if ((MthArr != "") && (Parrm != "") && (curPages > 0))  //hunhemoban
                    {

                    }
                    else
                    {
                        foreach (XmlNode xnod in xmlprndoc["Root"]["InstanceData"].ChildNodes)
                        {
                            if (xnod.Attributes["text"].Value.IndexOf(TableText) == -1)
                            {
                                string valueflag = xnod.Attributes["text"].Value;
                                if ((valueflag != "LHead") && (valueflag != "RHead") && (valueflag != "LFoot") && (valueflag != "RFoot"))
                                {
                                    if ((xuflag == "1") && (Startrow != 0) && (Startpage == Pages)) //续打
                                    {

                                    }
                                    else
                                    {

                                        DrawTxt1(xnod, g);

                                    }
                                }
                            }
                        }
                    }
                }
            }
            //MessageBox.Show(stPrintPos.ToString());
            XmlNode cellnod;
            if (cellDatanod != "") cellnod = xmlprndoc["Root"]["InstanceData"][cellDatanod];
            else cellnod = xmlprndoc["Root"]["InstanceData"][xHn];
            float fontnum = float.Parse(cellnod.Attributes["fontsize"].Value); //字体尺寸
            float RowH = float.Parse(cellnod.Attributes["height"].Value); //行高
            int y1 = POSY;
            int x1 = tabx;
            //if (stRow != 0) y1 = stPrintPos - (int)(stRow * RowH);
            if ((Startrow != 0) && (Pages == Startpage)) y1 = y1 + (int)(Startrow * RowH); //续打

            int PrnH = (pDoc.DefaultPageSettings.PaperSize.Height - pDoc.DefaultPageSettings.Margins.Bottom - y1);  //
            if (pDoc.DefaultPageSettings.Landscape == true)  //横向打印
            {
                PrnH = (pDoc.DefaultPageSettings.PaperSize.Width - pDoc.DefaultPageSettings.Margins.Bottom - y1);  //
            }

            //int PrnH = (ev.MarginBounds.Y + ev.MarginBounds.Height - y1);  //

            ItmCount = 0;
            if ((Row == 0) && (stPrintPos != 0)) y1 = stPrintPos;
            for (int i = 0; i < table.Columns.Count; i++)
            {
                if (head == null) continue;
                if (head[i] == "") continue;
                string[] th = head[i].Split('^');
                string[] afth = null;
                if (i < (table.Columns.Count - 1)) afth = head[i + 1].Split('^');
                int tx = 0, tw = 0, afx = 0;
                if (afth != null) afx = int.Parse(afth[1]);
                else
                {
                    afx = tabW;
                }
                tx = int.Parse(th[1]); tw = afx - x1;
                head[i] = head[i] + "^" + tw.ToString();
                x1 = x1 + tw;
            }

            int singcount = 0;
            if ((Startrow != 0) && (Pages == Startpage)) singcount = Startrow; //续打
            bool setflag = false;
            Hashtable rowlinedata = new Hashtable();
            int PH = pDoc.DefaultPageSettings.PaperSize.Height - pDoc.DefaultPageSettings.Margins.Bottom;//页面可利用高度
            //int PH = ev.PageBounds.Size.Height - pDoc.DefaultPageSettings.Margins.Bottom;//页面可利用高度
            PageRows = PrnH / (int)RowH; ///

            // pDoc.DefaultPageSettings.Bounds.Size.Height

            if ((Startrow != 0) && (Pages == Startpage))  //
            {
                PageRows = (pDoc.DefaultPageSettings.PaperSize.Height - POSY - pDoc.DefaultPageSettings.Margins.Bottom) / (int)RowH; //续打
                if (pDoc.DefaultPageSettings.Landscape == true)  //横向打印
                {
                    PageRows = (pDoc.DefaultPageSettings.PaperSize.Width - POSY - pDoc.DefaultPageSettings.Margins.Left) / (int)RowH; //续打
                }
            }
            //if ((Startrow != 0) && (Pages == Startpage)) PageRows = (ev.MarginBounds.Y + ev.MarginBounds.Height - POSY) / (int)RowH; //续打
            printPagesize = PageRows;
            if (stPrintPos == 0)
            {
                if ((xuflag == "1") && (Startrow != 0) && (Startpage == Pages)) //续打
                {

                }
                else
                {

                    if (AllLine == "Y") //续打
                    {
                        DrawLines(tabx, y1, tabW - tabx, PrnH, g, (int)RowH, head);
                    }
                    //   DrawLines(tabx, y1, tabW - tabx, PrnH, g, (int)RowH, head);

                }
            }
            if (dxflag == 0)
            {
                if ((xuflag == "1") && (Startrow != 0) && (Startpage == Pages)) //续打
                {

                }
                else
                {
                    //不管多少行记录都画线
                    for (int k = 0; k < PageRows; k++)
                    {
                        int hh = (int)RowH;

                        g.DrawLine(new Pen(Color.Black, 1), new Point(tabx, y1 + (k + 1) * hh), new Point(tabW, y1 + (k + 1) * hh));

                    }
                }
            }
            int flagxh = 0;//记录循环次数 转科
            int onum = 0; //双行个数
            do
            {
                x1 = tabx;
                int hh = 0;
                setflag = false;
                if (Row == 8)
                {
                    int retrowstest = 0;
                    if (tcoldata.Count > 0)
                    {
                        //setCell(table.Rows[Row], ev.Graphics, head,tcoldata );
                        //tcoldata.Clear();
                        //setflag = true;
                    }
                }
                if (table.Rows.Count == 0) break;
                if (Row >= table.Rows.Count) break;

                int htt = 0;
                float RH; //
                int retrows = 0;
                if (tcoldata.Count > 0)
                {
                    //setCell(table.Rows[Row], ev.Graphics, head,tcoldata );
                    //tcoldata.Clear();
                    setflag = true;
                }

                if (setflag == true)
                { // 分页数据
                    rowlinedata = tcoldata;
                    retrows = getrows(tcoldata, head);
                    htt = retrows * (int)RowH;

                }
                else
                {
                    //if (Row == 52)
                    //{
                    ////   rowlinedata = getRowH(table.Rows[Row], ev.Graphics, head, fontnum, RowH, out RH, out retrows, cellnod);
                    //   htt = (int)RH;  //每条总行高
                    //}
                    //else
                    //{

                    rowlinedata = getRowH(table.Rows[Row], g, head, fontnum, RowH, out RH, out retrows, cellnod);

                    htt = (int)RH;  //每条总行高
                    //}

                }

                //if (singcount ==PageRows )
                //{
                //    HasMorePages = true;
                //    break;
                //}
                //某条记录超过一页，且刚好打到最后一行
                if ((retrows == PageRows) & (tcoldata.Count != 0))
                {
                    Hashtable tcoldataback = new Hashtable();
                    tcoldataback = setbackhash(tcoldata);
                    rowlinedata = tcoldataback;
                    tcoldata.Clear();
                    // tcoldata.Clear();
                }
                bool rowflag = false;
                if (((singcount + retrows) > PageRows))
                {
                    retrows = PageRows - singcount;
                    htt = retrows * (int)RowH;
                    rowlinedata = splitRow1(rowlinedata, retrows, head);
                }
                else
                {
                    if (retrows == 1)
                    {   //数据是一行并且 余下的行数够加高一行 则单条数据行高为两行高度，否则
                        if ((((singcount + retrows + 1) <= PageRows)) && (dxflag == 1))
                        {
                            retrows = 2;
                            htt = (int)(RowH * 2);
                            rowflag = true;
                        }
                    }
                }

                bool lineSum = false; //
                bool lineNod = false;
                int stx = 0, sty = 0; //////////////
                string CareDate = "", CareTime = "";
                if (table.Columns.Contains("DiagnosDr"))
                {
                    if (table.Rows[Row]["DiagnosDr"] != null) CurrDiaID = table.Rows[Row]["DiagnosDr"].ToString();
                }
                //if ((((PrevDiaID != "") && (CurrDiaID != "")) && (CurrDiaID != PrevDiaID))&&(((PrevLocID != "") && (CurrLocID != "")) && (CurrLocID != PrevLocID)))  //同一科室诊断变了不换页
                if (((PrevDiaID != "") && (CurrDiaID != "")) && (CurrDiaID != PrevDiaID) && (tranflag == "1"))
                {  //

                    HasMorePages = true;
                    tcolbakdata.Clear();
                    tcoldata.Clear();
                    break;

                }
                if (NurseLocHuanYe == "Y") //按护士科室换页
                {
                    if (table.Columns.Contains("RecNurseLoc")) //转科
                    {
                        if (table.Rows[Row]["RecNurseLoc"] != null) CurrLocID = table.Rows[Row]["RecNurseLoc"].ToString();
                    }
                }
                else
                {
                    if (table.Columns.Contains("RecLoc")) //转科
                    {
                        if (table.Rows[Row]["RecLoc"] != null) CurrLocID = table.Rows[Row]["RecLoc"].ToString();
                    }

                }
                if (table.Columns.Contains("NextPageFlag")) //换页标志
                {
                    Currpageflag = table.Rows[Row]["NextPageFlag"].ToString();
                    if ((Currpageflag == "Y") & (!nextpagehastable.Contains(Row))) //转科
                    {  //
                        if ((MakeTemp == "Y") && (Row == 0)) //生成图片时第一条记录有换页标志不换页
                        { }
                        else
                        {
                            HasMorePages = true;
                            tcolbakdata.Clear();
                            tcoldata.Clear();
                            break;
                        }
                    }
                }
                if ((PrevLocID == "") && (flagxh != 0))
                {
                    string parr = EpisodeID + "!" + EmrCode;
                    try
                    {
                        // PrevLocID = PrnLoc; //转科
                        //  PrevLocID = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetFirstNulRecloc", "par:" + parr + "^");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("web.DHCNurRecPrint:GetFirstNulRecloc是否存在？params:" + parr + ex.Message);
                        //return;
                    }

                }
                if (((PrevLocID != "") && (CurrLocID != "")) && (CurrLocID != PrevLocID) && (tranflag == "1")) //转科
                {  //
                    if (ShowLocTran == "Y")
                    {    //转科不分页打印时科室处显示转科信息 2014.10.23
                        string oldtextloc = PageLocNod.Attributes["text"].Value;
                        string curlocdesc = "";
                        try
                        {
                            curlocdesc = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetNurloc", "par:" + CurrLocID + "^");
                        }
                        catch
                        { }
                        PageLocNod.Attributes["text"].Value = oldtextloc + "->" + curlocdesc;
                        DrawTxt(PageLocNod, g);
                    }
                    if (SplitPage == "1")
                    {
                        HasMorePages = true;
                        tcolbakdata.Clear();
                        tcoldata.Clear();
                        break;
                    }

                }

                if (table.Columns.Contains("RecBed"))       //转床打印床号 2014.10.23
                {
                    if (table.Rows[Row]["RecBed"] != null) CurrBed = table.Rows[Row]["RecBed"].ToString();
                }
                if (((PrevBed != "") && (CurrBed != "")) && (PrevBed != CurrBed)) //转床打印床号 2014.10.23
                {  //
                    //if (SplitPage == "1")
                    //{
                    if (PageBedNod != null)  //如果未启用转科
                    {
                        string oldtext = PageBedNod.Attributes["text"].Value;
                        PageBedNod.Attributes["text"].Value = oldtext + "->" + CurrBed;
                        DrawTxt(PageBedNod, g);
                    }
                    //break;
                    //}
                }

                if (table.Columns.Contains("HeadDR"))       //表头变化模板，变换表头后换页 --20141201
                {
                    if (table.Rows[Row]["HeadDR"] != null) CurrHeadDR = table.Rows[Row]["HeadDR"].ToString();
                    if ((Pages == curPages) && (Row == 0))
                    {
                        PrevHeadDR = table.Rows[Row]["HeadDR"].ToString();
                    }
                    if (Pages == 2)
                    {

                        CurrHeadDR = table.Rows[Row]["HeadDR"].ToString();
                    }
                }
               // if ((CurrHeadDR != "")) && (PrevHeadDR != CurrHeadDR)&&(Row!=0)&&(singcount!=0)) //表头变化模板，变换表头后换页 --20141201

                if ((PrevHeadDR != CurrHeadDR) && (singcount != 0)) //表头变化模板，变换表头后换页,singcount=0且前后表头id不同说明前一个表头的最后一页刚好打满换页，此处无需再换页 --20141201
                {
                    HasMorePages = true;
                    tcolbakdata.Clear();
                    tcoldata.Clear();
                    PrevHeadDR = CurrHeadDR;
                    curhead = CurrHeadDR;
                    break;
                }
                ///备份数据
                SetTable(tableBak, rowlinedata, head, retrows, rowflag);
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (head[i] == "") continue;
                    string[] th = head[i].Split('^');
                    string[] afth = null;
                    if (i < (table.Columns.Count - 1)) afth = head[i + 1].Split('^');
                    int tx = 0, tw = 0, afx = 0;
                    if (afth != null) afx = int.Parse(afth[1]);
                    else
                    {
                        afx = tabW;
                    }
                    tx = int.Parse(th[1]); tw = int.Parse(th[4]);
                    if (tw < 5) tw = 0; //列宽
                    // Graphics g = ev.Graphics;
                    if (table.Rows[Row][th[0]].ToString().IndexOf(BlueString) != -1)
                    {
                        lineSum = true;
                    }
                    if (table.Rows[Row][th[0]].ToString().IndexOf("入液量=") != -1) lineNod = true;
                    //if (y1 +htt)
                    //if (tw != 0) DrawString(table.Rows[Row][th[0]].ToString(), x1, y1, tw, htt, ev.Graphics);
                    if (rowlinedata.Count == 0)
                    {

                        ArrayList array22 = (ArrayList)rowlinedata[th[0]];
                        continue;

                    }
                    ArrayList array = (ArrayList)rowlinedata[th[0]];

                    if (tw != 0)
                    {
                        XmlNode xalign = null;
                        if (xmlprndoc["Root"]["THEAD"] != null)
                        {
                            xalign = xmlprndoc["Root"]["THEAD"][th[0]];
                        }
                        if (th[0] == "User")
                        {
                            if (IsVerifyCALoc == "1")
                            {
                                float yimage = y1;
                                for (int i3 = 0; i3 < array.Count; i3++)
                                {
                                    string userstr = array[i3].ToString();

                                    if (userstr != " ")
                                    {
                                        string[] useridstr = userstr.Split(' ');
                                        for (int i2 = 0; i2 < useridstr.Length; i2++)
                                        {
                                            string uid = useridstr[i2];
                                            string imageuser = null;
                                            try
                                            {
                                              imageuser = Comm.DocServComm.GetData("web.DHCNurSignVerify:GetUserSignImage", "par:" + uid + "^");
                                            }
                                            catch(Exception ex){
                                            
                                            }
                                            if (imageuser != null)
                                            {
                                                Comm cmg = new Comm();
                                                Image img = cmg.StringToImage(imageuser);

                                                int addh = 0;
                                                int addw = 0;
                                                if (qmprnorientation == 1)
                                                {
                                                    addw = qmwildth + qmhori;
                                                }
                                                if (qmprnorientation == 0)
                                                {
                                                    addh = qmheight + qmport;
                                                }
                                                if (blackflag == "Y") //红色转黑色
                                                {
                                                    Bitmap bmp = changecolor(img, 4);

                                                    g.DrawImage(bmp, x1 + qmleft + addw * i2, y1 + qmtop + addh * i2, qmwildth, qmheight);

                                                }
                                                else
                                                {
                                                    //g.DrawImage(img, x1 + 2, yimage + 5 + i2 * 10, qmwildth, qmheight);

                                                    g.DrawImage(img, x1 + qmleft + addw * i2, yimage + qmtop + addh * i2, qmwildth, qmheight);

                                                    //g.DrawImage(img, x + addw * i, y + i * addh, qmwildth, qmheight);
                                                }
                                            }
                                            else
                                            {

                                                DrawString1(array, x1, y1, tw, htt, g, (int)RowH, cellnod, rowflag, xalign);


                                            }


                                        }
                                    }
                                    yimage = yimage + RowH;
                                }
                                //DrawString1(array, x1, y1, tw, htt, ev.Graphics, (int)RowH, cellnod, rowflag, xalign);
                            }
                            else
                            {

                                DrawString1(array, x1, y1, tw, htt, g, (int)RowH, cellnod, rowflag, xalign);


                            }
                        }
                        else
                        {
                            // string PageLoc = Comm.DocServComm.GetData("web.DHCNurSignVerify:GetUserDesc", "par:"  + "^");

                            DrawString1(array, x1, y1, tw, htt, g, (int)RowH, cellnod, rowflag, xalign);

                        }
                        if (rowflag == true)  //记录行高
                        {
                            //tableBak.Rows 
                            BakRowH[tableBak.Rows.Count - 1] = 2;

                        }

                    }

                    if (th[0] == "CareDateTim") CareDateTim = table.Rows[Row][th[0]].ToString();
                    if (th[0] == "CareDate") CareDate = table.Rows[Row][th[0]].ToString();
                    if (th[0] == "CareTime") CareTime = table.Rows[Row][th[0]].ToString();
                    if ((CareDate != "") && (CareTime != ""))
                    {
                        CareDateTim = CareDate + "/" + CareTime;
                    }

                    //if (y1 == tabH)
                    //{
                    //    g.DrawLine(new Pen(Color.Black, 1), new Point(afx, taby), new Point(afx, tabH));
                    //}
                    if (i == 0)
                    {
                        stx = x1; sty = y1;
                    }
                    x1 = x1 + tw;
                    hh = htt;
                }

                string rowid = rowidha[Row].ToString();
                int row = singcount;
                int pagerows = PageRows;// +Startrow;
                if ((xuflag == "1") && (PreView == "0")) //续打
                {
                   
                    if (Row == table.Rows.Count - 1)   //续打201407
                    {
                        //row = row + retrows;
                    }
                    if (printinfo == "")
                    {

                        printinfo = rowid + "^" + Pages + "^" + row + "^" + y1 + "^" + retrows + "^" + RowH + "^" + pagerows;

                    }
                    else
                    {

                        printinfo = printinfo + "&" + rowid + "^" + Pages + "^" + row + "^" + y1 + "^" + retrows + "^" + RowH + "^" + pagerows;

                    }


                }
                if ((MakeTemp == "Y") || (curhead != ""))  //生成图片  表头变更
                {
                    //string rowid = rowidha[Row].ToString();
                    //int row = singcount;
                   // int pagerows = PageRows;// +Startrow;
                    try
                    {
                        if (rowid != "")
                        {
                            curhead = Comm.DocServComm.GetData("NurEmr.webheadchange:getrowhead", "par:" + rowid + "^");
                            if (curhead == null) curhead = "";
                        }
                    }
                    catch(Exception ex){
                        MessageBox.Show("NurEmr.webheadchange:getrowhead是否存在？");
                    }
                    if (printinfo == "") //第一条信息 每页包含的记录id及页码信息
                    {
                        int spagemake = Pages + 1;
                        int srowmake = row + 1;
                        printinfo = spagemake + "^" + rowid + "^" + srowmake + "^" + curhead + "^" + rowid;
                    }
                    else
                    {
                        printinfo = printinfo + "*" + rowid;
                    }
                    int strow = singcount;     //开始行
                    int edrow = singcount + retrows; //结束行
                    int rowh = retrows; //一条记录占用几行
                    int rownewpage = Pages + 1; //yema
                    if (edrow > pagerows)
                    {
                        edrow = pagerows;
                        rowh = pagerows - strow;
                    }
                    if (rowprintinfo == "")
                    {
                        rowprintinfo = rowid + "^" + rownewpage + "^" + strow + "^" + edrow + "^" + rowh + "^" + y1 + "^" + pagerows;

                    }
                    else
                    {
                        rowprintinfo = rowprintinfo + "&" + rowid + "^" + rownewpage + "^" + strow + "^" + edrow + "^" + rowh + "^" + y1 + "^" + pagerows;

                    }

                }
                if (AllLine != "Y") //按内容高度打印竖线
                {
                    DrawLines(tabx, y1, tabW - tabx, (int)RowH * retrows, g, (int)RowH, head);
                }
                singcount = singcount + retrows; //单页行数
                if (rowflag == true) onum = onum + 1;
                //PageRows *RowH
                if ((sty + htt) < (PH - RowH + 5))
                {
                    if (dxflag == 0)
                    {
                        if ((xuflag == "1") && (Startrow != 0) && (Startpage == Pages)) //续打
                        {

                        }
                        else
                        {


                            g.DrawLine(new Pen(Color.Black, 1), new Point(stx, sty + htt), new Point(x1, sty + htt));

                        }
                    }
                    else
                    {

                        g.DrawLine(new Pen(Color.Black, 1), new Point(stx, sty + htt), new Point(x1, sty + htt));

                    }
                }
                else
                {
                    //   MessageBox.Show ("dd");
                }
                if (lineSum == true)
                {

                    g.DrawLine(new Pen(Color.Blue, 2), new Point(stx, sty), new Point(x1, sty));
                    RecLine(Pages + 1, new Point(stx, sty), new Point(x1, sty));
                    g.DrawLine(new Pen(Color.Blue, 2), new Point(stx, sty + htt), new Point(x1, sty + htt));
                    RecLine(Pages + 1, new Point(stx, sty + htt), new Point(x1, sty + htt));


                }
                if (lineNod == true)
                {

                    g.DrawLine(new Pen(Color.Blue, 2), new Point(stx, sty + htt), new Point(x1, sty + htt));
                    RecLine(Pages + 1, new Point(stx, sty + htt), new Point(x1, sty + htt));



                }
                if (table.Columns.Contains("DiagnosDr"))
                {
                    if (table.Rows[Row]["DiagnosDr"] != null) PrevDiaID = table.Rows[Row]["DiagnosDr"].ToString();
                }
                if (NurseLocHuanYe == "Y") //按护士科室换页
                {
                    if (table.Columns.Contains("RecNurseLoc"))
                    {
                        if (table.Rows[Row]["RecNurseLoc"] != null) PrevLocID = table.Rows[Row]["RecNurseLoc"].ToString();
                    }
                }
                else
                {
                    if (table.Columns.Contains("RecLoc"))
                    {
                        if (table.Rows[Row]["RecLoc"] != null) PrevLocID = table.Rows[Row]["RecLoc"].ToString();
                    }

                }
                if (table.Columns.Contains("NextPageFlag"))
                {
                    if (table.Rows[Row]["NextPageFlag"] != null) Prevpageflag = table.Rows[Row]["NextPageFlag"].ToString();
                }
                if (table.Columns.Contains("RecBed"))  //转床打印转床信息 2014.10.23
                {
                    if (table.Rows[Row]["RecBed"] != null) PrevBed = table.Rows[Row]["RecBed"].ToString();
                }

                if (table.Columns.Contains("HeadDR")) //表头变换换页
                {
                    if (table.Rows[Row]["HeadDR"] != null) PrevHeadDR = table.Rows[Row]["HeadDR"].ToString();
                }

                y1 = y1 + hh;

                if (setflag == true)
                {
                    if (retrows == PageRows)  //一条记录超过二页
                    {
                        // setCell(table.Rows[Row], ev.Graphics, head, tcolbakdata); //复原数据
                    }
                    else
                    {

                        setCell(table.Rows[Row], g, head, tcolbakdata); //复原数据

                        tcolbakdata.Clear();
                        tcoldata.Clear();
                    }
                }
                if (singcount == PageRows)
                {
                    if (table.Columns.Contains("HeadDR")) //表头变换换页
                    {
                        if ((Row < table.Rows.Count - 1)&&(tcoldata.Count == 0))
                        {
                            string nextHeadDR = "";
                            if (table.Rows[Row + 1]["HeadDR"] != null) nextHeadDR = table.Rows[Row + 1]["HeadDR"].ToString();
                            if (PrevHeadDR != nextHeadDR) //表头变化模板，如果下条记录表头id有变化，更换当前表头id(换页后打印表头用)
                            {
                                curhead = nextHeadDR;
                            }
                        }
                    }
                  
                    HasMorePages = true;
                    if (tcoldata.Count == 0) Row++;
                    if (Row >= (table.Rows.Count)) HasMorePages = false;
                    break;

                }
                //if (y1 > (ev.PageBounds.Size.Height - pDoc.DefaultPageSettings.Margins.Bottom))
                //{
                //    //Row++;
                //    //if (Row >= table.Rows.Count) HasMorePages = false;
                //    //else HasMorePages = true;
                //    //Row++;
                //    HasMorePages = true;
                //    break;
                //}
                Row++;
                flagxh++; //转科


            } while ((Row < table.Rows.Count) && (y1 < (pDoc.DefaultPageSettings.Bounds.Size.Height - pDoc.DefaultPageSettings.Margins.Bottom)));
            // } while ((Row < table.Rows.Count) && (y1 < (ev.PageBounds.Size.Height - pDoc.DefaultPageSettings.Margins.Bottom)));

            if (HasMorePages == false)
            {


                if ((PreView == "0") && (xuflag == "1")) //续打201407
                {
                    try
                    {
                        Comm.DocServComm.Save(EpisodeID, ItmName, printinfo);
                        printinfo = "";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                Row = 0;
                string parm = EpisodeID + "^" + CareDateTim + "^" + y1.ToString() + "^" + Pages.ToString();
                if (HPagRow.Contains(Pages + 1) == false)
                {
                    HPagRow.Add(Pages + 1, singcount);
                    HPagRow1.Add(Pages + 1, singcount - onum);
                }
                printpagecount = Pages + 1;// -stPage + 1;

                // stPrintPos =y1;



                if (PreView == "1")
                {
                    if (MakeTemp != "Y")
                    {
                        Pages = Startpage; //续打
                    }
                    // y1 = stPrintPos;
                }
                else
                {
                    // stPage = Pages;
                    // stPrintPos = y1;
                    // stRow = singcount;
                    if (MakeTemp != "Y")
                    {
                        clearxuprint();//续打
                    }
                    PrnPreView.Dispose();
                }
                //首次是预览 将PreView=0，第二次是打印
                //if (PreView == "1") PreView = "0";
                // if (PreView == "0") PreView = "1";
                PreView = "0";
                PrnFlag++;

                // HFCaption["LHead"] = ChangePageDiag ;
                // if (ChangePageDiag == "") HFCaption["LHead"] = LHeadCaption;
                if (HLhead.Contains(printpagecount) == false) HLhead.Add(printpagecount, HFCaption["LHead"]);
                if (PageDiagNod != null)
                {  //201110116
                    if (PageDiagH.Contains(printpagecount) == false) PageDiagH.Add(printpagecount, PageDiagNod.Attributes["text"].Value);
                }
                if (PageLocNod != null)
                {  //201110116
                    if (PageLocH.Contains(printpagecount) == false) PageLocH.Add(printpagecount, PageLocNod.Attributes["text"].Value);
                    //PageLocNod.Attributes["text"].Value = table.Rows[Row]["RecLoc"].ToString();
                    //DrawTxt(PageLocNod, ev.Graphics);
                }
                if (PageBedNod != null)
                {  //201110116
                    if (PageBedH.Contains(printpagecount) == false) PageBedH.Add(printpagecount, PageBedNod.Attributes["text"].Value);
                }

                // MessageBox.Show(PrnFlag.ToString() ); 
                // HFCaption["LHead"] = HLhead[1].ToString();
            }
            else
            {

                if ((PreView == "0") && (xuflag == "1")) //续打201407
                {
                    try
                    {
                        Comm.DocServComm.Save(EpisodeID, ItmName, printinfo);
                        printinfo = "";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                stPrintPos = 0;
                stRow = 0;
                //换页取诊断
                Pages++;
                if (NurseLocHuanYe == "Y")
                {
                    if (table.Columns.Contains("RecNurseLoc"))
                    {
                        string recloc1 = table.Rows[Row]["RecNurseLoc"].ToString();

                        if (recloc1 == "")
                        {
                            string parr1 = EpisodeID + "!" + EmrCode;
                            try
                            {
                                recloc1 = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetFirstNulRecNurseloc", "par:" + parr1 + "^");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("web.DHCNurRecPrint:GetFirstNulRecNurseloc是否存在？params:" + parr1 + ex.Message);
                                return;
                            }
                        }
                        if (recloc1 == "") recloc1 = PrnLoc;
                        if (PageLocNod != null)
                        {  //转科
                            string PageLoc1 = "";
                            try
                            {
                                PageLoc1 = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetNurloc", "par:" + recloc1 + "^");
                            }
                            catch
                            { }

                            if (PageLocH.Contains(Pages) == false) PageLocH.Add(Pages, PageLocNod.Attributes["text"].Value);
                            PageLocNod.Attributes["text"].Value = PageLoc1;
                        }
                    }
                }
                else
                {
                    if (table.Columns.Contains("RecLoc"))
                    {
                        string recloc = table.Rows[Row]["RecLoc"].ToString();

                        if (recloc == "")
                        {
                            string parr = EpisodeID + "!" + EmrCode;
                            try
                            {
                                recloc = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetFirstNulRecloc", "par:" + parr + "^");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("web.DHCNurRecPrint:GetFirstNulRecloc是否存在？params:" + parr + ex.Message);
                                return;
                            }
                        }
                        if (recloc == "") recloc = Patcurloc;
                        if (PageLocNod != null)
                        {  //转科
                            string PageLoc = "";
                            try
                            {
                                PageLoc = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetNurloc", "par:" + recloc + "^");
                            }
                            catch
                            { }

                            if (PageLocH.Contains(Pages) == false) PageLocH.Add(Pages, PageLocNod.Attributes["text"].Value);
                            PageLocNod.Attributes["text"].Value = PageLoc;
                        }
                    }
                }
                if (table.Columns.Contains("RecBed"))
                {
                    string recbed = table.Rows[Row]["RecBed"].ToString();
                    if (recbed == "") recbed = PrnBed;
                    if (PageBedNod != null)
                    {  //转科
                        if (PageBedH.Contains(Pages) == false) PageBedH.Add(Pages, PageBedNod.Attributes["text"].Value);
                        PageBedNod.Attributes["text"].Value = recbed;
                    }

                }
                if (table.Columns.Contains("NextPageFlag"))
                {
                    if (table.Rows[Row]["NextPageFlag"] != null) Prevpageflag = table.Rows[Row]["NextPageFlag"].ToString();
                    if (Prevpageflag == "Y")
                    {
                        if (nextpagehastable.Contains(Row) == false)
                            nextpagehastable.Add(Row, "Y");
                    }
                }

                if (table.Columns.Contains("DiagnosDr"))
                {
                    //PAN
                    string PageDiag = Comm.DocServComm.GetData("Nur.DHCNurCopyDiagnos:GetNurDiagnos", "par:" + table.Rows[Row]["DiagnosDr"].ToString() + "^");
                    // string PageLoc = Comm.DocServComm.GetData("web.DHCNurRecPrint:GetNurloc", "par:" + table.Rows[Row]["RecLoc"].ToString() + "^");
                    if (PageDiagNod != null)
                    {  //201110116
                        if (PageDiagH.Contains(Pages) == false) PageDiagH.Add(Pages, PageDiagNod.Attributes["text"].Value);
                        PageDiagNod.Attributes["text"].Value = PageDiag;
                    }
                    if (LHeadCaption != "")
                    {
                        if (HLhead.Contains(Pages) == false) HLhead.Add(Pages, HFCaption["LHead"]);
                        if (HLasthead.Contains(Pages - 1) == false) HLasthead.Add(Pages - 1, HFCaption["LHead"]);  //沈阳医大诊断换页
                        //if (HLhead.Contains(Pages) == false) HLhead.Add(Pages, HFCaption["LHead"]);
                        int sindex = LHeadCaption.IndexOf("诊断:");
                        if (sindex > -1)
                        {
                            string oldcaption = LHeadCaption;
                            LHeadCaption = LHeadCaption.Substring(0, sindex) + "诊断:" + PageDiag;
                            HFCaption["LHead"] = LHeadCaption;
                            if (HLasthead.Contains(Pages) == false) HLasthead.Add(Pages, HFCaption["LHead"]); //沈阳医大诊断换页
                            //if (Pages == 0)
                            if (Pages == curPages)  //20130311集中打印修改
                            {
                                ChangePageDiag = oldcaption;
                            }
                        }

                    }
                }
                // LHeadCaption = "";
                if (HPagRow.Contains(Pages) == false)
                {
                    HPagRow.Add(Pages, singcount);
                    HPagRow1.Add(Pages, singcount - onum);
                }

            }

        }


        public void MakePicture()
        {
           //System.Diagnostics.Debugger.Launch();
            //生成图片
            if (MakeAllPages == "Y")
            {
                curPages = 0;
            }
            StartMakePic = "N";
            PrintOut();

            Bitmap image = null; // new Bitmap((int)(pDoc.DefaultPageSettings.PaperSize.Width), (int)(pDoc.DefaultPageSettings.PaperSize.Height));
            if (pDoc.DefaultPageSettings.Landscape == true)
            {
                image = new Bitmap((int)(pDoc.DefaultPageSettings.PaperSize.Height), (int)(pDoc.DefaultPageSettings.PaperSize.Width));
            }
            else
            {
                image = new Bitmap((int)(pDoc.DefaultPageSettings.PaperSize.Width), (int)(pDoc.DefaultPageSettings.PaperSize.Height));
            }
            Graphics gimage = Graphics.FromImage(image);
            string makeinfo = ""; //生成图片保存每页的记录信息
            string mekerecorderinfo = ""; //生成图片保存每条记录的打印信息
            gimage.Clear(Color.White);
            bool HasMorePages = false;
            int nowpage = 0;
            int savepage = 0;
            if ((NurRecId != "")&&(MthArr==""))
            {
                string curp = Comm.DocServComm.GetNurPage(NurRecId, EpisodeID, EmrCode);
                curPages = Convert.ToInt32(curp);
                Pages = curPages;

            }

            do
            {
                HasMorePages = false;
                if ((MthArr != "")&&(Parrm==""))  //评估单类生成图片
                {
                    makepgdpage(gimage, ref HasMorePages);

                    if (HasMorePages)
                    {
                        Pages++;
                        nowpage++;
                        savepage = Pages;
                    }
                    else
                    {
                        nowpage++;
                        if (NurRecId != "")
                        {
                            Comm.DocServComm.SaveNurPage(NurRecId, EpisodeID, EmrCode, Pages + 1 + "", nowpage + "");
                        }
                        printpagecount = Pages + 1;

                        savepage = Pages + 1;
                        ItmCount = 0;
                        PagOffY = 0;
                        PrnCount = 0;
                        Pages = curPages;
                        PgdPrintedArray.Clear();

                    }
                    if (IfUpload == "Y") //是否上传ftp
                    {
                        string thefullname = "c:\\" + (savepage) + ".gif";
                        image.Save(thefullname, System.Drawing.Imaging.ImageFormat.Gif);
                        string pathfold = Comm.DocServComm.GetPictureFilePath(EpisodeID, EmrCode);
                        UploadImage(ftppath, ftpport, ftpuer, ftppwd, ftpdealyTim, thefullname, pathfold);
                        Comm.DocServComm.SavePictureFilePath(EpisodeID, EmrCode, savepage + "");
                        File.Delete(thefullname);
                    }
                    gimage.Clear(Color.White);

                }
                else //记录单类生成图片(包括混合单)
                {
                    makepage(gimage, ref HasMorePages);
                    try
                    {
                        if (makeinfo == "") makeinfo = printinfo;
                        else makeinfo = makeinfo + "&" + printinfo;
                        if (mekerecorderinfo == "") mekerecorderinfo = rowprintinfo;
                        else mekerecorderinfo = mekerecorderinfo + "&" + rowprintinfo;
                        if (HasMorePages == true)
                        {
                            nowpage = Pages;                         
                        }
                        else
                        {
                            //if (makeinfo == "") makeinfo = printinfo;
                            //else makeinfo = makeinfo + "&" + printinfo;
                            //MessageBox.Show(makeinfo);
                            if (makeinfo == "") continue;
                            try
                            {
                                Comm.DocServComm.MakePictureHistory(EpisodeID, EmrCode, makeinfo); //生成图片保存每页记录信息
                                Comm.DocServComm.MakePictureRecorder(EpisodeID, EmrCode, mekerecorderinfo); //生成图片保存每条记录打印信息
                            }catch(Exception ex){
                            
                            }
                            nowpage = Pages + 1;
                            MakeTemp = "N";
                            MakeAllPages = "";
                            printinfo = "";
                            clearxuprint();//续打
                            makeinfo = "";
                            printinfo = "";
                            rowprintinfo = "";
                            Pages = 0;
                            PrevHeadDR = "";
                            CurrHeadDR = "";
                            curhead = "";
                           

                        }
                        if (IfUpload == "Y") //是否上传ftp
                        {
                            string thefullname = "c:\\" + (nowpage) + ".gif";
                            image.Save(thefullname, System.Drawing.Imaging.ImageFormat.Gif);
                            string pathfold = Comm.DocServComm.GetPictureFilePath(EpisodeID, EmrCode);
                            UploadImage(ftppath, ftpport, ftpuer, ftppwd, ftpdealyTim, thefullname, pathfold);
                            Comm.DocServComm.SavePictureFilePath(EpisodeID, EmrCode, nowpage + "");
                            File.Delete(thefullname);
                        }
                        gimage.Clear(Color.White);
                    }
                    catch (Exception ex)
                    {
                        MakeTemp = "N";
                        MakeAllPages = "";
                        MessageBox.Show(ex.Message);
                    }
                }



            }
            while (HasMorePages == true);



        }
        private void pd_PrintPage1(object sender, PrintPageEventArgs ev)
        {
            //int rowcount = 0;
            //int rowsperpage = 0;
            //if (MakeTemp == "Y") return;
            if (EdP > 0)
            {
                // MessageBox.Show(EdP.ToString());
                pd_PrintPage11(sender, ev);
                return;
            }
            bool HasMorePages = false;
            makepage(ev.Graphics, ref HasMorePages);

            if (HasMorePages == false)
            {
                if ((Pages == EdP) && (EdP != 0))
                {

                    HasMorePages = false;
                    Pages = 0;
                    ev.HasMorePages = HasMorePages;
                    return;
                }
            }
            /*
            if (MakeTemp == "Y")
            {
                try
                {
                    image.Save(thefullname, System.Drawing.Imaging.ImageFormat.Gif);
                    string pathfold = Comm.DocServComm.GetPictureFilePath(EpisodeID, EmrCode);
                    //UploadImage("10.56.32.6", "21", "admin", "1", "20", thefullname, pathfold);

                    UploadImage(ftppath, ftpport, ftpuer, ftppwd, ftpdealyTim, thefullname, pathfold);
                    Comm.DocServComm.SavePictureFilePath(EpisodeID,EmrCode,Pages+"");
                    File.Delete(thefullname);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
             * */
            ev.HasMorePages = HasMorePages;


        }

        private bool UploadImage(string ip, string port, string userName, string passWord, string delayTime, string localFilePath, string remoteDirectoryPath)
        {
            FTPFactory ff = new FTPFactory(ip, port, userName, passWord, delayTime);

            //登陆
            try
            {
                ff.login();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                ff.close();
                return false;
            }

            try
            {
                ff.setBinaryMode(true);

                //跳转到根目录
                //ff.chdir("/");

                //从根目录跳转到图片存储目录
                string[] arrRemotePath = remoteDirectoryPath.Split('/');
                for (int j = 0; j < arrRemotePath.Length; j++)
                {
                    ff.mkdir(arrRemotePath[j]);
                    ff.chdir(arrRemotePath[j]);
                }

                //上传文件
                ff.upload(localFilePath);


                ff.close();

                // MessageBox.Show("上传文件成功!");
            }
            catch (Exception ex1)
            {
                MessageBox.Show(ex1.Message);
                return false;
            }

            return true;
        }
        private void clearxuprint() //续打
        {
            Startpage = 0;//续打
            Startrow = 0;//续打
            if (MakeTemp != "Y")
            {
                Pages = 0;//续打
            }
            rowidha.Clear();//续打
            lastprninfo = "";//续打
            printinfo = "";//续打
            xuflag = "0";//续打
            curPages = 0;//续打
            AllLine = "Y";//竖线

        }
        private void RecLine(int pages, Point p1, Point p2)
        {
            if (PageLine.Contains(pages))
            {
                PageLine[pages] = PageLine[pages].ToString() + "^" + p1.X.ToString() + "," + p1.Y.ToString() + "|" + p2.X.ToString() + "," + p2.Y.ToString();
            }
            else
            {
                PageLine[pages] = p1.X.ToString() + "," + p1.Y.ToString() + "|" + p2.X.ToString() + "," + p2.Y.ToString();
            }
        }
        private void DrawRecLine(int page, Graphics g)
        {
            if (PageLine.Contains(page))
            {
                string[] tem = PageLine[page].ToString().Split('^');
                for (int i = 0; i < tem.Length; i++)
                {
                    if (tem[i] == "") continue;
                    string[] parr = tem[i].Split('|');
                    string[] p1 = parr[0].Split(',');
                    string[] p2 = parr[1].Split(',');
                    int x1, x2, y1, y2;
                    x1 = int.Parse(p1[0]); y1 = int.Parse(p1[1]);
                    x2 = int.Parse(p2[0]); y2 = int.Parse(p2[1]);
                    g.DrawLine(new Pen(Color.Blue, 2), new Point(x1, y1), new Point(x2, y2));
                }



            }

        }
        private void pd_PrintPage2(object sender, PrintPageEventArgs ev)
        {
            //int rowcount = 0;
            //int rowsperpage = 0;
            bool HasMorePages = false;
            int yh = 0;
            //ev.Graphics.DrawRectangle(new Pen(Color.Black), ev.MarginBounds);
            Rectangle rect = new Rectangle(new Point(0, 0), ev.PageBounds.Size);

            // ev.Graphics.DrawRectangle(new Pen(Color.Red ),rect);
            PrintShape(xmlprndoc["Root"]["SHAPES"], ev.Graphics, ev.PageBounds.Size.Height);
            if (xmlprndoc["Root"]["PageHeadFoot"] != null)
            {
                //int hf = 0;
                // int nodes = 0;
                foreach (XmlNode xn in xmlprndoc["Root"]["PageHeadFoot"].ChildNodes)
                {
                    if (xn.Name.IndexOf("HFNOD", 0) == -1)
                    {
                        if (xmlprndoc["Root"]["InstanceData"][xn.Name] != null)
                        {
                            string aa = xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value;
                            if (HFCaption.ContainsKey(aa))
                            {
                                xmlprndoc["Root"]["InstanceData"][xn.Name].Attributes["text"].Value = HFCaption[aa].ToString();
                            }

                            DrawTxt(xmlprndoc["Root"]["InstanceData"][xn.Name], ev.Graphics);
                        }
                    }
                }
            }
            foreach (XmlNode xnod in xmlprndoc["Root"]["InstanceData"].ChildNodes)
            {
                DrawTxt(xnod, ev.Graphics);
            }



            if (HasMorePages == false) Row = 0;
            else
            {
                Pages++;
            }
            ev.HasMorePages = HasMorePages;


        }
        private void setCell(DataRow drw, Graphics g, string[] head, Hashtable tdata)
        {
            for (int i = 0; i < table.Columns.Count; i++)
            {
                string[] th = head[i].Split('^');
                drw[th[0]] = tdata[th[0]];
            }

        }
        private int getrows(Hashtable aa, string[] head)
        {
            int rows = 0;
            for (int i = 0; i < head.Length; i++)
            {
                if (head[i] == "") continue;
                string[] th = head[i].Split('^');
                if (aa[th[0]] == null) continue;
                ArrayList dd = (ArrayList)aa[th[0]];
                if (dd.Count > rows) rows = dd.Count;

            }
            return rows;
        }
        private Hashtable setbackhash(Hashtable coldata33)
        {
            Hashtable tcoldatanew = new Hashtable();
            foreach (DictionaryEntry objDE in coldata33)
            {
                tcoldatanew.Add(objDE.Key.ToString(), objDE.Value);
                // Console.WriteLine(objDE.Key.ToString());
                //Console.WriteLine(objDE.Value.ToString());
            }
            return tcoldatanew;
        }
        private Hashtable splitRow1(Hashtable coldata, int getrows, string[] head)
        {
            Hashtable newhash = new Hashtable();
            Hashtable tcoldatanew = new Hashtable();
            tcoldatanew = setbackhash(coldata);
            tcoldata.Clear();
            int maxrow = 0;
            string maxitem = "";
            //换页加签名并靠下显示
            ArrayList Userstr1 = new ArrayList(); //日期
            ArrayList Userstr2 = new ArrayList(); //签名
            ArrayList Userstr3 = new ArrayList();
            ArrayList Userstr4 = new ArrayList();//时间
            ArrayList Userstr5 = new ArrayList();
            ArrayList Userstr6 = new ArrayList(); //日期/时间
            string User = "";
            for (int i = 0; i < head.Length; i++)
            {
                if (head[i] == "") continue;
                string[] th = head[i].Split('^');

                // ArrayList arr = (ArrayList)coldata[th[0]];
                ArrayList arr = (ArrayList)tcoldatanew[th[0]];
                if (arr == null) continue;
                ArrayList newarr = new ArrayList();
                ArrayList sarr = new ArrayList();
                //截取行数
                for (int g = 0; g < getrows; g++)
                {

                    if (g >= arr.Count) continue;
                    newarr.Add(arr[g]);
                }

                if (arr.Count >= getrows)
                {
                    for (int g = getrows; g < arr.Count; g++)
                    {
                        sarr.Add(arr[g]);
                    }
                    if ((arr.Count - getrows) > maxrow)
                    {
                        maxrow = arr.Count - getrows;
                        maxitem = th[0];
                    }
                }
                if (th[0] == "User") //换页加签名并靠下显示
                {
                    int ll = arr.Count;
                    for (int d = 0; d < ll; d++)
                    {
                        string str = (string)arr[d];
                        if (str != " ") Userstr2.Add(str);
                    }
                    //User = (string)arr[arr.Count - 1];
                }
                if (th[0] == "CareDate") //换页加签名并靠下显示
                {
                    int ll = arr.Count;
                    for (int d = 0; d < ll; d++)
                    {
                        string str = (string)arr[d];
                        if (str != " ") Userstr1.Add(str);
                    }
                    //User = (string)arr[arr.Count - 1];
                }
                if (th[0] == "CareDateTim") //换页加签名并靠下显示
                {
                    int ll = arr.Count;
                    for (int d = 0; d < ll; d++)
                    {
                        string str = (string)arr[d];
                        if (str != " ") Userstr6.Add(str);
                    }
                    //User = (string)arr[arr.Count - 1];
                }
                if (th[0] == "CareTime") //换页加签名并靠下显示
                {
                    int ll = arr.Count;
                    for (int d = 0; d < ll; d++)
                    {
                        string str = (string)arr[d];
                        if (str != " ") Userstr4.Add(str);
                    }
                    //User = (string)arr[arr.Count - 1];
                }
                newhash.Add(th[0], newarr);
                tcoldata.Add(th[0], sarr);


            }
            //前一页签名
            if (Userstr2.Count <= getrows) //换页加签名并靠下显示
            {
                int lll = getrows - Userstr2.Count;
                for (int i = 0; i < lll; i++)
                {
                    if (UserPrintDown == "Y")
                    {
                        Userstr3.Add(" ");
                    }

                }
                for (int q = 0; q < Userstr2.Count; q++)
                {
                    Userstr3.Add((string)Userstr2[q]);

                }
                if (lll >= 0)
                {
                    if (newhash.Contains("User"))
                    {
                        newhash.Remove("User");
                        newhash.Add("User", Userstr3);

                    }

                }

            }
            else
            {
                if (newhash.Contains("User"))
                {
                    newhash.Remove("User");
                    newhash.Add("User", Userstr2);

                }


            }
            //后一页签名
            if (Userstr2.Count <= maxrow) //换页加签名并靠下显示
            {
                int lll = maxrow - Userstr2.Count;
                for (int i = 0; i < lll; i++)
                {
                    if (UserPrintDown == "Y")
                    {
                        Userstr5.Add(" ");
                    }

                }
                for (int q = 0; q < Userstr2.Count; q++)
                {
                    Userstr5.Add((string)Userstr2[q]);

                }
                if (lll >= 0)
                {
                    if (newhash.Contains("User"))
                    {
                        tcoldata.Remove("User");
                        tcoldata.Add("User", Userstr5);

                    }

                }

            }
            else
            {
                if (newhash.Contains("User"))
                {
                    tcoldata.Remove("User");
                    tcoldata.Add("User", Userstr2);

                }


            }
            if (tcoldata.Contains("User")) //换页加签名并靠下显示
            {
                //tcoldata.Remove("User");            
            }
            // tcoldata.Add("User", Userstr2);
            if (tcoldata.Contains("CareDate")) //换页加签名并靠下显示
            {
                tcoldata.Remove("CareDate");
                tcoldata.Add("CareDate", Userstr1);
            }

            if (tcoldata.Contains("CareTime")) //换页加签名并靠下显示
            {
                tcoldata.Remove("CareTime");
                tcoldata.Add("CareTime", Userstr4);
            }


            if (tcoldata.Contains("CareDateTim")) //日期时间 //换页加签名并靠下显示
            {
                tcoldata.Remove("CareDateTim");
                tcoldata.Add("CareDateTim", Userstr6);
            }


            return newhash; //要打印的数据
        }
        private Hashtable splitRow1back(Hashtable coldata, int getrows, string[] head)
        {
            Hashtable newhash = new Hashtable();
            Hashtable tcoldatanew = new Hashtable();
            tcoldatanew = setbackhash(coldata);
            tcoldata.Clear();
            for (int i = 0; i < head.Length; i++)
            {
                if (head[i] == "") continue;
                string[] th = head[i].Split('^');

                ArrayList arr = (ArrayList)tcoldatanew[th[0]];
                if (arr == null) continue;
                ArrayList newarr = new ArrayList();
                ArrayList sarr = new ArrayList();
                //截取行数
                for (int g = 0; g < getrows; g++)
                {

                    if (g >= arr.Count) continue;
                    newarr.Add(arr[g]);
                }

                if (arr.Count >= getrows)
                {
                    for (int g = getrows; g < arr.Count; g++)
                    {
                        sarr.Add(arr[g]);
                    }
                }
                newhash.Add(th[0], newarr);
                tcoldata.Add(th[0], sarr);

            }
            tcoldatanew.Clear();
            return newhash; //要打印的数据
        }
        private Hashtable splitRow(DataRow drw, Graphics g, string[] head, int HR)
        {

            float hh1 = 0;
            //h=面积/w  求出一个最高的
            //H/RwH==Rows
            //Fontnum= Rows*RFontNum  
            //如果每单元字数>Fontnum  就拆成两部分
            Hashtable HRw = new Hashtable();
            Comm fn = new Comm();
            string CareDate = "", CareTime = "";

            for (int i = 0; i < table.Columns.Count; i++)
            {
                string[] th = head[i].Split('^');
                string coldata = drw[th[0]].ToString();
                //  SizeF sizf = g.MeasureString("宋", new Font("宋体", 11));

                int aa = getstrLen(coldata); //字数
                if (aa > 4) aa = aa / 2;
                int fh = 16; // (int )sizf.Height;
                int fw = 14; // (int)sizf.Width;
                int linfontnum = int.Parse(th[4]) / fw; //每行的字数  //双字节字数'
                if (linfontnum == 0) linfontnum = 1;
                int rows = aa / (linfontnum);           // 行数
                rows = HR / fh;// // 行数
                tcolbakdata.Add(th[0], coldata);
                //if (th[0] == "CareDate")
                //{
                //    CareDate = drw[th[0]].ToString();
                //    HRw.Add(th[0], CareDate); 
                //    continue;
                //}
                //if (th[0] == "CareTime")
                //{
                //    CareTime = drw[th[0]].ToString();
                //    HRw.Add(th[0], CareTime);
                //    continue;
                //}
                int SumFontNum = rows * linfontnum * 2;//
                if (coldata.Length > SumFontNum) drw[th[0]] = coldata.Substring(0, SumFontNum);
                else drw[th[0]] = coldata.Substring(0);
                if (coldata.Length > SumFontNum) HRw.Add(th[0], coldata.Substring(SumFontNum));
                else HRw.Add(th[0], "");
                //int Ht = (rows + 1) * fh;                //记录行高
                //SizeF sizf1 = fn.getFontSize(g, new Font("宋体", 11), "宋");
                //float Ht1 = (sizf1.Width * sizf1.Height) / float.Parse(th[4]);


            }

            return HRw;
        }
        private Hashtable getRowH(DataRow drw, Graphics g, string[] head, float fontnum, float swH, out float rowH, out int retrows, XmlNode td)
        {//计算行高

            // float hh1 = 0;
            //h=面积/w  求出一个最高的
            Hashtable rowdata = new Hashtable();
            Comm fn = new Comm();

            Brush brush = new SolidBrush(Color.FromName(td.Attributes["bgcolor"].InnerText));
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;
            if (bool.Parse(td.Attributes["i"].InnerText) == true)
                style |= FontStyle.Italic;
            if (bool.Parse(td.Attributes["u"].InnerText) == true)
                style |= FontStyle.Underline;
            GraphicsUnit fUnit = GraphicsUnit.World;
            if (td.Attributes["fontunit"] != null)
            {
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "World") fUnit = GraphicsUnit.World;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
                if (td.Attributes["fontunit"].Value == "Point") fUnit = GraphicsUnit.Point;
            }
            float fontsize = float.Parse(td.Attributes["fontsize"].InnerText);
            Font font = new Font(td.Attributes["fontname"].InnerText,
            fontsize, style, fUnit);
            tcolbakdata.Clear();
            SizeF sizf = g.MeasureString("宋", font);
            int rows = 0;
            ArrayList Userstr = new ArrayList(); //签名修改
            ArrayList Userstrtmp = new ArrayList(); //签名修改
            for (int i = 0; i < table.Columns.Count; i++)
            {
                string[] th = head[i].Split('^');
                string coldata = drw[th[0]].ToString();
                tcolbakdata.Add(th[0], coldata);　　//备份
                if (float.Parse(th[4]) < 5) continue;
                ArrayList arr = getlenstr(coldata, float.Parse(th[4]), g, font);
                if (th[0] == "User")
                {
                    int ll = arr.Count;
                    for (int d = 0; d < ll; d++)
                    {
                        string str = (string)arr[d];
                        if (str != " ") Userstr.Add(str);
                    }
                    //User = (string)arr[arr.Count - 1];
                }
                rowdata.Add(th[0], arr);
                if (arr.Count > rows) rows = arr.Count;
            }
            if (Userstr.Count < rows) //签名修改
            {
                int lll = rows - Userstr.Count;
                for (int i = 0; i < lll; i++)
                {
                    Userstrtmp.Add(" ");

                }
                for (int q = 0; q < Userstr.Count; q++)
                {
                    Userstrtmp.Add((string)Userstr[q]);

                }

                if (rowdata.Contains("User"))
                {
                    if (UserPrintDown == "Y")
                    {
                        rowdata.Remove("User");
                        rowdata.Add("User", Userstrtmp);
                    }

                }



            }
            rowH = rows * swH;
            retrows = rows;
            return rowdata;

        }
        private ArrayList getlenstr(string ss, float linew, Graphics g, Font font)
        {
            ArrayList arrlin = new ArrayList();
            ss = ss.Replace("\r\n", "\r"); //回车换行替换成换行
            char[] aa = ss.ToCharArray();
            int strlen = 0;
            string s1 = "", s2 = "";
            int i;
            float slinew = 0;
            SizeF fonsize;
            for (i = 0; i < aa.Length; i++)
            {

                fonsize = g.MeasureString(s1 + aa[i], font);
                slinew = fonsize.Width;
                if (slinew > linew)
                {
                    arrlin.Add(s1);
                    s1 = "";
                    slinew = 0;
                    i = i - 1;
                }
                else
                {

                    // string specN2 = aa[i].ToString;
                    if ((aa[i] == 38) || (aa[i] == 13) || (aa[i] == 10))
                    {
                        if (s1 == "") continue;
                        arrlin.Add(s1);
                        s1 = "";
                        slinew = 0;
                        //i = i - 1;

                    }
                    else
                    {
                        s1 = s1 + aa[i];
                    }
                }
            }
            if (s1 != "") arrlin.Add(s1);
            return arrlin;
        }
        private ArrayList getlenstrback(string ss, float linew, Graphics g, Font font)
        {
            ArrayList arrlin = new ArrayList();
            char[] aa = ss.ToCharArray();
            int strlen = 0;
            string s1 = "", s2 = "";
            int i;
            float slinew = 0;
            SizeF fonsize;
            for (i = 0; i < aa.Length; i++)
            {

                fonsize = g.MeasureString(s1 + aa[i], font);
                slinew = fonsize.Width;
                if (slinew > linew)
                {
                    arrlin.Add(s1);
                    s1 = "";
                    slinew = 0;
                    i = i - 1;
                }
                else
                {

                    s1 = s1 + aa[i];

                }
            }
            if (s1 != "") arrlin.Add(s1);
            return arrlin;
        }
        private int getstrLen(string linStr)
        {
            //If (Asc(Mid(segmentstr, i, 1)) >= 1 And Asc(Mid(segmentstr, i, 1)) <= 255) Then
            //'If Asc(Mid(segmentstr, i, 1)) > 0 Then
            //    CharNums = CharNums + 1
            //Else
            //    CharNums = CharNums + 2
            //End If
            char[] aa = linStr.ToCharArray();
            int strlen = 0;
            for (int i = 0; i < aa.Length; i++)
            {
                if ((aa[i] >= 1) && (aa[i] <= 255))
                {
                    strlen = strlen + 1;
                }
                else
                {
                    strlen = strlen + 2;
                }
            }
            return strlen;

        }
        private bool getsel(string txt, int selindex)
        {
            string[] tm = txt.Split('@');
            string[] data = tm[0].Split('!');
            string[] selitm = tm[1].Split('|');
            string sel = "";
            if (selitm.Length > 1) sel = selitm[1];
            else sel = tm[1];
            bool flag = false;
            if (sel == "") return flag;
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == "") continue;
                if (data[i] == sel)
                {
                    data[i] = data[i] + "√";
                    if (selindex == i)
                    {
                        flag = true;
                    }
                }

            }
            return flag;

        }
        private bool getmulsel(string txt, int selindex)
        {
            string[] tm = txt.Split('@');
            string[] data = tm[0].Split('!');
            string[] sel = tm[1].Split('^');
            bool flag = false;
            if (sel.Length == 0) return flag;
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == "") continue;
                for (int j = 0; j < sel.Length; j++)
                {
                    if (sel[j] == "") continue;
                    if (data[i] == sel[j])
                    {
                        if (i == selindex)
                        {
                            flag = true;
                        }
                    }
                }
            }

            return flag;
        }
        private string getmulseltxt(string txt)
        {
            if (txt == "") return "";
            string[] tm = txt.Split('@');
            string[] data = tm[0].Split('!');
            string[] sel = tm[1].Split('^');
            return tm[1].Replace('^', ' ');
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == "") continue;
                for (int j = 0; j < sel.Length; j++)
                {
                    if (sel[j] == "") continue;
                    if (data[i] == sel[j])
                    {
                        data[i] = data[i] + "√";
                    }
                }
            }
            txt = "";
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == "") continue;
                txt = txt + data[i] + "  ";
            }
            return txt;
        }


        private string gettxt(string txt)
        {
            if (txt == "") return "";
            string[] tm = txt.Split('@');
            string[] data = tm[0].Split('!');
            string[] sel = tm[1].Split('|');
            if (sel.Length > 1) return sel[1];
            else return tm[1];

        }
        private int getstpagrows(int stp, int stindex)
        {
            int strows = 0;
            for (int i = stindex; i < (stp + 1); i++)
            {
                if (HPagRow1.Count != 0)
                {
                    strows = strows + int.Parse(HPagRow1[i].ToString());
                }
                else
                {
                    strows = 0;
                }
            }
            return strows;
        }
        private void pDoc_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs ev)
        {
            if (previewPrint == "0")
            {
                if (printername != "")
                {
                    PrnDiaglog.PrinterSettings.PrinterName = printername;
                }
            }
            if (previewPrint == "0") return;
            if (showPrintDialog == true)
            {
                PrnDiaglog.Document = pDoc;
                PrnDiaglog.PrinterSettings = pDoc.PrinterSettings;
                // MessageBox.Show(PrnDiaglog.PrinterSettings.FromPage.ToString); 
                if (stPage == printpagecount)
                {
                    ev.Cancel = true;
                    return;
                }
                //PrnDiaglog.PrinterSettings.FromPage = stPage +1;
                PrnDiaglog.PrinterSettings.FromPage = curPages + 1;
                PrnDiaglog.PrinterSettings.ToPage = printpagecount;
                //PrnSumRows = 0;
                Row = 0;
                if (PrnDiaglog.ShowDialog() == DialogResult.OK)
                {
                    //取消当前的打印任务，点击打印时已生成一次打印任务，必须先取消 
                    ev.Cancel = true;
                    //页码范围正确，进行打印 
                    if ((PrnDiaglog.PrinterSettings.FromPage > 0) && (PrnDiaglog.PrinterSettings.ToPage >= PrnDiaglog.PrinterSettings.FromPage) && (PrnDiaglog.PrinterSettings.ToPage <= (printpagecount)))
                    {
                        //计算开始打印索引号 
                        StP = (PrnDiaglog.PrinterSettings.FromPage - 1);
                        //printIndex = getstpagrows(StP,1);//(this.PrnDiaglog.PrinterSettings.FromPage - 1) * this.printPagesize; 
                        printIndex = getstpagrows(StP, curPages + 1);//(this.PrnDiaglog.PrinterSettings.FromPage - 1) * this.printPagesize; 
                        Row = printIndex;
                        Pages = StP;
                        EdP = PrnDiaglog.PrinterSettings.ToPage;
                        // pDoc.PrinterSettings = PrnDiaglog.PrinterSettings; 
                        //确定打印后，必须将值设为false，否则会再继续打开打印设置 
                        showPrintDialog = false;

                        //开始打印 
                        // HPagRow.Clear();
                        pDoc.Print();
                    }
                    else
                    {
                        MessageBox.Show("输入的页码有误，请重新输入！ ");
                    }

                }
                else
                {
                    ev.Cancel = true;
                }
            }
            else
            {
                //ev.Cancel = true; 
            }
        }
        # region
        //没有特殊要求就采用一行数据
        //有单位需要小字体或上标下标的情况采用多行
        //多行时第一行要靠下不要据中
        //第一行行高采用==总高/行数
        //单行数据据中加空格
        private string PrintF(XmlNode td, Graphics g, string txt, int x, int y, int fh, int W, int rows)
        {


            // g.FillRectangle(brush, x + 1, y + 1, width, height);
            Brush brush = new SolidBrush(Color.Black);
            Point location = new Point(x, y);
            FontStyle style = FontStyle.Regular;
            //设置字体样式
            if (bool.Parse(td.Attributes["b"].InnerText) == true)
                style |= FontStyle.Bold;

            GraphicsUnit fUnit = GraphicsUnit.Point;
            string off = "";
            if (td.Attributes["offset"] != null)
            {
                off = td.Attributes["offset"].Value;
                location.Y = location.Y - int.Parse(off);
            }

            Font font = new Font(td.Attributes["fontname"].InnerText,
                float.Parse(td.Attributes["fontsize"].InnerText), style, fUnit);
            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            sf.Alignment = StringAlignment.Center;
            if ((rows > 1) && (fh != 0)) sf.LineAlignment = StringAlignment.Far;  //垂直方向靠下

            SizeF sizf = g.MeasureString(txt, font);
            Size sz = new Size(sizf.ToSize().Width + 5, sizf.ToSize().Height + 1);
            if (fh != 0) sz = new Size(sizf.ToSize().Width + 5, fh);
            if (sz.Width > W) sz.Width = W;
            //if (rows == 1) sz.Width = W;
            Rectangle rect = new Rectangle(location, sz);

            g.DrawString(txt, font, brush, rect, sf);
            return (x + ((int)sizf.Width - 3)) + "^" + sz.Height;
            // g.DrawRectangle(new Pen(Color.Black, 1), rect);


        }
        private int drawHead(Graphics g, XmlNode XHEAD)
        {
            //没有特殊要求就采用一行数据
            //有单位需要小字体或上标下标的情况采用多行
            //多行时第一行要靠下不要据中
            //第一行行高采用==总高/行数
            //单行数据据中加空格
            int RY = 0;
            int cc = 0;
            if (Pages > 0)
            {
                //tabx = MargLeft;
                cc = taby - MargTop;
                taby = MargTop;

            }

            int px = 0, py = 0;
            foreach (XmlNode xn in XHEAD.ChildNodes)
            {
                string[] pp = xn.Attributes["p"].Value.Split(',');
                string[] sz = xn.Attributes["Size"].Value.Split(',');
                Point P = new Point(int.Parse(pp[0]), int.Parse(pp[1]) - cc);
                Size PS = new Size(int.Parse(sz[0]), int.Parse(sz[1]));
                Rectangle rect = new Rectangle(P, PS);
                if (PS.Width < 5) continue;
                int X, Y;

                X = P.X; Y = P.Y;
                if (px == 0) { px = P.X; py = P.Y; };
                int fh = PS.Height / xn.ChildNodes.Count;
                int R = xn.ChildNodes.Count;
                int row = 0;
                foreach (XmlNode xl in xn.ChildNodes)
                {
                    int rh = 0;
                    if (row > 0) fh = 0;
                    string ret = "";
                    string preFs = "0";
                    XmlNode preXn = null;
                    foreach (XmlNode xf in xl.ChildNodes)
                    {
                        string fs = (xf.Attributes["fontsize"].Value);
                        string offset = (xf.Attributes["offset"].Value);
                        string dd = "";
                        if (xf.InnerText == "") dd = " ";
                        else dd = xf.InnerText;

                        if ((fs != preFs) && (preFs != "0"))  //&&(offset =="0")
                        {
                            //一行数据内有不同
                            // int offset = int.Parse(xf.Attributes["offset"].Value);
                            // Y = Y - offset;
                            //如果多行 第一行要靠下 
                            string vv = PrintF(preXn, g, ret, X, Y, fh, PS.Width, R);
                            ret = dd;
                            string[] tt = vv.Split('^');
                            // Y = Y + offset;
                            X = int.Parse(tt[0]);
                            rh = int.Parse(tt[1]);
                        }
                        else
                        {
                            ret = ret + dd;

                        }
                        preFs = fs;
                        preXn = xf;

                    }
                    if (ret != "")
                    {
                        //剩余数据打印
                        string vv = PrintF(preXn, g, ret, X, Y, fh, PS.Width, R);
                        string[] tt = vv.Split('^');
                        //  Y = Y + offset;
                        X = int.Parse(tt[0]);
                        rh = int.Parse(tt[1]);

                    }
                    row = row + 1;
                    X = P.X;
                    Y = Y + rh;

                }
                g.DrawRectangle(new Pen(Color.Black, 1), rect);
                if ((P.Y + rect.Height) > RY) RY = P.Y + rect.Height;
            }
            tabx = px; taby = py;
            return RY;
        }
        #endregion
        private void pDoc_EndPrint(object sender, System.Drawing.Printing.PrintEventArgs ev)
        {
            if (previewPrint == "0") return;

            //Comm.FactoryClass;
            //Comm.FactoryClass = null;
            showPrintDialog = true;
        }


    }
}
