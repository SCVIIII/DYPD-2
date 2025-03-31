using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using MySql.Data.MySqlClient;
using Sunny.UI;
using static DYPD_2.FClass;

namespace DYPD_2
{
    public partial class FPDY2 : UIPage
    {
        //定义所需的全局变量
        Database db;
        Editor ed;
        List<DiYaGanXian_CLASS> list_global_DiYaGanXian_calFromDiYaGui_Legal;
        List<DiYaGanXian_CLASS> list_global_DiYaGanXian_FromSec_Legal = new List<DiYaGanXian_CLASS>();
        public FPDY2()
        {
            InitializeComponent();
            // CAD所需变量初始化
            db = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.WorkingDatabase;
            ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            ////选择图块(干线及出线回路)
            //var result = usr_selectBlks_GanXianChuXian();
            //list_global_DiYaGanXian_calFromDiYaGui_Legal = result.Item1;
            //string strError_HuiLu = result.Item2;
            ////弹窗提示
            //if(!string.IsNullOrWhiteSpace(strError_HuiLu))
            //{
            //    MessageBox.Show("以下回路故障:" + strError_HuiLu);
            //}
            

        }

        private void FPDY2_Load(object sender, EventArgs e)
        {

        }

        #region 功能函数
        private (List<DiYaGanXian_CLASS>, string) usr_selectBlks_GanXianChuXian()
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //新建list 存放回路信息
                List<DiYaGui_CLASS> list_DiYaGui = new List<DiYaGui_CLASS>();
                //新建list 存放由低压柜出线计算出的所有竖向干线合法回路
                List<DiYaGanXian_CLASS> list_DiYaGanXian_CalFromDiYaGui_Legal = new List<DiYaGanXian_CLASS>();
                //新建list 存放平面中框选到的干线图块
                List<DiYaGanXian_CLASS> list_DiYaGanXian_FromSec = new List<DiYaGanXian_CLASS>();
                //新建string，存放合法性校验失败的回路编号及名称
                string strError_HuiLu = "";

                //筛选的块名存放于此
                //采用GetAnonymousBlk函数，防止匿名块导致的获取缺失
                List<string> listBlks = new List<string>()
                {
                    "DY-低压柜出线",
                    "TJSY-放射式干线-210803A",
                    "TJSY-通用干线-210812A"
                };

                //此功能需要在定义trans的前提下使用
                //筛选块的命令
                string blkTKNames = ARXTools.GetAnonymousBlk(db, listBlks);

                //防火分区框的筛选条件
                TypedValue[] acTypValAr_XT = new TypedValue[2];
                //acTypValAr_XT.SetValue(new TypedValue((int)DxfCode.LayerName, "E-火灾报警系统图专用图层"), 0);
                acTypValAr_XT.SetValue(new TypedValue((int)DxfCode.BlockName, blkTKNames), 0);
                acTypValAr_XT.SetValue(new TypedValue((int)DxfCode.Start, "INSERT"), 1);
                SelectionFilter acSelFtr_XT = new SelectionFilter(acTypValAr_XT);
                // Request for objects to be selected in the drawing area
                PromptSelectionResult acSSPrompt_XT = ed.GetSelection(acSelFtr_XT);

                //如果用户完成选择，开始执行
                if (acSSPrompt_XT.Status == PromptStatus.OK)  //当选中实体时执行
                {
                    //选中的图块数量
                    int num_acSSPrompt = acSSPrompt_XT.Value.Count;
                    //将ObjectId存为数组，方便编程调用
                    ObjectId[] id_ss = acSSPrompt_XT.Value.GetObjectIds();


                    //启用for循环,
                    
                    for (int i = 0; i < num_acSSPrompt; i++)
                    {
                        
                        //存为id便于书写
                        ObjectId id = id_ss[i];
                        string blk_Name = id.GetBlockName();

                        //对选择集中的所有低压柜出线回路进行处理
                        //将选择的块转化为class DiYaGui_CLASS 便于存入list
                        //if (blk_Name == "TJSY-放射式干线-210803A" || blk_Name == "TJSY-通用干线-210812A")
                        if(blk_Name == "DY-低压柜出线")
                        {
                            //新建第i个块的类
                            DiYaGui_CLASS diyagui = new DiYaGui_CLASS()
                            {
                                OBJID = id,
                                Pe = id.GetAttributeInBlockReference("安装功率") ?? "",
                                Purpose = id.GetAttributeInBlockReference("用途") ?? "",
                                ChangBeiYong = id.GetAttributeInBlockReference("备注2" ?? ""),
                                HuiLu = id.GetAttributeInBlockReference("回路编号") ?? "",
                                BianHao = id.GetAttributeInBlockReference("备注") ?? ""
                            }; //end of class
                            list_DiYaGui.Add(diyagui);
                        } //end of if if(blk_Name == "DY-低压柜出线")

                        //将干线图块转化为list
                        else
                        {
                            //新建干线的class
                            DiYaGanXian_CLASS ganxian_sec = new DiYaGanXian_CLASS();
                            //根据块名筛选
                            if(blk_Name == "TJSY-通用干线-210812A")
                            {
                                ganxian_sec.Purpose = id.GetAttributeInBlockReference("用途");
                                ganxian_sec.HuiLu = id.GetAttributeInBlockReference("回路编号");
                                ganxian_sec.OBJID = id;
                                list_global_DiYaGanXian_FromSec_Legal.Add(ganxian_sec);
                            }

                            else if(blk_Name == "TJSY-放射式干线-210803A")
                            {
                                //获取可见性
                                string vis = id.GetDynBlockValue("可见性1");
                                //放射式单回路
                                if(vis == "放射式单回路")
                                {
                                    //获取块属性
                                    ganxian_sec.Purpose = id.GetAttributeInBlockReference("用途");
                                    ganxian_sec.BianHao = id.GetAttributeInBlockReference("箱柜编号") ?? "";
                                    ganxian_sec.Pe = id.GetAttributeInBlockReference("安装功率") ?? "" ;
                                    ganxian_sec.HuiLu = id.GetAttributeInBlockReference("回路编号") ?? "";
                                    ganxian_sec.OBJID = id;
                                    //加入到list中
                                    list_global_DiYaGanXian_FromSec_Legal.Add(ganxian_sec);
                                }
                                //放射式双回路
                                else
                                {
                                    //获取块属性
                                    ganxian_sec.Purpose = id.GetAttributeInBlockReference("用途");
                                    ganxian_sec.BianHao = id.GetAttributeInBlockReference("箱柜编号") ?? "";
                                    ganxian_sec.Pe = id.GetAttributeInBlockReference("安装功率") ?? "";
                                    ganxian_sec.HuiLu1 = id.GetAttributeInBlockReference("常用回路") ?? "";
                                    ganxian_sec.HuiLu2 = id.GetAttributeInBlockReference("备用回路") ?? "";
                                    ganxian_sec.OBJID = id;
                                    //加入到list
                                    list_global_DiYaGanXian_FromSec_Legal.Add(ganxian_sec);
                                } //end of 放射式双回路
                            } //end of 放射式干线

                        } //end of else(干线图结束)


                    } //end of for(int i=0;i<num_acSSPrompt;i++)

                    //for循环后对list按回路编号排序
                    list_DiYaGui = list_DiYaGui.OrderBy(item => item.HuiLu).ToList();
                    //list_Purpose
                    //干线图进行计算时，剔除备用回路
                    List<string> list_Purpose = (from d in list_DiYaGui
                                                 where d.Purpose != "备用" && d.Purpose != "预留"
                                                 select d.Purpose).Distinct().ToList();

                    //使用for循环进行数据处理
                    for (int i = 0; i < list_Purpose.Count; i++)
                    {
                        //第i个Purpose对应的合理与否分析
                        List<DiYaGui_CLASS> list_i = (from d in list_DiYaGui
                                                      where d.Purpose == list_Purpose[i]
                                                      select d).ToList();

                        //单回路：用途名称 数量=1 且 常备用信息为空 
                        if (list_i.Count == 1 && list_i[0].ChangBeiYong == "" )
                        {
                            //判断是否为放射式
                            //箱柜编号不为空--放射式
                            if(!string.IsNullOrWhiteSpace(list_i[0].BianHao))
                            {
                                //将合法数据加入list
                                DiYaGanXian_CLASS ganxian = new DiYaGanXian_CLASS();
                                ganxian.Purpose = list_i[0].Purpose;
                                ganxian.Pe = list_i[0].Pe;
                                ganxian.HuiLu = list_i[0].HuiLu;
                                ganxian.BianHao = list_i[0].BianHao;
                                ganxian.FangShe = true;
                                list_DiYaGanXian_CalFromDiYaGui_Legal.Add(ganxian);
                            } //end of 放射式单回路

                            //箱柜编号为空--非放射式
                            else 
                            {
                                //将合法数据加入list
                                DiYaGanXian_CLASS ganxian = new DiYaGanXian_CLASS();
                                ganxian.Purpose = list_i[0].Purpose;
                                ganxian.Pe = list_i[0].Pe;
                                ganxian.HuiLu = list_i[0].HuiLu;
                                ganxian.FangShe = false;
                                list_DiYaGanXian_CalFromDiYaGui_Legal.Add(ganxian);
                            } //end of T接或预分支单回路

                        } //end of 单回路

                        //双回路：用途名称 数量=2 且 常备用信息有效
                        //双回路1：[0]常用 [1]备用 
                        else if (list_i.Count == 2 && list_i[0].ChangBeiYong == "常用回路" && list_i[1].ChangBeiYong == "备用回路")
                        {
                            //判断是否为放射式
                            //箱柜编号不为空--放射式
                            if (!string.IsNullOrWhiteSpace(list_i[0].BianHao))
                            {
                                //将合法数据加入list
                                DiYaGanXian_CLASS ganxian = new DiYaGanXian_CLASS();
                                ganxian.Purpose = list_i[0].Purpose;
                                ganxian.Pe = list_i[0].Pe;
                                ganxian.HuiLu1 = list_i[0].HuiLu;
                                ganxian.HuiLu2 = list_i[1].HuiLu;
                                ganxian.BianHao = list_i[0].BianHao;
                                ganxian.FangShe = true;
                                list_DiYaGanXian_CalFromDiYaGui_Legal.Add(ganxian);
                            } //end of 放射式双回路

                            //箱柜编号为空--非放射式
                            else
                            {
                                //将合法数据加入list
                                DiYaGanXian_CLASS ganxian = new DiYaGanXian_CLASS();
                                ganxian.Purpose = list_i[0].Purpose;
                                ganxian.Pe = list_i[0].Pe;
                                ganxian.HuiLu1 = list_i[0].HuiLu;
                                ganxian.HuiLu2 = list_i[1].HuiLu;
                                ganxian.FangShe = false;
                                list_DiYaGanXian_CalFromDiYaGui_Legal.Add(ganxian);
                            } //end of T接或预分支双回路
                        } //end of 双回路1

                        //双回路：用途名称 数量=2 且 常备用信息有效
                        //双回路2：[1]常用 [0]备用 
                        else if (list_i.Count == 2 && list_i[1].ChangBeiYong == "常用回路" && list_i[0].ChangBeiYong == "备用回路")
                        {
                            //判断是否为放射式
                            //箱柜编号不为空--放射式
                            if (!string.IsNullOrWhiteSpace(list_i[0].BianHao))
                            {
                                //将合法数据加入list
                                DiYaGanXian_CLASS ganxian = new DiYaGanXian_CLASS();
                                ganxian.Purpose = list_i[0].Purpose;
                                ganxian.Pe = list_i[0].Pe;
                                ganxian.HuiLu1 = list_i[1].HuiLu;
                                ganxian.HuiLu2 = list_i[0].HuiLu;
                                ganxian.BianHao = list_i[0].BianHao;
                                ganxian.FangShe = true;
                                list_DiYaGanXian_CalFromDiYaGui_Legal.Add(ganxian);
                            } //end of 放射式双回路

                            //箱柜编号为空--非放射式
                            else
                            {
                                //将合法数据加入list
                                DiYaGanXian_CLASS ganxian = new DiYaGanXian_CLASS();
                                ganxian.Purpose = list_i[0].Purpose;
                                ganxian.Pe = list_i[0].Pe;
                                ganxian.HuiLu1 = list_i[1].HuiLu;
                                ganxian.HuiLu2 = list_i[0].HuiLu;
                                ganxian.FangShe = false;
                                list_DiYaGanXian_CalFromDiYaGui_Legal.Add(ganxian);
                            } //end of T接或预分支双回路
                        } //end of 双回路2

                        //其他回路为异常回路
                        //弹窗提示
                        else
                        {
                            for(int j=0;j< list_i.Count;j++)
                            {
                                strError_HuiLu = strError_HuiLu +"\n"+ list_i[j].Purpose + "   " + list_i[j].HuiLu;
                            }
                        } //end of 异常回路


                    } //end of for









                    //无需提交事务处理
                    trans.Abort();
                } //end of if(acSSPrompt_XT.Status == PromptStatus.OK)  //当选中实体时执行

                //数据显示
                uiDG_Legal.DataSource = list_DiYaGanXian_CalFromDiYaGui_Legal;
                //返回值
                list_global_DiYaGanXian_calFromDiYaGui_Legal = list_DiYaGanXian_CalFromDiYaGui_Legal;
                return (list_DiYaGanXian_CalFromDiYaGui_Legal, strError_HuiLu);

            } //end of using trans
        }


        #endregion

        /// <summary>
        /// 按钮事件：框选图块
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton1_Click(object sender, EventArgs e)
        {
            //显示数据前，清空数据
            uiDG_Legal.DataSource = null;
            //选择图块(干线及出线回路)
            usr_selectBlks_GanXianChuXian();
        }

        /// <summary>
        /// 按钮事件：生成系统图
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiB_INSERT_Click(object sender, EventArgs e)
        {
            //启用事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //偏移距离
                int distance = 2660;
                //缩放比例
                Scale3d InsertScale = new Scale3d(1);  //插入比例
                //插入环境初始化
                ObjectId spaceid = db.CurrentSpaceId;
                //插入图层
                string layer_Insert = "0";
                //获取用户插入点
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Point3d ptStart = ARXTools.get_POINT3D(doc);

                //循环插入图块
                for (int i = 0; i < list_global_DiYaGanXian_calFromDiYaGui_Legal.Count; i++)
                {
                    var ganxian = list_global_DiYaGanXian_calFromDiYaGui_Legal[i];
                    //放射式单回路
                    if (ganxian.FangShe && !string.IsNullOrWhiteSpace(ganxian.HuiLu) )
                    {
                        //设置插入信息
                        //单回路插入点不偏移
                        Point3d ptInsert = ptStart;
                        //块名
                        string blkName = "TJSY-放射式干线-210803A";
                        Dictionary<string, string> atts1 = new Dictionary<string, string>();
                        atts1.Add("用途", ganxian.Purpose);
                        atts1.Add("箱柜编号", ganxian.BianHao);
                        atts1.Add("安装功率", ganxian.Pe + "kW");
                        ObjectId blk_objid = spaceid.InsertBlockReference(layer_Insert, blkName, ptInsert, InsertScale, 0, atts1);
                        //设置可见性、回路编号
                        Dictionary<string, string> atts2 = new Dictionary<string, string>();
                        atts2.Add("回路编号", ganxian.HuiLu);
                        blk_objid.SetDynBlockValue("可见性1", "放射式单回路");
                        blk_objid.UpdateAttributesInBlock(atts2);
                    } //end of 放射式单回路

                    //放射式双回路
                    else if (ganxian.FangShe && !string.IsNullOrWhiteSpace(ganxian.HuiLu1))
                    {
                        //设置插入信息
                        //放射式双回路插入点不偏移
                        Point3d ptInsert = ptStart;
                        //块名
                        string blkName = "TJSY-放射式干线-210803A";
                        Dictionary<string, string> atts1 = new Dictionary<string, string>();
                        atts1.Add("用途", ganxian.Purpose);
                        atts1.Add("箱柜编号", ganxian.BianHao);
                        atts1.Add("安装功率", ganxian.Pe + "kW");
                        ObjectId blk_objid = spaceid.InsertBlockReference(layer_Insert, blkName, ptInsert, InsertScale, 0, atts1);
                        //设置可见性、回路编号
                        Dictionary<string, string> atts2 = new Dictionary<string, string>();
                        atts2.Add("常用回路", "常用回路"+ganxian.HuiLu1);
                        atts2.Add("备用回路", "备用回路"+ ganxian.HuiLu2);
                        blk_objid.SetDynBlockValue("可见性1", "放射式双回路");
                        blk_objid.UpdateAttributesInBlock(atts2);
                    } //end of 放射式双回路

                    //T接或预分支的单回路
                    else if (!ganxian.FangShe && !string.IsNullOrWhiteSpace(ganxian.HuiLu))
                    {
                        //设置插入信息
                        //单回路插入点不偏移
                        Point3d ptInsert = ptStart;
                        //块名
                        string blkName = "TJSY-通用干线-210812A";
                        //块属性
                        Dictionary<string, string> atts1 = new Dictionary<string, string>();
                        atts1.Add("用途", ganxian.Purpose);
                        atts1.Add("安装功率", ganxian.Pe+"kW");
                        atts1.Add("回路编号", ganxian.HuiLu);
                        ObjectId blk_objid = spaceid.InsertBlockReference(layer_Insert, blkName, ptInsert, InsertScale, 0, atts1);
                        
                    } //end of T接或预分支的单回路

                    //T接或预分支的双回路
                    else if (!ganxian.FangShe && !string.IsNullOrWhiteSpace(ganxian.HuiLu1))
                    {
                        //设置插入信息
                        //放射式双回路插入点需偏移，且需插入两个通用干线
                        Point3d ptInsert1 = ptStart.PolarPoint(0, -1062.5);
                        Point3d ptInsert2 = ptStart.PolarPoint(0, 1062.5);
                        //块名
                        string blkName = "TJSY-通用干线-210812A";
                        Dictionary<string, string> atts1 = new Dictionary<string, string>();
                        atts1.Add("用途", ganxian.Purpose);
                        atts1.Add("回路编号", "常用回路"+ganxian.HuiLu1);
                        atts1.Add("安装功率", ganxian.Pe + "kW");

                        Dictionary<string, string> atts2 = new Dictionary<string, string>();
                        atts2.Add("用途", ganxian.Purpose);
                        atts2.Add("回路编号", "备用回路"+ganxian.HuiLu2);
                        atts2.Add("安装功率", ganxian.Pe + "kW");


                        ObjectId blk_objid1 = spaceid.InsertBlockReference(layer_Insert, blkName, ptInsert1, InsertScale, 0, atts1);
                        ObjectId blk_objid2 = spaceid.InsertBlockReference(layer_Insert, blkName, ptInsert2, InsertScale, 0, atts2);

                    } //end of T接或预分支的双回路

                    ptStart = ptStart.PolarPoint(0, distance);
                } //end of for i;

                trans.Commit();

            } //end of using trans

        } // end of private void uiB_INSERT_Click(object sender, EventArgs e)


        private void uiB_MODIFY_Click(object sender, EventArgs e)
        {
            //当有选中干线图块时才执行处理
            if(list_global_DiYaGanXian_FromSec_Legal.Count>0)
            {
                //启用事务处理
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {

                    string strError = "";
                    //循环插入图块
                    //根据低压柜的选取结果为基础
                    for (int i = 0; i < list_global_DiYaGanXian_calFromDiYaGui_Legal.Count; i++)
                    {
                        //由低压柜计算出的合法的回路
                        DiYaGanXian_CLASS ganxian_cal = list_global_DiYaGanXian_calFromDiYaGui_Legal[i];
                        //平面中已框选到的干线图块
                        List<DiYaGanXian_CLASS> ganxian_To_Modify = (from d in list_global_DiYaGanXian_FromSec_Legal
                                                                     where d.Purpose == ganxian_cal.Purpose
                                                                     select d).ToList();

                        //放射式单回路
                        if (ganxian_cal.FangShe && !string.IsNullOrWhiteSpace(ganxian_cal.HuiLu) && ganxian_To_Modify.Count == 1)
                        {
                            //设置插入信息
                            //获取OBJID
                            ObjectId blk_objid = ganxian_To_Modify[0].OBJID;
                            Dictionary<string, string> atts1 = new Dictionary<string, string>();
                            atts1.Add("用途", ganxian_cal.Purpose);
                            atts1.Add("箱柜编号", ganxian_cal.BianHao);
                            atts1.Add("安装功率", ganxian_cal.Pe + "kW");
                            atts1.Add("回路编号", ganxian_cal.HuiLu);
                            //设置块属性值
                            blk_objid.UpdateAttributesInBlock(atts1);
                        } //end of 放射式单回路

                        //放射式双回路
                        else if (ganxian_cal.FangShe && !string.IsNullOrWhiteSpace(ganxian_cal.HuiLu1) && !string.IsNullOrWhiteSpace(ganxian_cal.HuiLu1) && ganxian_To_Modify.Count == 1)
                        {
                            //设置插入信息
                            //获取OBJID
                            ObjectId blk_objid = ganxian_To_Modify[0].OBJID;
                            Dictionary<string, string> atts1 = new Dictionary<string, string>();
                            atts1.Add("用途", ganxian_cal.Purpose);
                            atts1.Add("箱柜编号", ganxian_cal.BianHao);
                            atts1.Add("安装功率", ganxian_cal.Pe + "kW");
                            atts1.Add("常用回路", "常用回路" + ganxian_cal.HuiLu1);
                            atts1.Add("备用回路", "备用回路" + ganxian_cal.HuiLu2);
                            //设置块属性值
                            blk_objid.UpdateAttributesInBlock(atts1);
                        } //end of 放射式双回路

                        //T接或预分支的单回路
                        else if (!ganxian_cal.FangShe && !string.IsNullOrWhiteSpace(ganxian_cal.HuiLu) && ganxian_To_Modify.Count == 1)
                        {
                            //设置插入信息
                            //获取OBJID
                            ObjectId blk_objid = ganxian_To_Modify[0].OBJID;
                            Dictionary<string, string> atts1 = new Dictionary<string, string>();
                            atts1.Add("用途", ganxian_cal.Purpose);
                            atts1.Add("安装功率", ganxian_cal.Pe + "kW");
                            atts1.Add("回路编号", ganxian_cal.HuiLu);
                            //设置块属性值
                            blk_objid.UpdateAttributesInBlock(atts1);

                        } //end of T接或预分支的单回路

                        //T接或预分支的双回路
                        else if (!ganxian_cal.FangShe && !string.IsNullOrWhiteSpace(ganxian_cal.HuiLu1) && ganxian_To_Modify.Count == 2)
                        {
                            //声明变量objid
                            ObjectId blk_objid1 = ganxian_To_Modify[0].OBJID;
                            ObjectId blk_objid2 = ganxian_To_Modify[1].OBJID;
                            //如果ganxian_To_Modify中的第2个元素为常用回路，纠正设置
                            if (ganxian_To_Modify[1].HuiLu1.StartsWith("常用回路"))
                            {
                                blk_objid1 = ganxian_To_Modify[1].OBJID;
                                blk_objid2 = ganxian_To_Modify[0].OBJID;

                            }

                            //块名属性
                            Dictionary<string, string> atts1 = new Dictionary<string, string>();
                            atts1.Add("用途", ganxian_cal.Purpose);
                            atts1.Add("回路编号", "常用回路" + ganxian_cal.HuiLu1);
                            atts1.Add("安装功率", ganxian_cal.Pe + "kW");

                            Dictionary<string, string> atts2 = new Dictionary<string, string>();
                            atts2.Add("用途", ganxian_cal.Purpose);
                            atts2.Add("回路编号", "备用回路" + ganxian_cal.HuiLu2);
                            atts2.Add("安装功率", ganxian_cal.Pe + "kW");

                            //设置块属性值
                            blk_objid1.UpdateAttributesInBlock(atts1);
                            blk_objid2.UpdateAttributesInBlock(atts2);

                        } //end of T接或预分支的双回路

                        //其余异常回路存入提示框内
                        else
                        {
                            strError = strError + "\n" + "用途:" + ganxian_cal.Purpose;
                        }
                    } //end of for i;

                    trans.Commit();

                    //弹窗提示
                    if (!string.IsNullOrWhiteSpace(strError))
                    {
                        MessageBox.Show("以下回路异常:" + strError);
                    }
                    //无异常时弹窗提示
                    else
                    {
                        MessageBox.Show("修改完成");
                    } //enmd of 弹窗提示

                } //end of using trans
            } //end of if(list_global_DiYaGanXian_FromSec_Legal.Count>0)
            
            //未框选图块时弹窗提示
            //防止程序崩溃
            else
            {
                MessageBox.Show("没有选择干线图块，请点击 框选 按钮重新框选图块");
            } //end of else

        } //end of private void uiB_MODIFY_Click(object sender, EventArgs e)

        private void FPDY2_Initialize(object sender, EventArgs e)
        {

        }
    } //end of class
} //end of namespace

