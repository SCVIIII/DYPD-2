using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using MySql.Data.MySqlClient;
using Org.BouncyCastle.Crypto;
using Sunny.UI;
using static DYPD_2.FClass;

namespace DYPD_2
{
    /// <summary>
    ///二次配的低压柜系统图拆分为单独Page
    ///便于程序按功能划分
    /// </summary>
    public partial class FPDY1 : UIPage
    {

        //定义所需的全局变量
        Database db;
        Editor ed;
        MySqlConnection conn;
        ObjectId[] id_ss;
        FMain form_this;

        public FPDY1(FMain form_this_in)
        {
            //窗口初始化程序
            InitializeComponent();
            form_this = form_this_in;

            //CAD所需变量初始化
            db = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.WorkingDatabase;
            ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            //数据库信息初始化
            //异常处理待添加：数据库连接失败时，返回的string为""
            string connStr = USRSQLTools.auto_login(db);

            //云数据库连接失败时弹窗提示
            if( string.IsNullOrWhiteSpace(connStr))
            {
                MessageBox.Show("云数据库连接失败，请尝试以下解决方法：" + "\n1.检查网络连接" + "\n2.登录控制台检查云数据库状态");
            }
            //数据库连接正常时继续
            else
            {
                //设置数据库连接
                conn = new MySqlConnection(connStr);
                //选择图块
                //usr_selectBlks();
            } //end of if( string.IsNullOrWhiteSpace(connStr))

        } //end of public FPDY1(FMain form_this_in)


        #region 子函数
        /// <summary>
        /// 为馈出回路datagridview设置列信息、添加按钮列
        /// </summary>
        private void InitializeDatagridview(UIDataGridView uiDataGridView1)
        {
            //行高度
            uiDataGridView1.RowTemplate.Height = 26;

            //创建DataGridViewButtonColumn列
            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();
            buttonColumn.HeaderText = "计算"; // 设置列标题
            buttonColumn.Name = "计算"; // 列的名称
            buttonColumn.Width = 70;

            // 在指定位置（第 16 列）插入列
            uiDataGridView1.Columns.Insert(16, buttonColumn);
            // 设置列的相关属性
            buttonColumn.Text = "计算"; // 按钮的文本
            buttonColumn.UseColumnTextForButtonValue = true; // 设置按钮的显示文本

            //列号命名：用名称替换序号，仅在开头定义后，后续调整无需复核列号
            //块名
            uiDataGridView1.Columns[1].Name = "块名";
            //块OBJID
            uiDataGridView1.Columns[2].Name = "OBJID";
            //块位置
            uiDataGridView1.Columns[3].Name = "块位置";
            //消防
            uiDataGridView1.Columns[4].Name = "消防";
            //切非
            uiDataGridView1.Columns[5].Name = "切非";
            //计量
            uiDataGridView1.Columns[6].Name = "计量";
            //常备用
            uiDataGridView1.Columns[7].Name = "常备用";
            //回路用途
            uiDataGridView1.Columns[8].Name = "回路用途";
            //箱名编号
            uiDataGridView1.Columns[9].Name = "箱名编号";
            //回路编号
            uiDataGridView1.Columns[10].Name = "回路编号";
            //Pe
            uiDataGridView1.Columns[11].Name = "Pe";
            //Kx
            uiDataGridView1.Columns[12].Name = "Kx";
            //Cos
            uiDataGridView1.Columns[13].Name = "Cos";
            //Ijs
            uiDataGridView1.Columns[14].Name = "Ijs";
            //Izd
            uiDataGridView1.Columns[15].Name = "Izd";
            //断路器
            uiDataGridView1.Columns[17].Name = "断路器";
            //电缆类型
            uiDataGridView1.Columns[18].Name = "电缆类型";
            //电缆截面
            uiDataGridView1.Columns[19].Name = "电缆截面";
            //互感器
            uiDataGridView1.Columns[20].Name = "互感器";

            //块名
            uiDataGridView1.Columns["块名"].Visible = false;
            uiDataGridView1.Columns["块名"].HeaderText = "块名";
            //块OBJID
            uiDataGridView1.Columns["OBJID"].Visible = false;
            uiDataGridView1.Columns["OBJID"].HeaderText = "OBJID";
            //块位置
            uiDataGridView1.Columns["块位置"].Visible = false;
            uiDataGridView1.Columns["块位置"].HeaderText = "块位置";
            //消防
            uiDataGridView1.Columns["消防"].Width = 50;
            uiDataGridView1.Columns["消防"].HeaderText = "消防";
            //切非
            uiDataGridView1.Columns["切非"].Width = 50;
            uiDataGridView1.Columns["切非"].HeaderText = "切非";
            //计量
            uiDataGridView1.Columns["计量"].Width = 50;
            uiDataGridView1.Columns["计量"].HeaderText = "计量";
            //常备用
            uiDataGridView1.Columns["常备用"].Width = 80;
            uiDataGridView1.Columns["常备用"].HeaderText = "常备用";
            //回路用途Purpose（中文）
            uiDataGridView1.Columns["回路用途"].Width = 150;
            uiDataGridView1.Columns["回路用途"].HeaderText = "回路用途";
            //箱名编号
            uiDataGridView1.Columns["箱名编号"].Width = 120;
            uiDataGridView1.Columns["箱名编号"].HeaderText = "箱名编号";
            //回路编号
            uiDataGridView1.Columns["回路编号"].Width = 100;
            //安装功率Pe
            uiDataGridView1.Columns["Pe"].Width = 50;
            //需要系数Kx
            uiDataGridView1.Columns["Kx"].Width = 50;
            //功率因数Cos
            uiDataGridView1.Columns["Cos"].Width = 50;
            //Ijs
            uiDataGridView1.Columns["Ijs"].Width = 70;
            //Izd
            uiDataGridView1.Columns["Izd"].Width = 50;
            //断路器
            uiDataGridView1.Columns["断路器"].Width = 120;
        } //end of InitializeDatagridview

        /// <summary>
        /// 由图块转为数据
        /// 函数:为DY_ChuXian类设置备用回路的所需信息
        /// 返回值:无
        /// </summary>
        /// <param name="id"></param>
        /// <param name="chuXian"></param>
        private void cal_blk_BeiYongHuiLu(ObjectId id, DY_ChuXian chuXian)
        {
            chuXian.HuiluBianhao = id.GetAttributeInBlockReference("回路编号");
            chuXian.Cable1 = "";
            chuXian.Cable2 = "";
            chuXian.Switch = id.GetAttributeInBlockReference("断路器");
            chuXian.Transformer = id.GetAttributeInBlockReference("互感器");
            //引入decimal.TryParse 防止因值为空时导致错误
            chuXian.Pe = (decimal)0;
            chuXian.Kx = (decimal)0.9;
            chuXian.Cos = (decimal)0.85;
            chuXian.ChangBei = "";
            chuXian.Ijs = (decimal)0;
            chuXian.Izd = int.TryParse(id.GetAttributeInBlockReference("整定电流"), out int tempIzd) ? tempIzd : 50;


            string strPurpose = id.GetAttributeInBlockReference("用途");
            //设置：是否为消防负荷
            chuXian.XiaoFang = autoIF_XiaoFang(id, strPurpose);
            //设置：是否切非
            chuXian.QieFei = autoIF_QieFei(id, strPurpose, chuXian.XiaoFang);
            //设置：是否计量
            chuXian.JiLiang = autoIF_JiLiang(id, strPurpose, chuXian.XiaoFang);
        } //end of private void set_BeiYongHuiLu(ObjectId id,DY_ChuXian chuXian)

        /// <summary>
        /// 由图块转为数据
        /// 函数:为DY_ChuXian类设置出线回路的所需信息
        /// 返回值:无
        /// </summary>
        /// <param name="id"></param>
        /// <param name="chuXian"></param>
        private void cal_blk_ChuXianHuiLu(ObjectId id, DY_ChuXian chuXian)
        {
            chuXian.Name = id.GetAttributeInBlockReference("备注");
            chuXian.HuiluBianhao = id.GetAttributeInBlockReference("回路编号");
            chuXian.Cable1 = id.GetAttributeInBlockReference("电缆类型");
            chuXian.Cable2 = id.GetAttributeInBlockReference("电缆规格");
            chuXian.Switch = id.GetAttributeInBlockReference("断路器");
            chuXian.Transformer = id.GetAttributeInBlockReference("互感器");
            //引入decimal.TryParse 防止因值为空时导致错误
            chuXian.Pe = decimal.TryParse(id.GetAttributeInBlockReference("安装功率"), out decimal tempPe) ? tempPe : (decimal)0;
            chuXian.Kx = decimal.TryParse(id.GetAttributeInBlockReference("需要系数"), out decimal tempKx) ? tempKx : (decimal)0.9;
            chuXian.Cos = decimal.TryParse(id.GetAttributeInBlockReference("功率因数"), out decimal tempCos) ? tempCos : (decimal)0.85;
            chuXian.ChangBei = id.GetAttributeInBlockReference("备注2");
            chuXian.Ijs = decimal.TryParse(id.GetAttributeInBlockReference("计算电流"), out decimal tempIjs) ? tempIjs : (decimal)0;
            chuXian.Izd = int.TryParse(id.GetAttributeInBlockReference("整定电流"), out int tempIzd) ? tempIzd : 50;
            //消防、切非、计量信息读取
            //功能待补充：当以上信息为空时，自动根据负荷名称设置默认值

            string strPurpose = id.GetAttributeInBlockReference("用途");
            //设置：是否为消防负荷
            chuXian.XiaoFang = autoIF_XiaoFang(id, strPurpose);
            //设置：是否切非
            chuXian.QieFei = autoIF_QieFei(id, strPurpose, chuXian.XiaoFang);
            //设置：是否计量
            chuXian.JiLiang = autoIF_JiLiang(id, strPurpose, chuXian.XiaoFang);

        } //end of private void set_ChuXianHuiLu(ObjectId id, DY_ChuXian chuXian)

        /// <summary>
        /// 由图块转为数据
        /// 函数:为DY_JinXian类设置出线回路的所需信息
        /// </summary>
        /// <param name="id_JinXian"></param>
        /// <returns></returns>
        private DY_JinXian cal_blk_JinXian(ObjectId id_JinXian, decimal sumPe)
        {
            //返回值
            DY_JinXian jinxian = new DY_JinXian();
            //获取基础信息
            jinxian.Name = id_JinXian.GetAttributeInBlockReference("配电箱编号") ?? "";
            jinxian.Switch = id_JinXian.GetAttributeInBlockReference("断路器") ?? "";
            jinxian.Transformer = id_JinXian.GetAttributeInBlockReference("互感器") ?? "";
            //引入decimal.TryParse 防止因值为空时导致错误
            jinxian.Pe = sumPe;
            jinxian.Kx = decimal.TryParse(id_JinXian.GetAttributeInBlockReference("需要系数"), out decimal tempKx2) ? tempKx2 : (decimal)0.9;
            jinxian.Cos = decimal.TryParse(id_JinXian.GetAttributeInBlockReference("功率因数"), out decimal tempCos2) ? tempCos2 : (decimal)0.85;
            jinxian.Ijs = decimal.TryParse(id_JinXian.GetAttributeInBlockReference("计算电流"), out decimal tempIjs2) ? tempIjs2 : (decimal)0;
            jinxian.Izd = int.TryParse(id_JinXian.GetAttributeInBlockReference("整定电流"), out int tempIzd2) ? tempIzd2 : 50;
            //设置进线柜的uiTextbox
            uiTextBox_Pe.Text = jinxian.Pe.ToString();
            uiTextBox_Kx.Text = jinxian.Kx.ToString();
            uiTextBox_Cos.Text = jinxian.Cos.ToString();
            uiTextBox_Ijs.Text = Math.Round(jinxian.Ijs, 2).ToString();
            uiTextBox_Izd.Text = jinxian.Izd.ToString();
            //返回值
            return jinxian;
        } //end of private DY_JinXian cal_JinXian(ObjectId id_JinXian)


        /// <summary>
        /// 从平面框选图块,并交由对应函数进行处理，最终由datagridview显示结果
        /// </summary>
        private void usr_selectBlks()
        {
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {

                //将选中的行转为list
                //List<FASJiSuanshuExcel> list_Insert = Common.MyTools.UIDGSecToFasList(uiDG_All);

                //筛选的块名存放于此
                //采用GetAnonymousBlk函数，防止匿名块导致的获取缺失
                List<string> listBlks = new List<string>()
                {
                    "DY-低压柜出线",
                    "TJSY-低压配进线-210803A"
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
                    id_ss = acSSPrompt_XT.Value.GetObjectIds();

                    //新建返回值:出线柜的list类：LIST_CHAXUN
                    List<DY_ChuXian> LIST_CHAXUN = new List<DY_ChuXian>();
                    //新建值：sumPe总和
                    decimal sumPe = 0;
                    decimal sumPe_XiaoFang = 0;
                    decimal sumPe_FeiXiaoFang = 0;
                    //进线模块的objid
                    ObjectId id_JinXian = new ObjectId();
                    //进线模块的class

                    //foreach遍历寻找进线模块
                    foreach (var id_i in id_ss)
                    {
                        if (id_i.GetBlockName().ToUpper() == "TJSY-低压配进线-210803A")
                        {
                            id_JinXian = id_i;
                            uiTextBox_OBJID.Text = id_JinXian.ToString();
                        }
                    }

                    //对选择集进行循环处理
                    for (int i = 0; i < num_acSSPrompt; i++)
                    {
                        //筛选低压配电图块"DY-低压柜出线"
                        ObjectId id = id_ss[i];
                        string BlkName = id.GetBlockName().ToUpper();
                        if (BlkName == "DY-低压柜出线")
                        {
                            //读取当前出线的信息
                            DY_ChuXian chuXian = new DY_ChuXian();
                            chuXian.BlkName = BlkName;
                            chuXian.BlkObjID = id.ToString();
                            chuXian.BlkPoint = id.GetInsertPoint();
                            //从块属性中读取信息
                            chuXian.Purpose = id.GetAttributeInBlockReference("用途");

                            //读取快信息
                            //首先判断，该回路是否为预留
                            if (chuXian.Purpose == "备用" || chuXian.Purpose == "预留")
                            {
                                //设置备用回路的信息
                                cal_blk_BeiYongHuiLu(id, chuXian);
                            } //end of if 备用

                            //馈出回路时正常运算：
                            else
                            {
                                //设置出线回路的信息
                                cal_blk_ChuXianHuiLu(id, chuXian);
                                //安装功率累加
                                //增加情况：消防、非消防负荷分开求和
                                if (id.GetAttributeInBlockReference("是否消防负荷") == "是")
                                {
                                    sumPe_XiaoFang = sumPe_XiaoFang + chuXian.Pe;
                                }
                                else
                                {
                                    sumPe_FeiXiaoFang = sumPe_FeiXiaoFang + chuXian.Pe;
                                }
                            } //end of else

                            //将这一出线回路添加至list中
                            LIST_CHAXUN.Add(chuXian);

                        } //end of if BlkName == DY-低压柜出线

                        else if (blkTKNames == "TJSY-低压配进线-210803A")
                        {
                            id_JinXian = id;
                        } //end of else if(blkTKNames == "TJSY-低压配进线-210803A")
                    } //end of for


                    //使用LINQ对list进行排序
                    LIST_CHAXUN = LIST_CHAXUN.OrderBy(item => item.BlkPoint.X).ToList();

                    //对进线模块进行信息处理
                    //增加if判断,未选中进线模块时，跳过此段
                    if (id_JinXian.IsValid)
                    {
                        //求sumPe
                        sumPe = sumPe_FeiXiaoFang;
                        if (sumPe_XiaoFang > sumPe_FeiXiaoFang)
                        {
                            sumPe = sumPe_XiaoFang;
                        }

                        //进线模块信息计算
                        DY_JinXian jinxian = cal_blk_JinXian(id_JinXian, sumPe);

                        //在uiTextBox6中显示进线柜状态
                        uiTextBox6.Text = "获取成功";
                        //uiTextBox_Name中显示柜名
                        uiTextBox_Name.Text = jinxian.Name;
                        //在radio中显示当前的开关类型
                        //默认值为负荷开关
                        string Switch_Type = id_JinXian.GetDynBlockValue("可见性1") ?? "负荷开关";
                        if (Switch_Type == "断路器")
                        {
                            uiRadioButtonGroup1.SelectedIndex = 1;
                        }
                        else if (Switch_Type == "熔断器")
                        {
                            uiRadioButtonGroup1.SelectedIndex = 2;
                        } //默认的开关类型设置结束


                        //for循环修改回路编号
                        string preHuiLu = jinxian.Name + "/WP";
                        for (int i = 0; i < LIST_CHAXUN.Count; i++)
                        {
                            //第i个的回路编号:1-1P1/WP(i+1)
                            LIST_CHAXUN[i].HuiluBianhao = preHuiLu + (i + 1).ToString();
                        } //end of for
                    } //end of if (id_JinXian.IsValid)

                    //id_JinXian不合法，说明没有框选到进线模块
                    //因此不对进线模块进行数据处理
                    else
                    {
                        //标识符置空
                        uiTextBox6.Text = "";
                    } //end of if/else

                    //绑定数据源;此操作应在赋回路编号后进行
                    uiDataGridView1.DataSource = LIST_CHAXUN;
                    //设置列信息、列名、添加按钮列
                    InitializeDatagridview(uiDataGridView1);

                    //无需提交事务处理
                    trans.Abort();
                } //end of if(acSSPrompt_XT.Status == PromptStatus.OK)  //当选中实体时执行
            } //end of using trans
        } //end of usr_selectBlks


        /// <summary>
        /// 子函数：判断是否为消防负荷
        /// 自带输入合法性校验、根据名称自动设置是否消防
        /// </summary>
        /// <param name="id"></param>
        /// <param name="chuXian"></param>
        /// <param name="strPurpose"></param>
        private static bool autoIF_XiaoFang(ObjectId id, string strPurpose)
        {
            string XiaoFang = id.GetAttributeInBlockReference("是否消防负荷");
            //创建返回值，而不是直接传入Chuxian.XiaoFang,便于主程序书写统一
            //消防负荷较少，所以默认返回值为否
            bool result = false;
            //有效输入：“是”——消防负荷
            if (XiaoFang == "是")
            {
                result = true;
            }
            //有效输入：“否”——非消防负荷
            else if (XiaoFang == "否")
            {
                result = false;
            }
            //其他情况为无效输出，自动设置默认值
            //不包括"非消防"且含有以下关键词为消防：消防、应急
            else if
                ((!strPurpose.Contains("非消防")) &&
                (strPurpose.Contains("消防")) || (strPurpose.Contains("应急"))
                )
            {
                result = true;
            }
            //其他情况，置否
            else
            {
                result = false;
            }
            //返回结果
            return result;
        } //end of  private static bool autoIF_XiaoFang

        /// <summary>
        /// 子函数：判断是否为切非
        /// 自带输入合法性校验、根据名称自动设置是否切非
        /// </summary>
        /// <param name="id"></param>
        /// <param name="chuXian"></param>
        /// <param name="strPurpose"></param>
        private static bool autoIF_QieFei(ObjectId id, string strPurpose, bool XiaoFang)
        {
            string QieFei = id.GetAttributeInBlockReference("是否切非");
            //创建返回值，而不是直接传入Chuxian.Qie,便于主程序书写统一
            bool result = true;
            //有效输入：“是”——消防负荷
            if (XiaoFang)
            {
                result = false;
            }
            //有效输入：“否”——非消防负荷
            else if (QieFei == "是")
            {
                result = true;
            }
            //其他情况为无效输出，自动设置默认值
            //不包括"非消防"且含有以下关键词为消防：消防、应急
            else if
                (
                (strPurpose.Contains("照明")) || (strPurpose.Contains("弱电机房")) || (strPurpose.Contains("备用"))
                )
            {
                result = false;
            }
            //其他情况，置否
            else
            {
                result = false;
            }
            //返回结果
            return result;
        } //end of  private static bool autoIF_XiaoFang

        /// <summary>
        /// 子函数：判断是否设置计量电表
        /// </summary>
        /// <param name="id"></param>
        /// <param name="strPurpose"></param>
        /// <param name="XiaoFang"></param>
        /// <returns></returns>
        private static bool autoIF_JiLiang(ObjectId id, string strPurpose, bool XiaoFang)
        {
            string JiLiang = id.GetAttributeInBlockReference("是否计量");
            //创建返回值，而不是直接传入Chuxian.Qie,便于主程序书写统一
            bool result = true;
            //有效输入：“否”——不计量
            if (JiLiang == "否")
            {
                result = false;
            }

            //其他情况，置是
            else
            {
                result = true;
            }
            //返回结果
            return result;
        } //end of  private static bool autoIF_XiaoFang



        //行计算
        private decimal cal_row(DataGridViewRow row)
        {
            //返回值:Pe
            decimal result_Pe = 0;

            string strPurpose = row.Cells["回路用途"].Value.ToString();
            //根据行号反推回路编号
            if (uiTextBox6.Text == "获取成功" && !string.IsNullOrWhiteSpace(uiTextBox_Name.Text))
            {
                //所在行号
                int row_index = row.Index;
                //回路编号=行号+1
                int HuiLuBianHao = row_index + 1;
                //前缀
                string preName = uiTextBox_Name.Text;
                //设置回路编号
                row.Cells["回路编号"].Value = preName + "/WP" + HuiLuBianHao;

            } //end of 行号

            //判断是否为备用回路
            //备用回路设置名称、回路编号、Izd、断路器、互感器
            //以下置0:Pe,Ijs
            //以下置空:电缆类型,电缆截面,常备用
            //以下置默认:Cos,Kx,
            //备用回路消防--否；切非--否；计量--是
            if (strPurpose == "备用" || strPurpose == "预留")
            {
                //读取整定电流,默认值为50A
                int Izd = 50;
                if (int.TryParse(row.Cells["Izd"].Value.ToString(), out Izd))
                {
                    //根据整定电流，从云数据库中查询电缆截面、保护管径
                    CABLE_CLASS result_Cable = USRSQLTools.query_Cable(conn, Izd);
                    MCB_CLASS result_MCB = USRSQLTools.query_MCB(conn, Izd);

                    //点击计算按钮后，自动勾选checkbox
                    row.Cells[0].Value = true;
                    //将计算结果显示在uiDataGridView1中
                    row.Cells["断路器"].Value = result_MCB.Info1 + result_MCB.Info2;
                    row.Cells["互感器"].Value = result_Cable.CT;
                    row.Cells["消防"].Value = false;
                    row.Cells["切非"].Value = false;
                    row.Cells["计量"].Value = true;
                    row.Cells["常备用"].Value = "";
                    row.Cells["回路用途"].Value = strPurpose;
                    row.Cells["Pe"].Value = (decimal)0;
                    row.Cells["Kx"].Value = (decimal)0;
                    row.Cells["Cos"].Value = (decimal)0;
                    row.Cells["Ijs"].Value = (decimal)0;
                    row.Cells["电缆类型"].Value = "";
                    row.Cells["电缆截面"].Value = "";
                }
                //整定电流错误时弹窗提示
                else
                {
                    MessageBox.Show("整定电流未设置");
                }
            } //end of if(strPurpose == "备用" || strPurpose == "预留")
            else
            {
                //声明Pe、Ijs
                decimal Pe, Ijs;
                decimal G3 = (decimal)Math.Sqrt(3);
                decimal Kx = (decimal)0.9;
                decimal Cos = (decimal)0.85;
                string tempPe = row.Cells["Pe"].Value.ToString();
                //尝试此行进行计算，负荷计算时，Pe不能为空
                //Kx、Cos可为空，默认值分别为0.9、0.85
                if (decimal.TryParse(tempPe, out Pe))
                {
                    //尝试根据现有的值设置Kx、Cos
                    //待补充：校验大小，Kx、COs均应在0~1之间
                    try
                    {
                        result_Pe = Pe;
                        Kx = decimal.Parse(row.Cells["Kx"].Value.ToString());
                        Cos = decimal.Parse(row.Cells["Cos"].Value.ToString());
                    }
                    //暂未设置catch
                    catch { }
                    //求计算电流
                    Ijs = Pe * Kx / Cos / (decimal)0.38 / G3;


                    //根据1.25倍计算电流求整定电流
                    int Izd = calCurrent_Izd(Ijs);

                    //根据整定电流，从云数据库中查询电缆截面、保护管径
                    CABLE_CLASS result_Cable = USRSQLTools.query_Cable(conn, Izd);
                    MCB_CLASS result_MCB = USRSQLTools.query_MCB(conn, Izd);

                    //点击计算按钮后，自动勾选checkbox
                    row.Cells[0].Value = true;
                    //将计算结果显示在uiDataGridView1中
                    row.Cells["Ijs"].Value = Ijs;
                    row.Cells["Izd"].Value = Izd;
                    //row.Cells["断路器"].Value = Izd;
                    row.Cells["电缆截面"].Value = "4x" + result_Cable.Cable + "+E" + result_Cable.CablePE;
                    row.Cells["互感器"].Value = result_Cable.CT;
                    //断路器信息设置
                    //消防负荷加MA
                    if (row.Cells["消防"].Value.ToString().ToUpper() == "TRUE")
                    {
                        row.Cells["断路器"].Value = result_Cable.MCB + "/MA";
                    }
                    //切非回路加MX+OF
                    else if (row.Cells["切非"].Value.ToString().ToUpper() == "TRUE")
                    {
                        row.Cells["断路器"].Value = result_Cable.MCB + "/MX+OF";
                    }
                    //其他负荷无附件
                    else
                    {
                        row.Cells["断路器"].Value = result_Cable.MCB;
                    }
                } //end of if(decimal.TryParse(tempPe, out Pe))


                //Pe值无效时，弹窗提示
                else
                {
                    // 转换失败，decimalValue 包含默认值 (0)
                    MessageBox.Show("转换失败，默认值为：" + tempPe);
                }
            } //end of else
            return result_Pe;
        } //end of private void cal_row(DataGridViewRow row)

        /// <summary>
        /// 二级子函数：根据1.25倍的Ijs求整定电流
        /// </summary>
        /// <param name="Ijs"></param>
        /// <returns></returns>
        private static int calCurrent_Izd(decimal Ijs)
        {
            //函数：求整定电流
            //计算电流的1.25倍：Ical
            decimal Ical = Ijs * (decimal)1.25;
            int Izd = 0;

            //整定电流20A
            if (Ical > 0 && Ical <= 20)
            {
                Izd = 20;
            }
            else if (Ical > 20 && Ical <= 25)
            {
                Izd = 25;
            } //
            else if (Ical > 25 && Ical <= 32)
            {
                Izd = 32;
            } //
            else if (Ical > 32 && Ical <= 40)
            {
                Izd = 40;
            } //
            else if (Ical > 40 && Ical <= 50)
            {
                Izd = 50;
            } //
            else if (Ical > 50 && Ical <= 63)
            {
                Izd = 63;
            } //
            else if (Ical > 63 && Ical <= 80)
            {
                Izd = 80;
            } //
            else if (Ical > 80 && Ical <= 100)
            {
                Izd = 100;
            } //
            else if (Ical > 100 && Ical <= 125)
            {
                Izd = 125;
            } //
            else if (Ical > 125 && Ical <= 140)
            {
                Izd = 140;
            } //
            else if (Ical > 140 && Ical <= 160)
            {
                Izd = 160;
            } //
            else if (Ical > 160 && Ical <= 180)
            {
                Izd = 180;
            } //
            else if (Ical > 180 && Ical <= 200)
            {
                Izd = 200;
            } //
            else if (Ical > 200 && Ical <= 225)
            {
                Izd = 225;
            } //
            else if (Ical > 225 && Ical <= 250)
            {
                Izd = 250;
            } //

            else if (Ical > 250 && Ical <= 315)
            {
                Izd = 315;
            }
            //整定电流350A
            else if (Ical > 315 && Ical <= 350)
            {
                Izd = 350;
            }

            //电流整定失败时弹窗提示
            else
            {
                MessageBox.Show("电流整定失败，当前计算电流：" + Math.Round(Ijs, 1) + "\n当前1.25倍Ijs：  " + Math.Round(Ical, 1));
            }

            return Izd;
            //整定电流计算结束
        }


        /// <summary>
        /// 一级子函数：回路进行负荷计算，仅影响datagridview，不做平面修改
        /// </summary>
        /// <returns></returns>
        private List<DY_ChuXian> calDG_ToChuXianClass()
        {

            //新建类，用于存放datagridview转回来的list
            List<DY_ChuXian> list_Xuanzhong = new List<DY_ChuXian>();
            for (int i = 0; i < uiDataGridView1.RowCount; i++)
            {
                DataGridViewRow row = uiDataGridView1.Rows[i];

                DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)uiDataGridView1.Rows[i].Cells[0];
                //Boolean flag = Convert.ToBoolean(checkCell.Value);
                //if (flag == true)     //查找被选择的数据行
                if (true)
                {
                    //新建DY_ChuXian类,并将数据逐格填入
                    DY_ChuXian chuxian = new DY_ChuXian
                    {
                        BlkObjID = row.Cells["OBJID"].Value.ToString(),

                        //= row.Cells[""].Value.ToString(),
                        Purpose = row.Cells["回路用途"].Value.ToString(),
                        //Name = row.Cells["箱名编号"].Value.ToString(),
                        Name = row.Cells["箱名编号"].Value?.ToString() ?? "", // 使用 null 合并运算符,
                        HuiluBianhao = row.Cells["回路编号"].Value.ToString(),

                        Pe = decimal.Parse(row.Cells["Pe"].Value.ToString()),
                        Kx = decimal.Parse(row.Cells["Kx"].Value.ToString()),
                        Cos = decimal.Parse(row.Cells["Cos"].Value.ToString()),
                        Ijs = decimal.Parse(row.Cells["Ijs"].Value.ToString()),
                        Izd = int.Parse(row.Cells["Izd"].Value.ToString()),

                        Switch = row.Cells["断路器"].Value.ToString(),
                        Cable1 = row.Cells["电缆类型"].Value?.ToString() ?? "", // 使用 null 合并运算符,
                        Cable2 = row.Cells["电缆截面"].Value?.ToString() ?? "",
                        Transformer = row.Cells["互感器"].Value?.ToString() ?? "",

                    }; //end of fas

                    //填入消防、切非、计量信息
                    //是否消防负荷
                    if ((bool)row.Cells["消防"].Value == true)
                    {
                        chuxian.XiaoFang = true;
                    }
                    else
                    {
                        chuxian.XiaoFang = false;
                    } //end of 消防
                    //是否切非
                    if ((bool)row.Cells["切非"].Value == true)
                    {
                        chuxian.QieFei = true;
                    }
                    else
                    {
                        chuxian.QieFei = false;
                    } //end of 切非
                    //是否计量
                    if ((bool)row.Cells["计量"].Value == true)
                    {
                        chuxian.JiLiang = true;
                    }
                    else
                    {
                        chuxian.JiLiang = false;
                    } //end of 计量


                    list_Xuanzhong.Add(chuxian);

                } //end of if
            } //end of for

            return list_Xuanzhong;
        } //end of private List<DY_ChuXian> calculate_uiDataGridview()

        /// <summary>
        /// 一级子函数：对馈出柜路、进线模块的dataGridView进行计算
        /// </summary>
        private void calDG_JinXian_Chuxian(UIDataGridView uiDataGridView1, UITextBox uiTextBox_Pe, UITextBox uiTextBox_Kx, UITextBox uiTextBox_Cos, UITextBox uiTextBox6, UITextBox uiTextBox_Izd)
        {
            //进线柜Pe
            decimal sumPe = 0;
            decimal sumPe_XiaoFang = 0;
            decimal sumPe_FeiXiaoFang = 0;

            //循环计算馈出回路对应的datagridview中每一行的数据
            foreach (DataGridViewRow row in uiDataGridView1.Rows)
            {
                if ((bool)row.Cells["消防"].Value == true)
                {
                    sumPe_XiaoFang = sumPe_XiaoFang + cal_row(row);
                }
                else
                {
                    sumPe_FeiXiaoFang = sumPe_FeiXiaoFang + cal_row(row);
                }
                //sumPe = sumPe + cal_row(row);
            } //end of foreach

            //安装功率负荷,消防、非消防取最大值
            sumPe = sumPe_FeiXiaoFang;
            if (sumPe_FeiXiaoFang < sumPe_XiaoFang)
            {
                sumPe = sumPe_XiaoFang;
            }

            //重新设置进线柜的Pe
            uiTextBox_Pe.Text = sumPe.ToString();
            //选中进线柜时才执行
            if ((uiTextBox6.Text == "获取成功"))
            {
                //求进线柜的计算电流
                //仅当Pe有设置时才进行
                if (decimal.TryParse(uiTextBox_Pe.Text.ToString(), out decimal Pe))
                {
                    decimal Kx = decimal.TryParse(uiTextBox_Kx.Text.ToString(), out decimal tempKx) ? tempKx : (decimal)0.9;
                    decimal Cos = decimal.TryParse(uiTextBox_Cos.Text.ToString(), out decimal tempCos) ? tempCos : (decimal)0.85;
                    decimal G3 = (decimal)Math.Sqrt(3);
                    decimal Ijs = Pe * Kx / Cos / (decimal)0.38 / G3;

                    //在textbox中显示Ijs
                    uiTextBox_Ijs.Text = Math.Round(Ijs, 2).ToString();

                    //校验是否使用Izd
                    if (int.TryParse(uiTextBox_Izd.Text.ToString(), out int Izd) && (Izd > Ijs * (decimal)1.25))
                    {
                        uiTextBox_Izd.Text = Izd.ToString();
                    } //end of if (decimal.TryParse(uiTextBox_Pe.Text.ToString(), out decimal Izd) && (Izd > Ijs*(decimal)1.25) )
                    //当Izd不合法时，重新计算Izd
                    else
                    {
                        Izd = calCurrent_Izd(Ijs);
                        uiTextBox_Izd.Text = Izd.ToString();
                    }
                } //end of if (decimal.TryParse(uiTextBox_Pe.Text.ToString(), out decimal Pe))


            } //end of ((uiTextBox6.Text == "获取成功"))
        } //end of private void cal_JinXian_Chuxian()

        #endregion

        #region 按钮事件
        private void FPDY1_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 事件：点击uiDataGridView1
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //此列为计算按钮时
            if (e.ColumnIndex == uiDataGridView1.Columns["计算"].Index && e.RowIndex >= 0)
            {
                //当前行转化为DataGridViewRow,便于统一调用
                DataGridViewRow row = uiDataGridView1.Rows[e.RowIndex];
                //由子函数完成当前行的计算
                cal_row(row);

            } //end of if (e.ColumnIndex == uiDataGridView1.Columns["计算"].Index && e.RowIndex >= 0)
        } //end of private void uiDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)

        #endregion


        /// <summary>
        /// 按钮事件：选择图块
        /// 选择前需清空uiDataGridView1防止重新选择时错误
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton1_Click(object sender, EventArgs e)
        {
            //清空uiDataGridView1防止重新选择时错误
            uiDataGridView1.DataSource = null;
            //删除计算按钮列
            while (uiDataGridView1.Columns.Count > 1)
            {
                uiDataGridView1.Columns.RemoveAt(1); // 移除索引为1的列，即第二列
            }
            //进线柜信息清空
            uiTextBox_Pe.Text = string.Empty;
            uiTextBox_Ijs.Text = string.Empty;
            uiTextBox_Name.Text = string.Empty;
            uiTextBox_Izd.Text = string.Empty;
            uiTextBox6.Text = string.Empty;
            //选择图块
            //此函数将自动设置datagridview信息
            usr_selectBlks();
        }

        /// <summary>
        /// 按钮事件：计算按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiB_calculate_Click(object sender, EventArgs e)
        {
            //计算功能移入一级子函数，便于多按钮重复调用
            //对馈线模块、进线模块进行计算
            calDG_JinXian_Chuxian(uiDataGridView1, uiTextBox_Pe, uiTextBox_Kx, uiTextBox_Cos, uiTextBox6, uiTextBox_Izd);
        } //end of private void uiB_calculate_Click(object sender, EventArgs e)


        /// <summary>
        /// 按钮事件：计算并修改
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uiButton3_Click(object sender, EventArgs e)
        {
            //对馈线模块、进线模块进行计算
            calDG_JinXian_Chuxian(uiDataGridView1, uiTextBox_Pe, uiTextBox_Kx, uiTextBox_Cos, uiTextBox6, uiTextBox_Izd);

            //启用事务处理
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {

                //一、馈出回路处理：
                //1.调用一级datagridview的行转为List<DY_ChuXian>
                List<DY_ChuXian> list_Xuanzhong = calDG_ToChuXianClass();
                //2.新建str 存放可见性
                string strVisibility;
                //3.使用for循环对每行数据进行处理
                for (int i = 0; i < list_Xuanzhong.Count; i++)
                {
                    //获取当前行对应的class
                    DY_ChuXian row_Insert = list_Xuanzhong[i];
                    //4.对选择集中的OBJID进行数据处理
                    //无法将string形式的OBJId转回ObjectId
                    //因此使用foreach循环，查找List<DY_ChuXian>[i]对应的ObjectId
                    foreach (ObjectId id in id_ss)
                    {
                        //按list_Xuanzhong的顺序进行数据处理
                        //所以应找到与第i个list对应的ObjectId
                        if (id.ToString() == row_Insert.BlkObjID)
                        {
                            //5.根据ObjID转化为块记录
                            BlockReference blkRef = (BlockReference)trans.GetObject(id, OpenMode.ForRead);
                            //6.修改块属性
                            Dictionary<string, string> atts = new Dictionary<string, string>
                                {
                                    { "用途", row_Insert.Purpose },
                                    { "备注", row_Insert.Name },

                                    { "需要系数",  row_Insert.Kx.ToString() },
                                    { "功率因数",  row_Insert.Cos.ToString() },
                                    //常备用待补充:"备注2"
                                    //{ "计算电流",  row_Insert.Ijs.ToString() },
                                    { "整定电流",  row_Insert.Izd.ToString() },
                                    { "断路器",  row_Insert.Switch.ToString() },
                                    { "回路编号",  row_Insert.HuiluBianhao },
                                    { "电缆类型",  row_Insert.Cable1 },
                                    { "电缆规格",  row_Insert.Cable2 },
                                }; //end of atts

                            //备用回路的Pe、Ijs置为空
                            if (row_Insert.Purpose == "备用" || row_Insert.Purpose == "预留")
                            {
                                atts.Add("安装功率", "");
                                atts.Add("计算电流", "");
                            }
                            else
                            {
                                atts.Add("安装功率", row_Insert.Pe.ToString());
                                atts.Add("计算电流", row_Insert.Ijs.ToString());
                            }

                            //7.消防、切非、计量的设置
                            //是否消防负荷
                            if (row_Insert.XiaoFang)
                            {
                                atts.Add("是否消防负荷", "是");
                            } //end of if (row_Insert.XiaoFang)
                            else
                            {
                                atts.Add("是否消防负荷", "否");
                            } //end of else

                            //切非、计量
                            if (row_Insert.QieFei && row_Insert.JiLiang)
                            {
                                strVisibility = "切非计量";
                                atts.Add("互感器", row_Insert.Transformer);
                                atts.Add("是否切非", "是");
                                atts.Add("是否计量", "是");
                            }
                            else if (row_Insert.QieFei && !row_Insert.JiLiang)
                            {
                                strVisibility = "切非不计量";
                                atts.Add("是否切非", "是");
                                atts.Add("是否计量", "否");
                            }
                            else if (!row_Insert.QieFei && !row_Insert.JiLiang)
                            {
                                strVisibility = "不切非不计量";
                                atts.Add("是否切非", "否");
                                atts.Add("是否计量", "否");
                            }
                            else if (!row_Insert.QieFei && row_Insert.JiLiang)
                            {
                                strVisibility = "不切非计量";
                                atts.Add("互感器", row_Insert.Transformer);
                                atts.Add("是否切非", "否");
                                atts.Add("是否计量", "是");
                            }
                            else
                            {
                                MessageBox.Show("切非设置故障，请复核");
                                strVisibility = "不切非不计量";
                                atts.Add("是否切非", "否");
                                atts.Add("是否计量", "否");
                            }

                            //9.更新块的可见性
                            id.SetDynBlockValue("可见性1", strVisibility);
                            //10.更新块属性
                            id.UpdateAttributesInBlock(atts);

                        } //end of if (id.ToString() == row_Insert.BlkObjID)
                    } //end of foreach
                } //end of for

                //二、进线模块修改
                //1.判断进线模块是否存在
                //校验条件:uiTextBox6是否显示 ---- "获取成功"
                //新建ObjectId用于存放进线模块的id
                //校验结果存为id_JinXian，如果有效即有合法输入，缩短if
                ObjectId id_JinXian = new ObjectId();
                int Izd = 0;
                if ((uiTextBox6.Text == "获取成功") && (int.TryParse(uiTextBox_Izd.Text.ToString(), out Izd)))
                {
                    //2.从id_ss中找到进线模块对应的objectId

                    foreach (ObjectId id_2 in id_ss)
                    {
                        if (id_2.ToString() == uiTextBox_OBJID.Text)
                        {
                            id_JinXian = id_2;
                        }
                    } //end of foreach
                } //end of ((uiTextBox6.Text == "获取成功"))

                //3.根据校验结果决定是否开始进线模块处理主函数
                if (id_JinXian.IsValid)
                {
                    //4.打开块记录
                    BlockReference blkRef_JinXian = (BlockReference)trans.GetObject(id_JinXian, OpenMode.ForRead);
                    //5.新建字典，将值填入块属性
                    Dictionary<string, string> atts_JinXian = new Dictionary<string, string>
                                {
                                    { "配电箱编号", uiTextBox_Name.Text },
                                    { "安装功率",uiTextBox_Pe.Text },
                                    { "需要系数",  uiTextBox_Kx.Text },
                                    { "功率因数",  uiTextBox_Cos.Text },
                                    //{ "互感器",  uiTextBox_Cos.ToString() },
                                    //{ "断路器",  row_Insert.Switch.ToString() },

                                    { "计算电流",  uiTextBox_Ijs.Text },
                                    { "整定电流",  uiTextBox_Izd.Text },

                                };

                    //6.根据Izd和形式整定互感器、电流互感器
                    //当Izd>0时，从云数据库中查询
                    if (Izd > 0)
                    {
                        //添加互感器信息
                        CABLE_CLASS result = USRSQLTools.query_Cable(conn, Izd);
                        atts_JinXian.Add("互感器", result.CT);

                        //7.根据radio按钮确定开关形式
                        string Switch = result.MCB + "/MX+OF";
                        string Switch_Vis = "断路器";
                        switch (uiRadioButtonGroup1.SelectedIndex)
                        {
                            case 0:
                                Switch = result.INS;
                                Switch_Vis = "负荷开关";
                                break;
                            case 2:
                                Switch = result.FU;
                                Switch_Vis = "熔断器";
                                break;
                        }
                        //添加断路器信息
                        atts_JinXian.Add("断路器", Switch);
                        //修改可见性：进线开关的类型
                        id_JinXian.SetDynBlockValue("可见性1", Switch_Vis);
                    } // end of if (Izd > 0)
                    
                    id_JinXian.UpdateAttributesInBlock(atts_JinXian);
                } //end of (id_JinXian.IsValid) 主函数结束

                //提交事务处理
                trans.Commit();


                //关闭窗体
                form_this.Close();
                //Application.Exit();

                //this.Close();
                //弹窗提示：修改完成
                Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("系统图修改已完成");
            } //end of using trans
        } //end of private void uiButton3_Click(object sender, EventArgs e)

        private void uiRadioButtonGroup1_ValueChanged(object sender, int index, string text)
        {

        }

        private void uiButton2_Click(object sender, EventArgs e)
        {

        }
    } //end of class
} //end of namespace
