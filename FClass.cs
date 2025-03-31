using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using Sunny.UI.Win32;

namespace DYPD_2
{
   
    public partial class FClass
    {

        #region 低压柜负荷计算
        /// <summary>
        /// 类定义：低压柜，每个回路均为一个class
        /// </summary>
        public class DY_ChuXian
        {
            public string BlkName { get; set; } //块名名称
            public string BlkObjID { get; set; } //块的OBJID
            public Point3d BlkPoint { get; set; } //块位置

            public bool XiaoFang { get; set; } = false; //是否为消防负荷
            public bool QieFei { get; set; } = false; //是否切非
            public bool JiLiang { get; set; } = false; //是否计量
            public string ChangBei { get; set; } //常备用选择

            public string Purpose { get; set; } //用途（中文）
            public string Name { get; set; } //箱名
            public string HuiluBianhao { get; set; } //回路编号
            

            public decimal Pe { get; set; } //安装功率

            public decimal Kx { get; set; } //需要系数
            public decimal Cos { get; set; } //功率因数


            private decimal Ijs_pri { get; set; } //私有字段，用于给计算电流取小数点后一位
            public decimal Ijs
            {
                get { return Ijs_pri; }
                set
                {
                    // 使用 Math.Round 将值舍入到小数点后一位
                    Ijs_pri = Math.Round(value, 1);
                }
            } //计算电流（小数点后一位）

            public int Izd { get; set; } //整定电流
            public string Switch { get; set; } //断路器 待修改
            
            public string Cable1 { get; set; } //电缆类型
            public string Cable2 { get; set; } //电缆规格
            
            public string Transformer { get; set; } //互感器

        } //end of DY_ChuXian

        //进线模块
        public class DY_JinXian
        {
            public string Name { get; set; } //箱名
            public decimal Pe { get; set; } //安装功率
            public decimal Kx { get; set; } //需要系数
            public decimal Cos { get; set; } //功率因数
            private decimal Ijs_pri { get; set; } //私有字段，用于给计算电流取小数点后一位
            public decimal Ijs
            {
                get { return Ijs_pri; }
                set
                {
                    // 使用 Math.Round 将值舍入到小数点后一位
                    Ijs_pri = Math.Round(value, 1);
                }
            } //计算电流（小数点后一位）

            public int Izd { get; set; } //整定电流
            public string Switch { get; set; } //断路器 待修改
            public string Transformer { get; set; } //互感器
            public bool XiaoFang { get; set; } //是否带消防负荷

        } //end of DY_ChuXian


        /// <summary>
        /// 定义类：存放MYSQL查询电缆表的返回值
        /// </summary>
        public class CABLE_CLASS
        {
            public int Izd { get; set; } //整定电流
            public string Cable { get; set; } //电缆截面(相线)
            public string CablePE { get; set; } //电缆截面(PE)
            public string SC { get; set; } //穿管管径
            public string MCB { get; set; } //断路器
            public string INS { get; set; } //负荷开关
            public string FU { get; set; } //熔断器
            public string CT { get; set; } //电流互感器

        } //end of CABLE_CLASS

        /// <summary>
        /// 定义类：存放MYSQL查询二次配塑壳断路器的返回值
        /// </summary>
        public class MCB_CLASS
        {
            public int Izd { get; set; } //整定电流
            public string Info1 { get; set; } //断路器壳架+In
            public string Info2 { get; set; } //3P
            public string Info3 { get; set; } //消防负荷MA
            public string Info4 { get; set; } //切非回路MX+OF

        } //end of CABLE_CLASS

        #endregion

        #region 所有低压柜回路校验
        /// <summary>
        /// 用于DYPD2程序：记录低压柜的馈出回路
        /// </summary>
        public class DiYaGui_CLASS
        {
            public string Purpose { get; set; } //用途
            public string Pe { get; set; } //安装功率
            public ObjectId OBJID { get; set; } //块对应的OBJID
            public string ChangBeiYong { get; set; } //常备用 对应属性为 备注2
            public string HuiLu { get; set; } //回路编号
            public string BianHao { get; set; } //箱柜编号 不一定存在
        } //end of CABLE_CLASS

        /// <summary>
        /// 用于DYPD2程序：记录低压柜的竖向干线回路
        /// </summary>
        public class DiYaGanXian_CLASS
        {
            public string Purpose { get; set; } //用途
            public string Pe { get; set; } //安装功率
            public string BianHao { get; set; } //箱柜
            public bool FangShe { get; set; } = false; //是否为放射式回路
            public string BlkVis { get; set; } = "放射式单回路"; //单、双回路对应的可见性
            public string HuiLu { get; set; } = ""; //单回路的回路编号
            public string HuiLu1 { get; set; } = "";//常用回路编号
            public string HuiLu2 { get; set; } = "";//备用回路编号
            public ObjectId OBJID { get; set; } //OBJID1
            //public ObjectId OBJD2 { get; set; } //OBJID2

        } //end of DiYaGanXian_CLASS

        #endregion
    }
}
