using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;

using Sunny.UI;

using MySql.Data;
using MySql.Data.MySqlClient;
using Autodesk.AutoCAD.Geometry;
using static DYPD_2.FClass;
using Google.Protobuf.WellKnownTypes;
using MySqlX.XDevAPI.Relational;

namespace DYPD_2
{
    public partial class FMain : UIForm
    {
        //环境初始化
        Database db;
        Editor ed;
        MySqlConnection conn;
        string tablename;
        ObjectId[] id_ss;

        public FMain()
        {

            //窗口初始化程序
            InitializeComponent();

            //将窗体传至子页面，便于关闭窗体
            FPDY1 FPage1 = new FPDY1(this);
            FPDY2 FPage2 = new FPDY2();
            FPage1.Text = "低压系统图";
            FPage2.Text = "干线系统图";
            uiTabControl1.AddPage(FPage1);
            uiTabControl1.AddPage(FPage2);


        }

        


        private void FMain_Load(object sender, EventArgs e)
        {

        }


        

    } //end of class
} //end of namespace
