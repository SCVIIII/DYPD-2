﻿using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;


using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;






namespace DYPD_2

{
    public static  class ARXTools
    {

        /// <summary>
        /// 获取用户插入点,需在using trans中使用
        /// </summary>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static Point3d get_POINT3D(Document doc)
        {
            //获取用户输入点
            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("")
            {
                Message = "\n选择低压柜的插入点: "
            };
            pPtRes = doc.Editor.GetPoint(pPtOpts);
            Point3d ptReturn = pPtRes.Value;
            return ptReturn;
        }

        /// <summary>
        /// 动态块的动态属性类型
        /// </summary>
        public enum DynBlockPropTypeCode
        {
            /// <summary>
            /// 字符串
            /// </summary>
            String = 1,
            /// <summary>
            /// 实数
            /// </summary>
            Real = 40,
            /// <summary>
            /// 短整型
            /// </summary>
            Short = 70,
            /// <summary>
            /// 长整型
            /// </summary>
            Long = 90
        }


        /// <summary>
        /// 获得动态块的所有动态属性
        /// </summary>
        /// <param name="blockId">动态块的Id</param>
        /// <returns>返回动态块的所有属性</returns>
        public static DynamicBlockReferencePropertyCollection GetDynProperties(this ObjectId blockId)
        {
            //获取块参照
            BlockReference br = blockId.GetObject(OpenMode.ForRead) as BlockReference;
            //如果不是动态块，则返回
            if (br == null && !br.IsDynamicBlock) return null;
            //返回动态块的动态属性
            return br.DynamicBlockReferencePropertyCollection;
        }

        /// <summary>
        /// 获取动态块的动态属性值
        /// </summary>
        /// <param name="blockId">动态块的Id</param>
        /// <param name="propName">需要查找的动态属性名</param>
        /// <returns>返回指定动态属性的值</returns>
        public static string GetDynBlockValue(this ObjectId blockId, string propName)
        {
            string propValue = null;//用于返回动态属性值的变量
            var props = blockId.GetDynProperties();//获得动态块的所有动态属性
            //遍历动态属性
            foreach (DynamicBlockReferenceProperty prop in props)
            {
                //如果动态属性的名称与输入的名称相同
                if (prop.PropertyName == propName)
                {
                    //获取动态属性值并结束遍历
                    propValue = prop.Value.ToString();
                    break;
                }
            }
            return propValue;//返回动态属性值
        }

        /// <summary>
        /// 设置动态块的动态属性
        /// </summary>
        /// <param name="blockId">动态块的ObjectId</param>
        /// <param name="propName">动态属性的名称</param>
        /// <param name="value">动态属性的值</param>
        public static void SetDynBlockValue(this ObjectId blockId, string propName, object value)
        {
            var props = blockId.GetDynProperties();//获得动态块的所有动态属性
            //遍历动态属性
            foreach (DynamicBlockReferenceProperty prop in props)
            {
                //如果动态属性的名称与输入的名称相同且为可读
                if (prop.ReadOnly == false && prop.PropertyName == propName)
                {
                    //判断动态属性的类型并通过类型转化设置正确的动态属性值
                    switch (prop.PropertyTypeCode)
                    {
                        case (short)DynBlockPropTypeCode.Short://短整型
                            prop.Value = Convert.ToInt16(value);
                            break;
                        case (short)DynBlockPropTypeCode.Long://长整型
                            prop.Value = Convert.ToInt64(value);
                            break;
                        case (short)DynBlockPropTypeCode.Real://实型
                            prop.Value = Convert.ToDouble(value);
                            break;
                        default://其它
                            prop.Value = value;
                            break;
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// 将string 转换为Point3d
        /// </summary>
        /// <param name="pointString">此处输入应为string (X,Y,Z)形式</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static Point3d StringToPoint3d(string pointString)
        {
            // 去除括号，并按逗号分隔字符串获取 x、y 和 z 坐标值
            string[] coordinates = pointString.Trim('(', ')').Split(',');

            if (coordinates.Length != 3)
            {
                throw new ArgumentException("Invalid point string format.");
            }

            // 将坐标值转换为 double 类型
            double x = double.Parse(coordinates[0]);
            double y = double.Parse(coordinates[1]);
            double z = double.Parse(coordinates[2]);

            // 创建 Point3d 对象
            Point3d point = new Point3d(x, y, z);

            return point;
        } //end of public static Point3d StringToPoint3d(string pointString)


        //可见性
        //public static void ToggleDynamicBlockVisibility()
        //{
        //    Document doc = Application.DocumentManager.MdiActiveDocument;
        //    Database db = doc.Database;
        //    Editor editor = doc.Editor;

        //    PromptResult blockNameResult = editor.GetString("输入要切换可见性的动态块的名称：");

        //    if (blockNameResult.Status == PromptStatus.OK)
        //    {
        //        string blockName = blockNameResult.StringResult;

        //        using (Transaction tr = db.TransactionManager.StartTransaction())
        //        {
        //            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        //            if (bt.Has(blockName))
        //            {
        //                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[blockName], OpenMode.ForRead);

        //                if (btr.IsDynamicBlock)
        //                {
        //                    //btr.GetBlockReferenceIds
        //                    // 获取动态块定义
        //                    BlockTableRecord dynBlockDef = (BlockTableRecord)tr.GetObject(btr.GetBlockReferenceId(0), OpenMode.ForWrite);

        //                    //cTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

        //                    // 获取动态块的属性定义
        //                    foreach (ObjectId attId in dynBlockDef)
        //                    {
        //                        AttributeReference attRef = tr.GetObject(attId, OpenMode.ForRead) as AttributeReference;
        //                        if (attRef != null && attRef.Tag == "Visibility")
        //                        {
        //                            // 切换可见性状态
        //                            attRef.UpgradeOpen();
        //                            attRef.TextString = (attRef.TextString == "1") ? "0" : "1";
        //                        }
        //                    }

        //                    tr.Commit();
        //                }
        //                else
        //                {
        //                    editor.WriteMessage($"'{blockName}' 不是动态块。");
        //                }
        //            }
        //            else
        //            {
        //                editor.WriteMessage($"未找到名称为 '{blockName}' 的块。");
        //            }
        //        }
        //    }
        //}
        //end of 可见性


        /// <summary>
        /// 修改可见性
        /// </summary>
        /// <param name="blockRefId">对应图块的objectid</param>
        /// <param name="db"></param>
        /// <param name="DynName">可见性的名称</param>
        /// <param name="DynValue">可见性的状态</param>
        public static void UpdateDynInBlock(Transaction trans, ObjectId blockRefId,  string DynName,string DynValue)
        {

            //转化为BlockReference
            BlockReference blkRef = (BlockReference)trans.GetObject(blockRefId, OpenMode.ForRead);
            // 获取动态块参照的 DynamicBlockReferencePropertyCollection 对象
            DynamicBlockReferencePropertyCollection props = blkRef.DynamicBlockReferencePropertyCollection;
            //测试语句:使用第一个可见性
            DynamicBlockReferenceProperty prop = props[0];
            prop.Value = "设置剩余电流监控";
            //foreach (DynamicBlockReferenceProperty prop in props)
            //{
            //    if (prop.PropertyName.Equals(DynName, StringComparison.OrdinalIgnoreCase))
            //    {
            //        // 替换 "YOUR_PROPERTY_NAME" 为实际的动态块属性名称
            //        // 通常情况下，"设置" 可见性状态对应着属性值为 1 或 true
            //        // 请根据动态块定义中的属性值来设置对应的状态值
            //        prop.Value = DynValue;
            //        break;
            //    }
            //} //end of foreach

            //props.

            //BlockTableRecord currentSpace = acTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

            //currentSpace.AppendEntity(blkRef);
            //tr.AddNewlyCreatedDBObject(blockRef, true);

            //// 获取动态块参照的 DynamicBlockReferencePropertyCollection 对象
            //DynamicBlockReferencePropertyCollection props = blockRef.DynamicBlockReferencePropertyCollection;

            //BlockTableRecord space = (BlockTableRecord)spaceId.GetObject(OpenMode.ForWrite);
            //BlockReference br = new BlockReference(position, bt[blockName]);
            //br.ScaleFactors = scale;//设置块参照的缩放比例
            //br.Layer = layer;//设置块参照的层名
            //br.Rotation = rotateAngle;//设置块参照的旋转角度
            //ObjectId btrId = bt[blockName];//获取块表记录的Id
            ////打开块表记录
            //BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            ////添加可缩放性支持
            //if (record.Annotative == AnnotativeStates.True)
            //{
            //    ObjectContextCollection contextCollection = db.ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES");
            //    ObjectContexts.AddContext(br, contextCollection.GetContext("1:1"));
            //}
            //blockRefId = space.AppendEntity(br);//在空间中加入创建的块参照
            //db.TransactionManager.AddNewlyCreatedDBObject(br, true);//通知事务处理加入创建的块参照
            //space.DowngradeOpen();//为了安全，将块表状态改为读
            //return blockRefId;//返回添加的块参照的Id

            ////设置可见性
            ////循环查找所有可见性,直至找到对应的可见性名称
            //foreach (DynamicBlockReferenceProperty prop in props)
            //{
            //    if (prop.PropertyName.Equals(DynName, StringComparison.OrdinalIgnoreCase))
            //    {
            //        // 替换 "YOUR_PROPERTY_NAME" 为实际的动态块属性名称
            //        // 通常情况下，"设置" 可见性状态对应着属性值为 1 或 true
            //        // 请根据动态块定义中的属性值来设置对应的状态值
            //        prop.Value = DynValue;
            //        break;
            //    }
            //} //end of foreach
        } //end of UpdateDynInBlock


        /// <summary>
        /// 在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">插入点</param>
        /// <param name="scale">缩放比例</param>
        /// <param name="rotateAngle">旋转角度</param>
        /// <returns>返回块参照的Id</returns>
        public static ObjectId InsertBlockReference(this ObjectId spaceId, string layer, string blockName, Point3d position, Scale3d scale, double rotateAngle)
        {
            ObjectId blockRefId;//存储要插入的块参照的Id
            Database db = spaceId.Database;//获取数据库对象
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //如果没有blockName表示的块，则程序返回
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)spaceId.GetObject(OpenMode.ForWrite);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName]);
            br.ScaleFactors = scale;//设置块参照的缩放比例
            br.Layer = layer;//设置块参照的层名
            br.Rotation = rotateAngle;//设置块参照的旋转角度
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //添加可缩放性支持
            if (record.Annotative == AnnotativeStates.True)
            {
                ObjectContextCollection contextCollection = db.ObjectContextManager.GetContextCollection("ACDB_ANNOTATIONSCALES");
                ObjectContexts.AddContext(br, contextCollection.GetContext("1:1"));
            }
            blockRefId = space.AppendEntity(br);//在空间中加入创建的块参照
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);//通知事务处理加入创建的块参照
            space.DowngradeOpen();//为了安全，将块表状态改为读
            return blockRefId;//返回添加的块参照的Id
        }
        /// <summary>
        /// 在AutoCAD图形中插入块参照
        /// </summary>
        /// <param name="spaceId">块参照要加入的模型空间或图纸空间的Id</param>
        /// <param name="layer">块参照要加入的图层名</param>
        /// <param name="blockName">块参照所属的块名</param>
        /// <param name="position">插入点</param>
        /// <param name="scale">缩放比例</param>
        /// <param name="rotateAngle">旋转角度</param>
        /// <param name="attNameValues">属性的名称与取值</param>
        /// <returns>返回块参照的Id</returns>
        public static ObjectId InsertBlockReference(this ObjectId spaceId, string layer, string blockName, Point3d position, Scale3d scale, double rotateAngle, Dictionary<string, string> attNameValues)
        {
            Database db = spaceId.Database;//获取数据库对象
            //以读的方式打开块表
            BlockTable bt = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
            //如果没有blockName表示的块，则程序返回
            if (!bt.Has(blockName)) return ObjectId.Null;
            //以写的方式打开空间（模型空间或图纸空间）
            BlockTableRecord space = (BlockTableRecord)spaceId.GetObject(OpenMode.ForWrite);
            ObjectId btrId = bt[blockName];//获取块表记录的Id
            //打开块表记录
            BlockTableRecord record = (BlockTableRecord)btrId.GetObject(OpenMode.ForRead);
            //创建一个块参照并设置插入点
            BlockReference br = new BlockReference(position, bt[blockName])
            {
                ScaleFactors = scale,//设置块参照的缩放比例
                Layer = layer,//设置块参照的层名
                Rotation = rotateAngle//设置块参照的旋转角度
            };
            space.AppendEntity(br);//为了安全，将块表状态改为读 
            //判断块表记录是否包含属性定义
            if (record.HasAttributeDefinitions)
            {
                //若包含属性定义，则遍历属性定义
                foreach (ObjectId id in record)
                {
                    //检查是否是属性定义
                    AttributeDefinition attDef = id.GetObject(OpenMode.ForRead) as AttributeDefinition;
                    if (attDef != null)
                    {
                        //创建一个新的属性对象
                        AttributeReference attribute = new AttributeReference();
                        //从属性定义获得属性对象的对象特性
                        attribute.SetAttributeFromBlock(attDef, br.BlockTransform);
                        //设置属性对象的其它特性
                        attribute.Position = attDef.Position.TransformBy(br.BlockTransform);
                        attribute.Rotation = attDef.Rotation;
                        attribute.AdjustAlignment(db);
                        //判断是否包含指定的属性名称
                        if (attNameValues.ContainsKey(attDef.Tag.ToUpper()))
                        {
                            //设置属性值
                            attribute.TextString = attNameValues[attDef.Tag.ToUpper()].ToString();
                        }
                        //向块参照添加属性对象
                        br.AttributeCollection.AppendAttribute(attribute);
                        db.TransactionManager.AddNewlyCreatedDBObject(attribute, true);
                    }
                }
            }
            db.TransactionManager.AddNewlyCreatedDBObject(br, true);
            return br.ObjectId;//返回添加的块参照的Id
        }

        /// <summary>
        /// 更新块参照中的属性值
        /// </summary>
        /// <param name="blockRefId">块参照的Id</param>
        /// <param name="attNameValues">需要更新的属性名称与取值</param>
        public static void UpdateAttributesInBlock(this ObjectId blockRefId, Dictionary<string, string> attNameValues)
        {

            //获取块参照对象
            BlockReference blockRef = blockRefId.GetObject(OpenMode.ForRead) as BlockReference;
            if (blockRef != null)
            {
                //遍历块参照中的属性
                foreach (ObjectId id in blockRef.AttributeCollection)
                {
                    //获取属性
                    AttributeReference attref = id.GetObject(OpenMode.ForRead) as AttributeReference;
                    //判断是否包含指定的属性名称
                    if (attNameValues.ContainsKey(attref.Tag.ToUpper()))
                    {
                        attref.UpgradeOpen();//切换属性对象为写的状态
                        //设置属性值
                        attref.TextString = attNameValues[attref.Tag.ToUpper()].ToString();
                        attref.DowngradeOpen();//为了安全，将属性对象的状态改为读
                    }
                }
            }
        }

        /// <summary>
        /// 获取块的插入点
        /// </summary>
        /// <param name="blockReferenceId"></param>
        /// <returns></returns>
        public static Point3d GetInsertPoint(this ObjectId blockReferenceId)
        {
            // 获取当前文档和数据库
            Database db = blockReferenceId.Database;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //打开块参照
                BlockReference bref = (BlockReference)trans.GetObject(blockReferenceId, OpenMode.ForRead);
                //返回块插入点
                return bref.Position;
                trans.Commit();
            }
        } // end of public static Point3d GetInsertPoint

        /// <summary>
        /// 获取指定名称的块属性值
        /// </summary>
        /// <param name="blockReferenceId">块参照的Id</param>
        /// <param name="attributeName">属性名</param>
        /// <returns>返回指定名称的块属性值</returns>
        public static string GetAttributeInBlockReference(this ObjectId blockReferenceId, string attributeName)
        {
            string attributeValue = string.Empty; // 属性值
            Database db = blockReferenceId.Database;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 获取块参照
                BlockReference bref = (BlockReference)trans.GetObject(blockReferenceId, OpenMode.ForRead);
                // 遍历块参照的属性
                foreach (ObjectId attId in bref.AttributeCollection)
                {
                    // 获取块参照属性对象
                    AttributeReference attRef = (AttributeReference)trans.GetObject(attId, OpenMode.ForRead);
                    //判断属性名是否为指定的属性名
                    if (attRef.Tag.ToUpper() == attributeName.ToUpper())
                    {
                        attributeValue = attRef.TextString;//获取属性值
                        break;
                    }
                }
                trans.Commit();
            }
            return attributeValue; //返回块属性值
        }






        /// <summary>
        /// 获取块参照的块名（包括动态块）
        /// </summary>
        /// <param name="id">块参照的Id</param>
        /// <returns>返回块名</returns>
        public static string GetBlockName(this ObjectId id)
        {
            //获取块参照
            BlockReference bref = id.GetObject(OpenMode.ForRead) as BlockReference;
            if (bref != null)//如果是块参照
                return GetBlockName(bref);
            else
                return null;
        }

        /// <summary>
        /// 获取块参照的块名（包括动态块）
        /// </summary>
        /// <param name="bref">块参照</param>
        /// <returns>返回块名</returns>
        public static string GetBlockName(this BlockReference bref)
        {
            string blockName;//存储块名
            if (bref == null) return null;//如果块参照不存在，则返回
            if (bref.IsDynamicBlock) //如果是动态块
            {
                //获取动态块所属的动态块表记录
                ObjectId idDyn = bref.DynamicBlockTableRecord;
                //打开动态块表记录
                BlockTableRecord btr = (BlockTableRecord)idDyn.GetObject(OpenMode.ForRead);
                blockName = btr.Name;//获取块名
            }
            else //非动态块
                blockName = bref.Name; //获取块名
            return blockName;//返回块名
        }

     



        /// <summary>
        /// 获取指定名称的 匿名块+普通块参照 的筛选信息
        /// </summary>
        /// <param name="BlkNamesA0">块名</param>
        /// <returns>返回按块名进行选择的筛选项(new TypedValue((int)DxfCode.BlockName, listTKblk)</returns>
     
        public static string GetAnonymousBlk(this Database db, List<string> listTKblk)
        {

            //Database db = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.WorkingDatabase;
            //返回值,用于存放匿名块+普通块 对应块名筛选项的string
            string BlkNamesA0;

            //查表,确定对应的BlockReference
            var BlkssListA0 = (from d in db.GetEntsInModelSpace<BlockReference>()
                               where
                               listTKblk.Contains(d.GetBlockName())
                               //d.GetBlockName() == "TJAD_TK_A0"
                               //select d.GetBlockName().Distinct().ToString();
                               //select d.BlockName.Distinct().ToString();
                               //select d.Name.Distinct().ToString();
                               select d.Name.ToString()).Distinct().ToList(); //返回去重的列表

            //列表元素的数量
            int numBlkssListA0 = BlkssListA0.Count;


            //如果选择集为空
            if (numBlkssListA0 < 1)
            {
                BlkNamesA0 = null;
            }


            else if (numBlkssListA0 == 1)  //如果选择集只有一个元素
            {

                if ((BlkssListA0[0].StartsWith("*U")) || (BlkssListA0[0].StartsWith("*u"))) //列表第一个元素为匿名块时
                {

                    BlkNamesA0 = "`" + BlkssListA0[0];  //匿名块,需添加前缀"`",AutoCAD中"*"为通配符
                }

                else
                {
                    BlkNamesA0 = BlkssListA0[0]; //普通块,直接返回块名
                }

            }


            else  //如果选择集有多个元素
            {

                if ((BlkssListA0[0].StartsWith("*U")) || (BlkssListA0[0].StartsWith("*u")))  //列表第一个元素为匿名块时
                {

                    BlkNamesA0 = "`" + BlkssListA0[0];
                }

                else
                {

                    BlkNamesA0 = BlkssListA0[0];  //普通块,直接返回块名
                }


                //添加筛选信息
                for (int i = 1; i < numBlkssListA0; i++)
                {
                    if ((BlkssListA0[i].StartsWith("*U")) || (BlkssListA0[i].StartsWith("*u")))
                        BlkNamesA0 = BlkNamesA0 + ",`" + BlkssListA0[i];

                    BlkNamesA0 = BlkNamesA0 + "," + BlkssListA0[i];
                }
            }


            //返回值参考值:"`U8,TJAD_TK_A0,A0+1／4L,`U32,A0+1／2L,A0+3／4L"
            return BlkNamesA0;

        }




        /// <summary>
        /// 获取与给定点指定角度和距离的点
        /// </summary>
        /// <param name="point">给定点</param>
        /// <param name="angle">角度</param>
        /// <param name="dist">距离</param>
        /// <returns>返回与给定点指定角度和距离的点</returns>
        public static Point3d PolarPoint(this Point3d point, double angle, double dist)
        {
            return new Point3d(point.X + dist * Math.Cos(angle), point.Y + dist * Math.Sin(angle), point.Z);
        }

    

        //图层操作类
        /// <summary>
        /// 创建新图层
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layerName">图层名</param>
        /// <returns>返回新建图层的ObjectId</returns>
        public static ObjectId AddLayer(this Database db, string layerName)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            if (!lt.Has(layerName))//如果不存在名为layerName的图层，则新建一个图层
            {
                //定义一个新的层表记录
                LayerTableRecord ltr = new LayerTableRecord
                {
                    Name = layerName//设置图层名
                };
                lt.UpgradeOpen();//切换层表的状态为写以添加新的图层
                //将层表记录的信息添加到层表中
                lt.Add(ltr);
                //把层表记录添加到事务处理中
                db.TransactionManager.AddNewlyCreatedDBObject(ltr, true);
                lt.DowngradeOpen();//为了安全，将层表的状态切换为读
            }
            return lt[layerName];//返回新添加的层表记录的ObjectId
        }

        /// <summary>
        /// 设置图层的颜色
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layerName">图层名</param>
        /// <param name="colorIndex">颜色索引</param>
        /// <returns>如果成功设置图层颜色，则返回true，否则返回false</returns>
        public static bool SetLayerColor(this Database db, string layerName, short colorIndex)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            //如果不存在名为layerName的图层，则返回
            if (!lt.Has(layerName)) return false;
            ObjectId layerId = lt[layerName];//获取名为layerName的层表记录的Id
            //以写的方式打开名为layerName的层表记录
            LayerTableRecord ltr = (LayerTableRecord)layerId.GetObject(OpenMode.ForWrite);
            //设置图层的颜色
            ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
            ltr.DowngradeOpen();//为了安全，将图层的状态切换为读
            return true;//设置图层颜色成功
        }

      

 


        

        


        /// <summary>
        /// 获取模型空间中类型为T的所有实体(对象打开为读）
        /// </summary>
        /// <typeparam name="T">实体的类型</typeparam>
        /// <param name="db">数据库对象</param>
        /// <returns>返回模型空间中类型为T的实体</returns>
        public static List<T> GetEntsInModelSpace<T>(this Database db) where T : Entity
        {
            return GetEntsInModelSpace<T>(db, OpenMode.ForRead, false);
        }

        /// <summary>
        /// 获取模型空间中类型为T的所有实体
        /// </summary>
        /// <typeparam name="T">实体的类型</typeparam>
        /// <param name="db">数据库对象</param>
        /// <param name="mode">实体打开方式</param>
        /// <param name="openErased">是否打开已删除的实体</param>
        /// <returns>返回模型空间中类型为T的实体</returns>
        public static List<T> GetEntsInModelSpace<T>(this Database db, OpenMode mode, bool openErased) where T : Entity
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            //声明一个List类的变量，用于返回类型为T为的实体列表
            List<T> ents = new List<T>();
            //获取类型T代表的DXF代码名用于构建选择集过滤器
            string dxfname = RXClass.GetClass(typeof(T)).DxfName;
            //构建选择集过滤器        
            TypedValue[] values = { new TypedValue((int)DxfCode.Start, dxfname),
                                    new TypedValue((int)DxfCode.LayoutName,"Model")};
            SelectionFilter filter = new SelectionFilter(values);
            //选择符合条件的所有实体
            PromptSelectionResult entSelected = ed.SelectAll(filter);
            if (entSelected.Status == PromptStatus.OK)
            {
                //循环遍历符合条件的实体
                foreach (var id in entSelected.Value.GetObjectIds())
                {
                    //将实体强制转化为T类型的对象
                    //不能将实体直接转化成泛型T，必须首先转换成object类
                    T t = (T)(object)id.GetObject(mode, openErased);
                    ents.Add(t);//将实体添加到返回列表中
                }
            }
            return ents;//返回类型为T为的实体列表
        }

   


    } // end of class


    /// <summary>
    /// TypedValue列表类，简化选择集过滤器的构造
    /// </summary>
    public class TypedValueList : List<TypedValue>
    {
        /// <summary>
        /// 接受可变参数的构造函数
        /// </summary>
        /// <param name="args">TypedValue对象</param>
        public TypedValueList(params TypedValue[] args)
        {
            AddRange(args);
        }

        



        /// <summary>
        /// TypedValueList隐式转换为SelectionFilter
        /// </summary>
        /// <param name="src">要转换的TypedValueList对象</param>
        /// <returns>返回对应的SelectionFilter类对象</returns>
        public static implicit operator SelectionFilter(TypedValueList src)
        {
            return src != null ? new SelectionFilter(src) : null;
        }

        /// <summary>
        /// TypedValueList隐式转换为ResultBuffer
        /// </summary>
        /// <param name="src">要转换的TypedValueList对象</param>
        /// <returns>返回对应的ResultBuffer对象</returns>
        public static implicit operator ResultBuffer(TypedValueList src)
        {
            return src != null ? new ResultBuffer(src) : null;
        }

        /// <summary>
        /// TypedValueList隐式转换为TypedValue数组
        /// </summary>
        /// <param name="src">要转换的TypedValueList对象</param>
        /// <returns>返回对应的TypedValue数组</returns>
        public static implicit operator TypedValue[](TypedValueList src)
        {
            return src != null ? src.ToArray() : null;
        }

        /// <summary>
        /// TypedValue数组隐式转换为TypedValueList
        /// </summary>
        /// <param name="src">要转换的TypedValue数组</param>
        /// <returns>返回对应的TypedValueList</returns>
        public static implicit operator TypedValueList(TypedValue[] src)
        {
            return src != null ? new TypedValueList(src) : null;
        }

        /// <summary>
        /// SelectionFilter隐式转换为TypedValueList
        /// </summary>
        /// <param name="src">要转换的SelectionFilter</param>
        /// <returns>返回对应的TypedValueList</returns>
        public static implicit operator TypedValueList(SelectionFilter src)
        {
            return src != null ? new TypedValueList(src.GetFilter()) : null;
        }

        /// <summary>
        /// ResultBuffer隐式转换为TypedValueList
        /// </summary>
        /// <param name="src">要转换的ResultBuffer</param>
        /// <returns>返回对应的TypedValueList</returns>
        public static implicit operator TypedValueList(ResultBuffer src)
        {
            return src != null ? new TypedValueList(src.AsArray()) : null;
        }

    }

    /// <summary>
    /// Point3d列表类
    /// </summary>
    public class Point3dList : List<Point3d>
    {
        /// <summary>
        /// 接受可变参数的构造函数
        /// </summary>
        /// <param name="args">Point3d类对象</param>
        public Point3dList(params Point3d[] args)
        {
            AddRange(args);
        }

        /// <summary>
        /// Point3dList隐式转换为Point3d数组
        /// </summary>
        /// <param name="src">要转换的Point3dList对象</param>
        /// <returns>返回对应的Point3d数组</returns>
        public static implicit operator Point3d[](Point3dList src)
        {
            return src != null ? src.ToArray() : null;
        }

        /// <summary>
        /// Point3dList隐式转换为Point3dCollection
        /// </summary>
        /// <param name="src">要转换的Point3dList对象</param>
        /// <returns>返回对应的Point3dCollection</returns>
        public static implicit operator Point3dCollection(Point3dList src)
        {
            return src != null ? new Point3dCollection(src) : null;
        }

        /// <summary>
        /// Point3d数组隐式转换为Point3dList
        /// </summary>
        /// <param name="src">要转换的Point3d数组</param>
        /// <returns>返回对应的Point3dList</returns>
        public static implicit operator Point3dList(Point3d[] src)
        {
            return src != null ? new Point3dList(src) : null;
        }

        /// <summary>
        /// Point3dCollection隐式转换为Point3dList
        /// </summary>
        /// <param name="src">要转换的Point3dCollection</param>
        /// <returns>返回对应的Point3dList</returns>
        public static implicit operator Point3dList(Point3dCollection src)
        {
            if (src != null)
            {
                Point3d[] ids = new Point3d[src.Count];
                src.CopyTo(ids, 0);
                return new Point3dList(ids);
            }
            else
            {
                return null;
            }
        }
    }

    /// <summary>
    /// ObjectId列表
    /// </summary>
    public class ObjectIdList : List<ObjectId>
    {
        /// <summary>
        /// 接受可变参数的构造函数
        /// </summary>
        /// <param name="args">ObjectId对象</param>
        public ObjectIdList(params ObjectId[] args)
        {
            AddRange(args);
        }

        /// <summary>
        /// ObjectIdList隐式转换为ObjectId数组
        /// </summary>
        /// <param name="src">要转换的ObjectIdList对象</param>
        /// <returns>返回对应的ObjectId数组</returns>
        public static implicit operator ObjectId[](ObjectIdList src)
        {
            return src != null ? src.ToArray() : null;
        }

        /// <summary>
        /// ObjectIdList隐式转换为ObjectIdCollection
        /// </summary>
        /// <param name="src">要转换的ObjectIdList对象</param>
        /// <returns>返回对应的ObjectIdCollection</returns>
        public static implicit operator ObjectIdCollection(ObjectIdList src)
        {
            return src != null ? new ObjectIdCollection(src) : null;
        }

        /// <summary>
        /// ObjectId数组隐式转换为ObjectIdList
        /// </summary>
        /// <param name="src">要转换的ObjectId数组</param>
        /// <returns>返回对应的ObjectIdList</returns>
        public static implicit operator ObjectIdList(ObjectId[] src)
        {
            return src != null ? new ObjectIdList(src) : null;
        }

        /// <summary>
        /// ObjectIdCollection隐式转换为ObjectIdList
        /// </summary>
        /// <param name="src">要转换的ObjectIdCollection</param>
        /// <returns>返回对应的ObjectIdList</returns>
        public static implicit operator ObjectIdList(ObjectIdCollection src)
        {
            if (src != null)
            {
                ObjectId[] ids = new ObjectId[src.Count];
                src.CopyTo(ids, 0);
                return new ObjectIdList(ids);
            }
            else
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Entity列表
    /// </summary>
    public class EntityList : List<Entity>
    {
        /// <summary>
        /// 接受可变参数的构造函数
        /// </summary>
        /// <param name="args">实体对象</param>
        public EntityList(params Entity[] args)
        {
            AddRange(args);
        }

        /// <summary>
        /// EntityList隐式转换为Entity数组
        /// </summary>
        /// <param name="src">要转换的EntityList</param>
        /// <returns>返回对应的Entity数组</returns>
        public static implicit operator Entity[](EntityList src)
        {
            return src != null ? src.ToArray() : null;
        }

        /// <summary>
        /// Entity数组隐式转换为EntityList
        /// </summary>
        /// <param name="src">要转换的Entity数组</param>
        /// <returns>返回对应的EntityList</returns>
        public static implicit operator EntityList(Entity[] src)
        {
            return src != null ? new EntityList(src) : null;
        }


    }

}
