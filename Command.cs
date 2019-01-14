//柱帽偏移主要识别条件为结构柱族类型的柱帽构件
//对其他族类型不支持。
//2019.01.11有柱帽类型计算错误，需要重新验证：单个验证数据计算无错误，错误未解决
//2019.01.11对组添加名称无法将时间值给到组名称，（格式错误？？）//重新收集资料

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Selection;

namespace ColumnHat
{
    [Transaction(TransactionMode.Manual)]
    [Journaling(JournalingMode.UsingCommandData)]
    public class Command : IExternalCommand
    {
        private Curve curve;
        private double Lx;
        private Solid solid;
        public Solid msolid { get => solid;}

        private Solid floorsolid;
        public Solid mfloorsolid { get => solid; }

        private UV uV = new UV(0,0);
        public double H { get; private set; }

        private XYZ p;
        public XYZ mp
        {
            get => p;
        }
  

        Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            FloorSelection selection1 = new FloorSelection();
            ColumnSelection selection2 = new ColumnSelection();
            Reference refer = uidoc.Selection.PickObject(ObjectType.Element,selection1,"请选取车库顶板");
            if (null != refer)
            {
                Element el = uidoc.Document.GetElement(refer);
                Solid s = GetFloorSolid(el);
                Face f = GetAndFindTopFace(s);
                List<ElementId> ids = new List<ElementId>();
                IList<Reference> refer2 = uidoc.Selection.PickObjects(ObjectType.Element, selection2, "请选取柱帽");
                using (Transaction setvalue = new Transaction(uidoc.Document))
                {
                    setvalue.Start("set");
                    foreach (Reference re in refer2)
                    {
                        Element e = uidoc.Document.GetElement(re);
                        Solid s2 = GetSolid(e);
                        Face f2 = GetAndFindTopFace(s2);

                        if (IsInRangeFromFace(e, el, uidoc))
                        {
                            
                                
                                double value = GetProjectPoint(f2, f);
                                SetOffsetValue(value / 304.8, e);

                                
                        }
                        else
                        {
                            //TaskDialog.Show("Revit", "柱帽不再柱帽范围内");
                            ElementId elid = uidoc.Document.GetElement(re).Id;
                            ids.Add(elid);
                        }
                    }

                    if (0 != ids.Count)
                    {
                        MakeGroup(ids, uidoc.Document);
                    }
                    
                    setvalue.Commit();

                }
            }
            //TaskDialog.Show("Revit", f.Evaluate(new UV(0,0)).ToString() + "" + f2.Evaluate(new UV(0,0)).ToString()); 
            return Result.Succeeded;
        }

        /// <summary>
        /// 获取Solid
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        private Solid GetSolid(Element element)
        {
            
            Options option = new Options();
            option.ComputeReferences = true;
            option.DetailLevel = ViewDetailLevel.Fine;

            FamilyInstance instance = element as FamilyInstance;
            GeometryElement ge = instance.get_Geometry(option);
            foreach (GeometryObject geobject in ge)
            {
                if (geobject is Solid)
                {
                    solid = geobject as Solid;
                }
                else if (geobject is GeometryInstance)
                {
                    GeometryInstance geoInstance = geobject as GeometryInstance;
                    GeometryElement geoElemet = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject obj in geoElemet)
                    {
                        solid = obj as Solid;
                    }
                }
            }
            return solid;
        }
        /// <summary>
        /// 获取Solid顶部面
        /// </summary>
        /// <param name="solid"></param>
        /// <returns></returns>
        private Face GetAndFindTopFace(Solid solid)
        {
            PlanarFace pf = null;
            foreach (Face fa in solid.Faces)
            {
                pf = fa as PlanarFace;

                if (pf!=null)
                {
                    if (Math.Abs(pf.FaceNormal.X) < 0.01 && Math.Abs(pf.FaceNormal.Y) < 0.01 && pf.FaceNormal.Z > 0) //Z=-1为最低Z=1为最高    Z轴
                    {
                        break;
                    }
                }

            }

            return pf;

        }
        /// <summary>
        /// 获取顶板Solid
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        private Solid GetFloorSolid(Element element)
        {
            Options option = new Options();
            option.ComputeReferences = true;
            option.DetailLevel = ViewDetailLevel.Fine;
            Floor f = element as Floor;
            GeometryElement ge = f.get_Geometry(option);
            foreach (GeometryObject geobject in ge)
            {
                if (geobject is Solid)
                {
                    floorsolid = geobject as Solid;
                }
                else if (geobject is GeometryInstance)
                {
                    GeometryInstance geoInstance = geobject as GeometryInstance;
                    GeometryElement geoElemet = geoInstance.GetInstanceGeometry();
                    foreach (GeometryObject obj in geoElemet)
                    {
                        floorsolid = obj as Solid;
                    }
                }
            }

            return floorsolid;
        }
        /// <summary>
        /// 获取底部Face
        /// </summary>
        /// <param name="solid"></param>
        /// <returns></returns>
        private Face GetAndFindBottomFace(Solid solid)
        {
            PlanarFace pf = null;
            foreach (Face fa in solid.Faces)
            {
                pf = fa as PlanarFace;
                if (pf != null)
                {
                    if (Math.Abs(pf.FaceNormal.X) < 0.01 && Math.Abs(pf.FaceNormal.Y) < 0.01 && pf.FaceNormal.Z < 0) //Z=-1为最低Z=1为最高    Z轴
                    {
                        break;
                    }
                }
   
            }

            return pf;

        }
        /// <summary>
        /// 投影方法：获取柱帽偏移高度
        /// </summary>
        /// <param name="face"></param>
        /// <param name="face2"></param>
        /// <returns></returns>
        private double GetProjectPoint(Face face, Face floorface)
        {
            try
            {
                p = GetMinimumPointOnFace(face);

                XYZ p1 = new XYZ(p.X, p.Y, floorface.Evaluate(new UV(0, 0)).Z);
                XYZ p0 = GetMinimumPointOnFace(floorface);
                double lx = Math.Abs(p0.X - p1.X);
                EdgeArrayArray arrayArray = floorface.EdgeLoops;
                Dictionary<Curve, double> keys = new Dictionary<Curve, double>();
                foreach (EdgeArray array in arrayArray)
                {
                    for (int i = 0; i < array.Size; i++)
                    {
                        Edge e = array.get_Item(i);
                        Curve curve = e.AsCurve();
                        XYZ pe = curve.GetEndPoint(0);
                        XYZ pr = curve.GetEndPoint(1);
                        if ((float)pe.Z!= (float)pr.Z)
                        {
                            keys.Add(curve, curve.Length);
                        }
                        
                       
                       
                    }
                }

               
                Dictionary<Curve, double> dic = keys.OrderByDescending(m => m.Value)
                    .ToDictionary(m => m.Key, n => n.Value);

                Curve curve2 = dic.First().Key as Curve;
                curve = curve2;
                GetLx(face, floorface);
                Line li = curve2 as Line;
                XYZ v1 = li.Direction;
                if ((float)curve.GetEndPoint(0).X == (float)curve.GetEndPoint(1).X)
                {
                    XYZ p00 = new XYZ(p1.X, p0.Y, p1.Z);
                    Line li2 = Line.CreateBound(p00, p1);
                    XYZ v2 = li2.Direction;
                    double angle = v1.AngleTo(v2);
                    double h = Math.Tan(angle) * Lx; //获取垂直边长度
                    double h2 = p0.Z - p.Z;
                    H = Math.Abs(h) * 304.8 + h2 * 304.8;
                }
                else
                {
                    XYZ p00 = new XYZ(p0.X, p1.Y, p1.Z);
                    Line li2 = Line.CreateBound(p00, p1);
                    XYZ v2 = li2.Direction;
                    double angle = v1.AngleTo(v2);
                    double h = Math.Tan(angle) * Lx; //获取垂直边长度
                    double h2 = p0.Z - p.Z;
                    H = Math.Abs(h) * 304.8 + h2 * 304.8;
                }
                
            }
            catch (Exception e)
            {
                string mess = e.ToString();
                TaskDialog.Show("Revit", mess);
            }

             
            return H;
        }

       

        /// <summary>
        /// 获取面上最小点
        /// </summary>
        /// <param name="face"></param>
        /// <returns></returns>
        private XYZ GetMinimumPointOnFace(Face face)
        {
            List<XYZ> xYZs =GetFaceXyz(face);
             
            var maxxyz = xYZs[0];
            for (int a = 0; a < xYZs.Count; a++)
            {
                if (maxxyz.X > xYZs[a].X&& maxxyz.Y > xYZs[a].Y&& maxxyz.Z >= xYZs[a].Z)
                {
                    maxxyz = xYZs[a];
                }
            }
            
                   
            return maxxyz;
        }
        /// <summary>
        /// 获取两个面中距目标点最短距离
        /// </summary>
        /// <param name="face"></param>
        /// <param name="floorface"></param>
        /// <returns></returns>
        private double GetLx(Face face, Face floorface)
        {
            
            XYZ p0 = curve.GetEndPoint(0);
            XYZ p1 = curve.GetEndPoint(1);
            if (p0.X != p1.X && p0.Y == p1.Y)   //再假设Z值一致的情况下，边界线只存在平行于Y轴或X轴
            {
                List<XYZ> xYZs1 = GetFaceXyz(face);      //柱帽顶部四个点
                List<XYZ> xYZs2 = GetFaceXyz(floorface);   //板顶面四个点
                for (int i = 0; i < xYZs1.Count; i++)
                {
                    for (int j = xYZs1.Count - 1; j > 1; j--) //删除重复值，获取长度值两边
                    {
                        if (i == j) continue;
                        if (xYZs1[i].X == xYZs1[j].X  )
                        {
                            xYZs1.RemoveAt(j);
                        }
                    }
                }
                for (int a = 0; a < xYZs1.Count; a++)
                {
                    for (int b = xYZs1.Count - 1; b > 1; b--)
                    {
                        if (a == b) continue;
                        if (xYZs1[a].X == xYZs1[b].X)
                        {
                            xYZs1.RemoveAt(b);
                        }
                    }
                }

                List<double> ml = new List<double>();
                for (int m = 0; m < xYZs1.Count; m++)
                {
                    for (int n = 0; n < xYZs2.Count; n++)
                    {
                        double l = Math.Abs(xYZs1[m].X - xYZs2[n].X);
                        
                        if (l-0 != 0)
                        {
                            ml.Add(l);
                        }

                    }
                }

                //TaskDialog.Show("Revit", ml[0].ToString() + " " + ml[1].ToString());
                Lx = ml.Min();
                ml.Clear();
            }
            else
            {
                List<XYZ> xYZs1 = GetFaceXyz(face);      //柱帽顶部四个点
                List<XYZ> xYZs2 = GetFaceXyz(floorface);   //板顶面四个点
                for (int i = 0; i < xYZs1.Count; i++)
                {
                    for (int j = xYZs1.Count - 1; j > 1; j--) //删除重复值，获取长度值两边
                    {
                        if (i == j) continue;
                        if (xYZs1[i].Y == xYZs1[j].Y)
                        {
                            xYZs1.RemoveAt(j);
                        }
                    }
                }
                for (int a = 0; a < xYZs1.Count; a++)
                {
                    for (int b = xYZs1.Count - 1; b > 1; b--)
                    {
                        if (a == b) continue;
                        if (xYZs1[a].Y == xYZs1[b].Y)
                        {
                            xYZs1.RemoveAt(b);
                        }
                    }
                }

                List<double> ml = new List<double>();
                for (int m = 0; m < xYZs1.Count; m++)
                {
                    for (int n = 0; n < xYZs2.Count; n++)
                    {
                        double l = Math.Abs(xYZs1[m].Y - xYZs2[n].Y);
                        ml.Add(l);
                    }
                }

                //TaskDialog.Show("Revit", ml.Min().ToString());
                Lx = ml.Min();
            }

            return Lx;
        }
        /// <summary>
        ///获取面四个点
        /// </summary>
        /// <param name="face"></param>
        /// <returns></returns>
        private List<XYZ> GetFaceXyz(Face face)
        {
            List<XYZ> xYZs = new List<XYZ>();
            var arrayarray = face.EdgeLoops;
            foreach (EdgeArray array in arrayarray)
            {
                foreach (Edge e in array)
                {
                    Curve curve = e.AsCurve();
                    xYZs.Add(curve.GetEndPoint(0));
                    xYZs.Add(curve.GetEndPoint(1));
                }
            }

            for (int i = 0; i < xYZs.Count; i++)
            {
                for (int j = xYZs.Count - 1; j > 1; j--)
                {
                    if (i == j) continue;
                    if (xYZs[i].X == xYZs[j].X && xYZs[i].Y == xYZs[j].Y && xYZs[i].Z == xYZs[j].Z)
                    {
                        xYZs.RemoveAt(j);
                    }
                }
            }

            return xYZs;
        }
        /// <summary>
        /// 车库顶板过滤器（结构板过滤）
        /// </summary>
        public class FloorSelection : ISelectionFilter
        {
            bool ISelectionFilter.AllowElement(Element elem)
            {
                if ( elem.Category.Name == "楼板")
                {
                    return true;
                }

                return true;
            }

            bool ISelectionFilter.AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
        /// <summary>
        /// 柱帽过滤器（柱帽结构柱制作）
        /// </summary>
        public class ColumnSelection : ISelectionFilter
        {
            bool ISelectionFilter.AllowElement(Element elem)
            {
                if (elem.Category.Name == "结构柱")
                {
                    return true;
                }

                return true;
            }

            bool ISelectionFilter.AllowReference(Reference reference, XYZ position)
            {
                return false;
            }
        }
        /// <summary>
        /// 自动筛选楼板下的柱帽
        /// </summary>
        /// <param name="face"></param>
        /// <param name="floorface"></param>
        private bool IsInRangeFromFace(Element column,Element floor,UIDocument uIDocument)
        {
            bool result = false;
            View view = uIDocument.ActiveView;
            Floor fl = floor as Floor;
            BoundingBoxXYZ boundingBox1 = fl.get_BoundingBox(view);
            BoundingBoxXYZ boundingBox2 = column.get_BoundingBox(view);
            XYZ nmax = boundingBox1.Max;
            XYZ nmin = boundingBox1.Min;
            XYZ nmax2 = boundingBox2.Max;
            XYZ nmin2 = boundingBox2.Min;
            if (nmin.Y <= nmax2.Y && nmin.X <= nmax2.X||nmax.Y<=nmin2.Y&&nmax.X<=nmin2.X)
            {
                result = false;
            }
            if (nmin.X <= nmin2.X && nmin2.X <= nmax.X && nmin.Y <= nmin2.Y && nmin2.Y <= nmax.Y &&
                nmin.X <= nmax2.X && nmax2.X <= nmax.X && nmin.Y <= nmax2.Y && nmax2.Y <= nmax.Y)
            {
                result = true;
            }

            return result;
        }
        /// <summary>
        /// 传入数据对柱帽进行偏移
        /// </summary>
        /// <param name="value"></param>
        /// <param name="element"></param>
        private void SetOffsetValue(double value, Element element)
        {
            if (element.Category.Name == "结构柱")
            {
                Parameter para = element.get_Parameter(BuiltInParameter.SCHEDULE_TOP_LEVEL_OFFSET_PARAM);
                double d = para.AsDouble();
                double revalue = d + value;
                para.Set(revalue);
            }
        }
        /// <summary>
        /// 将不再楼板内的柱帽打组，后续操作需要用户手动调整
        /// 对组进行时间赋值失败
        /// </summary>
        /// <param name="elements"></param>
        /// <param name="document"></param>
        private void MakeGroup(List<ElementId> elements,Document document)
        {
            if (elements.Count != 0)
            {
                Group group = null;
                group = document.Create.NewGroup(elements);
                //System.DateTime time = new System.DateTime();
                //time = DateTime.Now;
                //string str = time.ToString("yyyyMMdd");
                //string name = str + time.Second.ToString();
                

            }
        }

    }

}
