using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL_Analise
{
    internal static class Sorter
    {
        private static Dictionary<Product, int> GetAllProduct(this IEnumerable<Date> dates)
        {
            var products = new Dictionary<Product, int>();

            var b = false;

            foreach (var date in dates)
            {

                foreach (var product in date.Products)
                {
                    b = false;

                    foreach (var p in products)
                    {
                        if (p.Key.Equals(product.Key))
                        {
                            products[p.Key] += product.Value;
                            b = true;
                            break;
                        }
                    }

                    if (!b)
                        products.Add(product.Key, product.Value);
                }
            }

            return products;
        }
        private static IConstruction GetConstruction(this Product product, EConstructionClass cclass)
        {
            switch (cclass)
            {
                case EConstructionClass.Product:
                    return product;

                case EConstructionClass.Aktiv:
                    return product.Active;

                case EConstructionClass.Passiv:
                    return product.Passive;

                case EConstructionClass.LVstavka:
                    return product.LeftVstavka;

                case EConstructionClass.RVstavka:
                    return product.RightVstavka;

                case EConstructionClass.Framuga:
                   return product.Framuga;

                default: throw new Exception("Неизвестный класс конструкции");
            }
        }

        public static Dictionary<IConstruction, int> SumDics(this Dictionary<IConstruction, int> dic1, Dictionary<IConstruction, int> dic2)
        {
            var dic = dic1;

            foreach(var p in dic2)
            {
                if (dic.ContainsKey(p.Key))
                    dic[p.Key] += p.Value;
                else
                    dic.Add(p.Key, p.Value);
            }

            return dic;
        }
        public static Dictionary<IConstruction, int> ReversDics(this Dictionary<IConstruction, int> dic)
        {
            var h = 0;
            var w = 0;
            var res = new Dictionary<IConstruction, int>();

            foreach(var d in dic)
            {
                var el = d.Key;
                el.setGabaryte(new Gabaryte(d.Key.Gabaryte.Width, d.Key.Gabaryte.Height));
                res.Add(el, d.Value);
            }

            return res;
        }

        public static Dictionary<IConstruction, int> GetAllConctruction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var products = dates.GetAllProduct();

            var constructions = new Dictionary<IConstruction, int>();

            var b = false;

            foreach (var product in products) 
            {
                b= false;

                var c = product.Key.GetConstruction(cclass);

                if (c == null) continue;

                if(constructions.ContainsKey(c))
                    constructions[c] += product.Value;
                else
                    constructions.Add(c, product.Value);
            }

            return constructions;
        }

        public static int AllConstructionCount(this IEnumerable<Date> dates, EConstructionClass cclass )
        {
            var cons = dates.GetAllConctruction(cclass);

            return cons.AllConstructionCount();  
        }
        public static int AllConstructionCount(this Dictionary<IConstruction, int> constructions)
        {
            int count = 0;

            foreach (var con in constructions)
                count += con.Value;

            return count;
        }

        public static Dictionary<int, int> GetAllHeightsConstruction(this Dictionary<IConstruction, int> constructions)
        {
            var allHeights = new Dictionary<int, int>();

            foreach (var c in constructions)
            {
                if (c.Key.Gabaryte.Height == 0)
                    continue;

                if (allHeights.ContainsKey(c.Key.Gabaryte.Height))
                    allHeights[c.Key.Gabaryte.Height] += c.Value;
                else
                    allHeights.Add(c.Key.Gabaryte.Height, c.Value);
            }

            return allHeights;
        }
        public static Dictionary<int, int> GetAllHeightsConstruction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var allConsts = dates.GetAllConctruction(cclass);

            return allConsts.GetAllHeightsConstruction();
        }

        public static Dictionary<int, int> GetAllWidthConstruction(this Dictionary<IConstruction, int> constructions)
        {
            var allWidth = new Dictionary<int, int>();

            foreach (var c in constructions)
            {
                if (c.Key.Gabaryte.Width == 0)
                    continue;

                if (allWidth.ContainsKey(c.Key.Gabaryte.Width))
                    allWidth[c.Key.Gabaryte.Width] += c.Value;
                else
                    allWidth.Add(c.Key.Gabaryte.Width, c.Value);
            }

            return allWidth;
        }
        public static Dictionary<int, int> GetAllWidthConstruction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var allConstruction = dates.GetAllConctruction(cclass);

            return allConstruction.GetAllWidthConstruction();
        }
        

        public static Gabaryte GetMinConstruction(this Dictionary<IConstruction, int> constructions)
        {
            var h = 0;
            var w = 0;

            var allH = constructions.GetAllHeightsConstruction();
            var allW = constructions.GetAllWidthConstruction();

            foreach (var height in allH)
            {
                if (h == 0)
                    h = height.Key;
                else if (h > height.Key)
                    h = height.Key;
            }

            foreach(var width in allW)
            { 
                if (w == 0)
                    w = width.Key;
                else if (w > width.Key)
                    w = width.Key;
            }

            return new Gabaryte(h, w);
        }
        public static Gabaryte GetMinConstruction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var constructions = dates.GetAllConctruction(cclass);

            return constructions.GetMinConstruction();
        }

        public static Gabaryte GetMaxConstruction(this Dictionary<IConstruction, int> constructions)
        {
            var h = 0;
            var w = 0;

            var allH = constructions.GetAllHeightsConstruction();
            var allW = constructions.GetAllWidthConstruction();

            foreach (var height in allH)
            {
                if (h < height.Key)
                    h = height.Key;
            }

            foreach(var width in allW) 
            {
                if (w < width.Key)
                    w = width.Key;
            }

            return new Gabaryte(h, w);
        }
        public static Gabaryte GetMaxConstruction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var constructions = dates.GetAllConctruction(cclass);

            return constructions.GetMaxConstruction();
        }
        
        public static Gabaryte GetMidConstruction(this Dictionary<IConstruction, int> constructions)
        {
            var min = constructions.GetMinConstruction();
            var max = constructions.GetMaxConstruction();

            return new Gabaryte((min.Height + max.Height) / 2, (min.Width + max.Width) / 2);
        }
        public static Gabaryte GetMidConstruction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var min = dates.GetMinConstruction(cclass);
            var max = dates.GetMaxConstruction(cclass);

            return new Gabaryte((min.Height + max.Height) / 2, (min.Width + max.Width) / 2);
        }

        public static Gabaryte GetModaConstruction(this Dictionary<IConstruction, int> constructions)
        {
            var allHeight = constructions.GetAllHeightsConstruction();
            var allWidth = constructions.GetAllWidthConstruction();

            var h = 0;
            var w = 0;
            var hc = 0;
            var wc = 0;

            foreach (var height in allHeight)
            {
                if (h == 0)
                {
                    h = height.Key;
                    hc = height.Value;
                    continue;
                }

                if (hc < height.Value)
                {
                    h = height.Key;
                    hc = height.Value;
                    continue;
                }
            }

            foreach (var width in allWidth)
            {
                if (w == 0)
                {
                    w = width.Key;
                    wc = width.Value;
                    continue;
                }

                if (wc < width.Value)
                {
                    w = width.Key;
                    wc = width.Value;
                    continue;
                }
            }

            return new Gabaryte(h, w);
        }
        public static Gabaryte GetModaConstruction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var allHeight = dates.GetAllHeightsConstruction(cclass);
            var allWidth = dates.GetAllWidthConstruction(cclass);

            var h = 0;
            var w = 0;
            var hc = 0;
            var wc = 0;

            foreach(var height in allHeight)
            {
                if (h == 0)
                {
                    h = height.Key;
                    hc = height.Value;
                    continue;
                }

                if(hc < height.Value)
                {
                    h = height.Key;
                    hc = height.Value;
                    continue;
                }
            }

            foreach(var width in allWidth)
            {
                if (w == 0)
                {
                    w = width.Key;
                    wc = width.Value;
                    continue;
                }

                if (wc < width.Value)
                {
                    w = width.Key;
                    wc = width.Value;
                    continue;
                }
            }

            return new Gabaryte(h, w);
        }

        public static IEnumerable<Gabaryte> Get75Construction(this Dictionary<IConstruction, int> constructions)
        {
            var allH = constructions
                .GetAllHeightsConstruction()
                .OrderByDescending(kvp => kvp.Value)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            var allW = constructions
                .GetAllWidthConstruction()
                .OrderByDescending(kvp => kvp.Value)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            var count = constructions.AllConstructionCount();

            var listH = new List<int>();
            var listW = new List<int>();

            var perc = 0.0;

            foreach (var h in allH)
            {
                perc += h.Value * 100 / count;
                listH.Add(h.Key);
                if (perc >= 75) break;
            }

            perc = 0.0;

            foreach (var w in allW)
            {
                perc += w.Value * 100 / count;
                listW.Add(w.Key);
                if (perc >= 75) break;
            }

            listH.Sort();
            listW.Sort();

            var res = new List<Gabaryte>();

            var g = new Gabaryte(listH[0], listW[0]);

            res.Add(g);

            g = new Gabaryte(listH[listH.Count - 1], listW[listW.Count - 1]);

            res.Add(g);

            return res;
        }
        public static IEnumerable<Gabaryte> Get75Construction(this IEnumerable<Date> dates, EConstructionClass cclass)
        {
            var allH = dates
                .GetAllHeightsConstruction(cclass)
                .OrderByDescending(kvp => kvp.Value)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            var allW = dates
                .GetAllWidthConstruction(cclass)
                .OrderByDescending(kvp => kvp.Value)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            var count = dates.AllConstructionCount(cclass);

            var listH = new List<int>();
            var listW = new List<int>();

            var perc = 0.0;

            foreach(var h in allH)
            {
                perc += h.Value * 100 / count;
                listH.Add(h.Key);
                if (perc >= 75) break;
            }

            perc = 0.0;

            foreach(var w in allW)
            {
                perc += w.Value * 100 / count;
                listW.Add(w.Key);
                if (perc >= 75) break;
            }

            listH.Sort();
            listW.Sort();

            var res = new List<Gabaryte>();

            var g = new Gabaryte(listH[0], listW[0]);

            res.Add(g);

            g = new Gabaryte(listH[listH.Count - 1], listW[listW.Count - 1]);

            res.Add(g);

            return res;
        }

        public static Dictionary<string, int> GetDiapHeights(this Dictionary<IConstruction, int> constructions)
        {
            var allH = constructions
                .GetAllHeightsConstruction()
                .OrderBy(pair => pair.Key)
                .ToDictionary(pair => pair.Key, pair => pair.Value);

            var d = 0;
            var c = 0;

            var res = new List<Dictionary<int, int>>();
            var dic = new Dictionary<int, int>();

            foreach(var h in allH)
            {
                c = h.Key / 100;

                if (d == 0)
                    d = c;

                if (d == c)
                    dic.Add(h.Key, h.Value);
                else
                {
                    res.Add(dic);

                    dic = new Dictionary<int, int>
                    {
                        { h.Key, h.Value }
                    };

                    d = c;
                }
            }

            var tmp = new Dictionary<string, int>();

            foreach(var diaps in res)
            {
                var str = $"{diaps.GetMin()}-{diaps.GetMax()}";
                var count = 0;

                foreach(var diap in diaps)
                    count += diap.Value;

                tmp.Add(str, count);
            }

            return tmp;
        }
        public static Dictionary<string, int> GetDiapWidth(this Dictionary<IConstruction, int> constructions)
        {
            var allW = constructions
                .GetAllWidthConstruction()
                .OrderBy(pair => pair.Key)
                .ToDictionary(pair => pair.Key, pair => pair.Value);

            var d = 0;
            var c = 0;

            var res = new List<Dictionary<int, int>>();
            var dic = new Dictionary<int, int>();

            foreach (var w in allW)
            {
                c = w.Key / 100;

                if (d == 0)
                    d = c;

                if (d == c)
                    dic.Add(w.Key, w.Value);
                else
                {
                    res.Add(dic);

                    dic = new Dictionary<int, int>
                    {
                        { w.Key, w.Value }
                    };

                    d = c;
                }
            }

            var tmp = new Dictionary<string, int>();

            foreach (var diaps in res)
            {
                var str = $"{diaps.GetMin()}-{diaps.GetMax()}";
                var count = 0;

                foreach (var diap in diaps)
                    count += diap.Value;

                tmp.Add(str, count);
            }

            return tmp;
        }

        private static int GetMin(this Dictionary<int, int> dic)
        {
            var e = 0;

            foreach(var d in dic)
            {
                if (e == 0)
                    e = d.Key;

                if (e > d.Key)
                    e = d.Key;
            }

            return e;
        }
        private static int GetMax(this Dictionary<int, int> dic)
        {
            var e = 0;

            foreach (var d in dic)
            {
                if (e < d.Key)
                    e = d.Key;
            }

            return e;
        }
        public static int GetCount(this Dictionary<string, int> dic)
        {
            var count = 0;

            foreach (var d in dic)
                count += d.Value;

            return count;
        }
    }
}
