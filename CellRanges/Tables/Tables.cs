using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.ComponentModel;

namespace ODSFiles
{
    public class Tables
    {
        internal int Index1From { get; set; }
        internal int Index2From { get; set; }

        internal Cells Cells { get; set; }

        public void LoadFromList<T>(List<T> list, bool printHeaders = false)
        {
            Type type = typeof(T);

            //печатаем заголовок или нет
            int header = printHeaders ? 1 : 0;

            int n = list.Count;
            int m = type.GetProperties().Length;

            IEnumerable<MethodInfo> value = type
                .GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.FlattenHierarchy)
                .Where(method => method.Name.StartsWith("get_"));


            for (int i = Index1From; i < Index1From + header + n; i++)
            {
                for (int j = Index2From; j < Index2From + m; j++)
                {
                    DisplayNameAttribute name = type
                        .GetProperties()[j - Index2From]
                        .GetCustomAttributes(typeof(DisplayNameAttribute), true)
                        .Cast<DisplayNameAttribute>()
                        .SingleOrDefault();

                    Cells[i, j].Value =
                        printHeaders && (i == Index1From)
                            ? name == null
                                ? type.GetProperties()[j - Index2From].Name
                                : name.DisplayName
                            : value.ElementAt(j - Index2From).Invoke(list[i - Index1From - header], parameters: null);
                }
            }
        }
    }
}