using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

[assembly: InternalsVisibleTo("OpenXmlBinder.UnitTests")]

namespace OpenXmlBinder
{

    internal class ObjectParser
    {

        internal Dictionary<string, string?> parseResult = new Dictionary<string, string?>();

        public void ParseObject(object obj,string prefix="")
        {

            foreach (PropertyInfo prop in obj.GetType().GetProperties(BindingFlags.Public|BindingFlags.Instance))
            {
                if (prop.PropertyType.IsClass && prop.PropertyType!=typeof(String))
                {
                    object? nested = prop.GetValue(obj, null);
                    // Loop in
                    if (nested != null)
                    {
                        ParseObject(nested, prefix + prop.Name+".");

                    }
                }
                if (prop.PropertyType == typeof(DateTime))
                {
                    DateTime? date = (DateTime?)prop.GetValue(obj);

                    if (date != null)
                    {
                        parseResult[prefix + prop.Name] = (date).Value.ToString("D") ?? "";

                    }

                }
                else
                {
                    parseResult[prefix+prop.Name] = prop.GetValue(obj)?.ToString() ?? "";
                }
            }
        }
    }
}
