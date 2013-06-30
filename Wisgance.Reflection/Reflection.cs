using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Wisgance.Reflection
{
    public class ObjUtility
    {
        public static List<string> GetPropertyInfo<T>()
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
        }

        public static List<string> GetPropertyInfo(object T)
        {
            var propertyInfos = T.GetType().GetProperties();
            return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
        }

    }
}
