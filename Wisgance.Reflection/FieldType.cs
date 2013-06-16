using System;

namespace Wisgance.Reflection
{
    /// <summary>
    /// This class is a pattern for dynamically class properties, 
    /// set all dynamic class properties in array of this class and then use it in class builders functions
    /// </summary>
    public class FieldMask
    {
        public string FieldName { get; set; }
        public Type FieldType { get; set; }
    }
}
