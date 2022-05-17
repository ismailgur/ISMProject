using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ISMExtensions.Extensions
{
    public static class TypeExtension
    {
        public static bool IsNullable(this Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>));
        }

        public static bool IsNullAssignable(this Type type)
        {
            return IsNullable(type) || !type.IsValueType;
        }

        public static Type AsNonNullable(this Type type)
        {
            return type.IsNullable() ? Nullable.GetUnderlyingType(type) : type;
        }
    }
}
