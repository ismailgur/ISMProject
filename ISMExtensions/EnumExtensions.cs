using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ISMExtensions.Extensions
{
    public static class EnumExtensions
    {
        public static string GetDisplay(this Enum value)
        {
            if (value == null)
                return string.Empty;

            Type type = value.GetType();
            return GetEnumDisplay(value.ToString(), type);
        }

        public static string GetFlagsDisplay(this Enum value)
        {
            Type type = value.GetType();
            var sb = new StringBuilder();


            foreach (Enum item in Enum.GetValues(type))
            {
                if (Convert.ToInt32(item) != 0)
                {

                    if ((value != null) && (value.HasFlag(item)))
                    {

                        sb.Append(GetEnumDisplay(item.ToString(), type));
                        sb.Append(",");
                    }
                }
            }

            return sb.Length > 0 ? sb.ToString().Substring(0, sb.ToString().Length - 1) : string.Empty;
        }

        public static string GetDisplayForEnum(this Type type, object value)
        {
            return GetEnumDisplay(value.ToString(), type);
        }

        private static string GetEnumDisplay(string value, Type type)
        {
            MemberInfo[] memberInfo = null;

            if (type.IsEnum)
            {
                memberInfo = type.GetMember(value);                
            }
            else if (type.IsNullableEnum())
            {
                memberInfo = Nullable.GetUnderlyingType(type).GetMember(value);
            }


            if (memberInfo != null && memberInfo.Length > 0)
            {
                // default to the first member info, it's for the specific enum value
                var info = memberInfo.First().GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault();

                if (info != null)
                    return ((DisplayAttribute)info).Name;
            }

            // no description - return the string value of the enum
            return value;
        }

        public static string GetDescription(this Enum value)
        {
            Type type = value.GetType();

            return GetEnumDescription(value.ToString(), type);
        }

        public static string GetDescriptionForEnum(this Type type, object value)
        {
            return GetEnumDescription(value.ToString(), type);
        }

        private static string GetEnumDescription(string value, Type type)
        {
            MemberInfo[] memberInfo = type.GetMember(value);

            if (memberInfo != null && memberInfo.Length > 0)
            {
                // default to the first member info, it's for the specific enum value
                var info = memberInfo.First().GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault();

                if (info != null)
                    return ((DescriptionAttribute)info).Description;
            }

            // no description - return the string value of the enum
            return value;
        }

        public static bool IsNullableEnum(this Type t)
        {
            Type u = Nullable.GetUnderlyingType(t);
            return t.BaseType.Name == "Enum" || ((u != null) && u.IsEnum);
        }

        public static bool IsDefined(this Enum value)
        {
            var values = Enum.GetValues(value.GetType()).Cast<int>().OrderBy(x => x);
            return values.Contains(Convert.ToInt32(value));
        }

        public static object GetEnumValueByDisplayName(this Type type, string displayName)
        {
            foreach (var item in Enum.GetNames(type))
            {
                MemberInfo[] memberInfo = null;

                if (type.IsEnum)
                {
                    memberInfo = type.GetMember(item);
                }
                else if (type.IsNullableEnum())
                {
                    memberInfo = Nullable.GetUnderlyingType(type).GetMember(item);
                }

                if (memberInfo != null && memberInfo.Length > 0)
                {
                    var info = memberInfo.First().GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault();

                    if (info != null)
                    {
                        if (((DisplayAttribute)info).Name.ToLower() == displayName.ToLower())
                        {
                            return Enum.Parse(type, item);
                        }
                    }
                }
            }
            return null;
        }
    }
}
