using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace ISMExtensions.Extensions
{
    public static class LinqExtensions
    {
        public static IEnumerable<TSource> DistinctBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            HashSet<TKey> seenKeys = new HashSet<TKey>();
            foreach (TSource element in source)
            {
                if (seenKeys.Add(keySelector(element)))
                {
                    yield return element;
                }
            }
        }

        public static IOrderedQueryable<T> OrderBy<T>(this IQueryable<T> source, string property)
        {
            return ApplyOrder<T>(source, property, "OrderBy");
        }

        public static IOrderedQueryable<T> OrderByDescending<T>(this IQueryable<T> source, string property)
        {
            return ApplyOrder<T>(source, property, "OrderByDescending");
        }

        static IOrderedQueryable<T> ApplyOrder<T>(IQueryable<T> source, string property, string methodName)
        {
            string[] props = property.Split('.');
            Type type = typeof(T);
            ParameterExpression arg = Expression.Parameter(type, "x");
            Expression expr = arg;
            foreach (string prop in props)
            {
                // use reflection (not ComponentModel) to mirror LINQ
                PropertyInfo pi = type.GetProperty(prop);
                expr = Expression.Property(expr, pi);
                type = pi.PropertyType;
            }
            Type delegateType = typeof(Func<,>).MakeGenericType(typeof(T), type);
            LambdaExpression lambda = Expression.Lambda(delegateType, expr, arg);

            object result = typeof(Queryable).GetMethods().Single(
                    method => method.Name == methodName
                            && method.IsGenericMethodDefinition
                            && method.GetGenericArguments().Length == 2
                            && method.GetParameters().Length == 2)
                    .MakeGenericMethod(typeof(T), type)
                    .Invoke(null, new object[] { source, lambda });
            return (IOrderedQueryable<T>)result;
        }

        public static void Batch<TSource>(
                  this IQueryable<TSource> source, int size, Action<TSource> action)
        {
            var totalCount = source.Count();

            var loopCnt = totalCount / size + (totalCount % size > 0 ? 1 : 0);
            for (int i = 0; i < loopCnt; i++)
            {
                if (i * size == totalCount)
                    break;

                foreach (var item in source.Skip(i * size).Take(size))
                {
                    action(item);
                }
            }
        }

        public static IEnumerable<T> Randomize<T>(this IEnumerable<T> pCol)
        {
            List<T> lResultado = new List<T>();
            List<T> lLista = pCol.ToList();
            Random lRandom = new Random();
            int lintPos = 0;

            while (lLista.Count > 0)
            {
                lintPos = lRandom.Next(lLista.Count);
                lResultado.Add(lLista[lintPos]);
                lLista.RemoveAt(lintPos);
            }

            return lResultado;
        }
    }
}
