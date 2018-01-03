using System;
using System.Reflection;

namespace SharpDocx.Extensions
{
    internal static class AssemblyExtensions
    {
        public static object Invoke(
            this Assembly assembly,
            string typeName,
            object obj,
            string method,
            object[] parameters)
        {
            var mods = assembly.GetModules(false);
            var t = mods[0].GetType(typeName);
            var mi = t.GetMethod(method);
            return mi?.Invoke(obj, parameters);
        }

        public static object CreateInstance(
            this Assembly assembly,
            string typeName,
            object[] parameters)
        {
            var mods = assembly.GetModules(false);
            var t = mods[0].GetType(typeName);
            return Activator.CreateInstance(t, parameters);
        }
    }
}