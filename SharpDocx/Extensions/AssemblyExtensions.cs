using System;
using System.Reflection;

namespace SharpDocx.Extensions
{
    public static class AssemblyExtensions
    {
        public static object Invoke(
            this Assembly assembly,
            string typeName,
            object obj,
            string method,
            object[] parameters)
        {
            Module[] mods = assembly.GetModules(false);
            Type t = mods[0].GetType(typeName);
            MethodInfo mi = t.GetMethod(method);
            return mi.Invoke(obj, parameters);
        }

        public static object CreateInstance(
            this Assembly assembly, 
            string typeName, 
            object[] parameters)
        {
            Module[] mods = assembly.GetModules(false);
            Type t = mods[0].GetType(typeName);
            return Activator.CreateInstance(t, parameters);
        }
    }
}
