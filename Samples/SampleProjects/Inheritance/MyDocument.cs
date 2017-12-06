using System.Collections.Generic;
using SharpDocx;

namespace Inheritance
{
    public abstract class MyDocument : DocumentBase
    {
        protected MyDocument()
        {
            MyProperty = "very thorough";
        }

        public string MyProperty { get; set; }

        public new static List<string> GetUsingDirectives()
        {
            return new List<string>
            {
                "using Inheritance;"

                //"using static System.Math;" 
                // Requires support for C# 6.
                // See https://stackoverflow.com/questions/31639602/using-c-sharp-6-features-with-codedomprovider-rosyln
            };
        }

        public new static List<string> GetReferencedAssemblies()
        {
            return new List<string>
            {
                typeof(MyDocument).Assembly.Location
            };
        }
    }
}