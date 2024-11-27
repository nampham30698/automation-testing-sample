using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationTestingSample.Core.Utilities
{
    public static class Parser
    {
        public static double ParseDoube(string value)
        {
            if (double.TryParse(value, out var result))
            {
                return result;
            }
            return 0;
        }
    }
}
