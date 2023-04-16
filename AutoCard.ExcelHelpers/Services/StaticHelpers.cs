using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCard.ExcelHelpers.Services
{
    public static class StaticHelpers
    {
        public static object ConvertStringToType(string input, Type targetType)
        {
            if (targetType == typeof(string))
            {
                return input;
            }
            else if (targetType == typeof(int?))
            {
                if (int.TryParse(input, out int result))
                {
                    return result;
                }
            }
            else if (targetType == typeof(double?))
            {
                if (double.TryParse(input, out double result))
                {
                    return result;
                }
            }
            else if (targetType == typeof(DirectionType?))
            {
                if (Enum.TryParse(input, out DirectionType result))
                {
                    return result;
                }
            }

            return null;

        }
    }
}
