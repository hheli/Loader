using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LoaderModel
{

    public static class Helpers
    {
        public static double ParseToDouble(string value)
        {
            double result = double.NaN;
            value = value.Trim();
            if (!double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("ru-RU"), out result))
            {
                if (!double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("en-US"), out result))
                {
                    
                    throw new Exception(string.Format("Не удалось преобразовать в double '{0}'", value));
                    
                }
            }
            return result;
        }
        public static double ParseToDoubleEx(string value, string shname, int i)
        {
            // расширенный метод с выбросом исключения если не удалось используется для КПФ филиалов т.к. там вручную заполняют и могут косячить.
            double result = double.NaN;
            value = value.Trim();
            if (!double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("ru-RU"), out result))
            {
                if (!double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("en-US"), out result))
                {

                    throw new Exception(string.Format("Не удалось преобразовать в double '{0}' (лист:'{1}', строка '{2}')", value, shname, i));

                }
            }
            return result;
        }
        public static float ParseToFloat(string value)
        {
            float result = float.NaN;
            value = value.Trim();
            if (!float.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("ru-RU"), out result))
            {
                if (!float.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("en-US"), out result))
                {
                    return float.NaN;
                }
            }
            return result;
        }
        public static int CqConverter(string value)
        {
            switch (value)
            {
                case "Первая":
                    return 1;
                case "Вторая":
                    return 2;
                case "Третья":
                    return 3;
                case "Четвертая":
                    return 4;
                case "Пятая":
                    return 5;
                case "1":
                    return 1;
                case "2":
                    return 2;
                case "3":
                    return 3;
                case "4":
                    return 4;
                case "5":
                    return 5;
                case "Первая (283-П)":
                    return 1;
                case "Вторая (283-П)":
                    return 2;
                case "Третья (283-П)":
                    return 3;
                case "Четвертая (283-П)":
                    return 4;
                case "Пятая (283-П)":
                    return 5;
                default:
                    return 0;
            };
        }
    }

}
