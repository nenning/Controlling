using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;

namespace Controlling
{
    public static class ControllingExtensions
    {
        public static DateOnly ParseDate(this string cellValue)
        {
            return DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(cellValue)));
        }

        public static string GetCellValue(this Cell cell, SharedStringTablePart stringTablePart)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int index = int.Parse(cell.InnerText);
                return stringTablePart.SharedStringTable.ElementAt(index).InnerText;
            }
            else
            {
                return cell.InnerText;
            }
        }

        public static string DayMonth(this DateOnly date)
        {
            return $"{date.Day}.{date.Month}.";
        }

        public static bool IsMoreDaysAgoThan(this DateOnly date, int daysAgo)
        {
            return date.AddDays(daysAgo) < DateOnly.FromDateTime(DateTime.Now);
        }
        public static bool TryGetString(this JsonElement element, out string value)
        {
            if (element.ValueKind == JsonValueKind.String)
            {
                value = element.GetString();
                return true;
            }
            else
            {
                value = null;
                return false;
            }
        }

        public static bool TryGetDouble(this JsonElement element, out double value)
        {
            if (element.ValueKind == JsonValueKind.Number)
            {
                value = element.GetDouble();
                return true;
            }
            else
            {
                value = 0;
                return false;
            }
        }
        //public static T TryGetValue<T>(this JsonElement element)
        //{
        //    if (element.ValueKind == JsonValueKind.Null) { return default(T);
        //    } else if (element.ValueKind == JsonValueKind.Number) { return element.GetDouble(); }

        //    return default(T);
        //}

        public static JsonElement? TryGetProperty(this JsonElement element, string propertyName) =>
            element.TryGetProperty(propertyName, out var result) ? result : null;

    }
}
