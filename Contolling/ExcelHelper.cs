﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Controlling
{
    public static class ExcelHelper
    {
        public static DateOnly ParseDate(string cellValue)
        {
            return DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(cellValue)));
        }

        public static string GetCellValue(Cell cell, SharedStringTablePart stringTablePart)
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
    }

    public static class DateExtensions
    {
        public static string DayMonth(this DateOnly date)
        {
            return $"{date.Day}.{date.Month}.";
        }
        public static bool IsMoreDaysAgoThan(this DateOnly date, int daysAgo)
        {
            return date.AddDays(21) < DateOnly.FromDateTime(DateTime.Now);
        }
    }
}