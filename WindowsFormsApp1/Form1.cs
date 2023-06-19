using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Worksheet RemoveCellBorders(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, у которой нужно удалить границы
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Удаление границ ячейки
            Style style = cell.GetStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.None;
            cell.SetStyle(style);
            return worksheet;
        }

        private Worksheet ButtomCellBorder(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, у которой нужно удалить границы
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Удаление границ ячейки
            Style style = cell.GetStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.None;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.None;
            cell.SetStyle(style);
            return worksheet;
        }
        private Worksheet ThinCellBorder(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, у которой нужно удалить границы
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Удаление границ ячейки
            Style style = cell.GetStyle();
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            cell.SetStyle(style);
            return worksheet;
        }

        private Worksheet CellsBorder(Worksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn,string style)
        {
            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
            {
                for (int colIndex = firstColumn; colIndex <= lastColumn; colIndex++)
                {
                    switch (style)
                    {
                        case "BottomThin":
                            worksheet = ButtomCellBorder(worksheet, rowIndex, colIndex);
                            break;
                        case "RemoveBorder":
                            worksheet = RemoveCellBorders(worksheet, rowIndex, colIndex);
                            break;
                        case "ThinBorder":
                            worksheet = ThinCellBorder(worksheet, rowIndex, colIndex);
                            break;
                    }
                }
            }

            return worksheet;
        }

        private Worksheet MergeCells(Worksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            string firstCellAddress = CellsHelper.CellIndexToName(firstRow, firstColumn);
            string lastCellAddress = CellsHelper.CellIndexToName(lastRow, lastColumn);
            Range range = worksheet.Cells.CreateRange(firstCellAddress, lastCellAddress);
            range.Merge();
            return worksheet;
        }


        private Worksheet SetBoldText(Worksheet worksheet, int rowIndex, int columnIndex)
        {
            // Получение ячейки, для которой нужно установить полужирный текст
            Cell cell = worksheet.Cells[rowIndex, columnIndex];
            // Установка полужирного шрифта для ячейки
            Style style = cell.GetStyle();
            style.Font.IsBold = true;
            cell.SetStyle(style);
            return worksheet;
        }

        private Worksheet SetStylesCells(Worksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn, string[] style, string putValue) {
            for (int indexElement=0;indexElement<style.Length;indexElement++ ) { 
                switch (style[indexElement])
                {
                    case "RemoveBorder":
                        worksheet = CellsBorder(worksheet, firstRow, firstColumn, lastRow, lastColumn, "RemoveBorder");
                        break;
                    case "BottomThin":
                        worksheet = CellsBorder(worksheet, firstRow, firstColumn, lastRow, lastColumn, "RemoveBorder");
                        break;
                    case "ThinBorder":
                        worksheet = CellsBorder(worksheet, firstRow, firstColumn, lastRow, lastColumn, "ThinBorder");
                        break;
                    case "MergeCells":
                        worksheet = MergeCells(worksheet, firstRow, firstColumn, lastRow, lastColumn);
                        break;
                    case "SetBoldText":
                        worksheet = SetBoldText(worksheet, firstRow, firstColumn);
                        break;
                }
            }
            Cell cell = worksheet.Cells[firstRow, firstColumn];
            cell.PutValue(putValue);
            return worksheet;
        }

        private string GetPrice(string startStr, string finishStr) {
            
            int startInt;
            int finishInt;
            string price;

            if (int.TryParse(startStr, out startInt) && int.TryParse(finishStr, out finishInt))
            {
                price = (finishInt - startInt).ToString();
            }
            else
            {
                MessageBox.Show("Введите корректные данные!");
                price = "0";
            }
            return price;
        }

        private string GetSumm(string startStr, string finishStr)
        {

            double startInt;
            double finishInt;
            string price;

            if (double.TryParse(startStr, out startInt) && double.TryParse(finishStr, out finishInt))
            {
                price = (finishInt*startInt).ToString();
            }
            else
            {
                MessageBox.Show("Введите корректные данные!");
                price = "0";
            }
            return price;
        }

        private string GetFinal(string startStr, string finishStr)
        {

            double startInt;
            double finishInt;
            string price;

            if (double.TryParse(startStr, out startInt) && double.TryParse(finishStr, out finishInt))
            {
                price = (finishInt + startInt).ToString();
            }
            else
            {
                price = startStr;
            }
            return price;
        }

        private string GetMonth(int month) {
            switch (month)
            {
                case 1:
                    return "Январь";   
                case 2:
                    return "Февраль";   
                case 3:
                    return "Март";    
                case 4:
                    return "Апрель";    
                case 5:
                    return "Май";    
                case 6:
                    return "Июнь";   
                case 7:
                    return "Июль";    
                case 8:
                    return "Август";   
                case 9:
                    return "Сентябрь";   
                case 10:
                    return "Октябрь";    
                case 11:
                    return "Ноябрь";   
                case 12:
                    return "Декабрь";     
                default:
                    return "Некорректный номер месяца";       
            }
        }

        private void GenerateExcelTemplate()
        {
            string price = GetPrice(this.textBox3.Text, this.textBox4.Text);
            string summ = GetSumm(price, "33,76");
            string finalSumm = GetFinal(summ, this.textBox5.Text);

            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet = CellsBorder(worksheet, 1, 1, 10, 10, "RemoveBorder");
            worksheet = SetStylesCells(worksheet, 1, 1, 1, 3, new string[] { "RemoveBorder", "MergeCells", "SetBoldText" }, "СЧЕТ - ИЗВЕЩЕНИЯ за ");
            worksheet = SetStylesCells(worksheet, 1, 4, 1, 4, new string[] { "BottomThin", "SetBoldText" }, GetMonth(this.dateTimePicker1.Value.Month));
            worksheet = SetStylesCells(worksheet, 1, 5, 1, 5, new string[] { "BottomThin", "SetBoldText" }, this.dateTimePicker1.Value.Year.ToString()+" г.");
            worksheet = SetStylesCells(worksheet, 2, 1, 2, 1, new string[] {  }, "Получатель: ");
            worksheet = SetStylesCells(worksheet, 2, 2, 2, 3, new string[] { "SetBoldText", "MergeCells" }, "МУП \"Архиповка\"");
            worksheet = SetStylesCells(worksheet, 3, 1, 3, 1, new string[] { "SetBoldText"}, "Абонент: ");
            worksheet = SetStylesCells(worksheet, 3, 2, 3, 5, new string[] { "MergeCells", "BottomThin", "SetBoldText"  }, this.textBox2.Text);
            worksheet = SetStylesCells(worksheet, 4, 1, 4, 1, new string[] {  }, "Адрес: ");
            worksheet = SetStylesCells(worksheet, 4, 2, 4, 5, new string[] { }, this.textBox1.Text);
            worksheet = CellsBorder(worksheet, 5, 1, 8, 6, "ThinBorder");
            worksheet = SetStylesCells(worksheet, 5, 1, 5, 1, new string[] { }, "Показ. Счетч.");
            worksheet = SetStylesCells(worksheet, 5, 2, 5, 2, new string[] { }, "Начальное");
            worksheet = SetStylesCells(worksheet, 5, 3, 5, 3, new string[] { }, "Конечное");
            worksheet = SetStylesCells(worksheet, 5, 4, 5, 4, new string[] { }, "Расход, куб.м");
            worksheet = SetStylesCells(worksheet, 5, 5, 5, 5, new string[] { }, "Цена");
            worksheet = SetStylesCells(worksheet, 5, 6, 6, 6, new string[] { }, "Сумма");

            worksheet = SetStylesCells(worksheet, 6, 1, 6, 1, new string[] { }, "ХВС");
            worksheet = SetStylesCells(worksheet, 6, 2, 6, 2, new string[] { }, this.textBox3.Text);
            worksheet = SetStylesCells(worksheet, 6, 3, 6, 3, new string[] { }, this.textBox4.Text);
            worksheet = SetStylesCells(worksheet, 6, 4, 6, 4, new string[] { }, "33,76");
            worksheet = SetStylesCells(worksheet, 6, 5, 6, 5, new string[] { }, price);
            worksheet = SetStylesCells(worksheet, 6, 6, 6, 6, new string[] { }, summ);

            worksheet = SetStylesCells(worksheet, 7, 1, 7, 1, new string[] { }, "ГВС");
            worksheet = SetStylesCells(worksheet, 7, 2, 7, 2, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 3, 7, 3, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 4, 7, 4, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 5, 7, 5, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 7, 6, 7, 6, new string[] { }, "");

            worksheet = SetStylesCells(worksheet, 8, 1, 8, 1, new string[] { }, "Водоотвед.");
            worksheet = SetStylesCells(worksheet, 8, 2, 8, 2, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 3, 8, 3, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 4, 8, 4, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 5, 8, 5, new string[] { }, "");
            worksheet = SetStylesCells(worksheet, 8, 6, 8, 6, new string[] { }, "");

            worksheet = SetStylesCells(worksheet, 9, 1, 9, 2, new string[] { "MergeCells"}, "Задолж-ть (переплата)");
            worksheet = SetStylesCells(worksheet, 9, 3, 9, 3, new string[] { "BottomThin", "SetBoldText" }, this.textBox5.Text);

            worksheet = SetStylesCells(worksheet, 9, 4, 9, 5, new string[] { "MergeCells", "SetBoldText" }, "Всего к оплате");
            worksheet = SetStylesCells(worksheet, 9, 6, 9, 6, new string[] { "BottomThin", "SetBoldText" }, finalSumm);
            // Сохранение книги Excel в файл
            workbook.Save("Template.xlsx");
        }


        private void ButtonExl(object sender, EventArgs e)
        {
            GenerateExcelTemplate();
            MessageBox.Show("Шаблон Excel успешно создан!");
        }

       
    }
}
