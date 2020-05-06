using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace widget
{
    class CellXlsX
    {
        private string LocalPath;
        private string RemotePath;
        private string FileName;
        private string Sheet;
        private int Row;
        private int Column;
        public CellXlsX(string cesta, string sheet, int row, int column)
        {
            this.FileName = Path.GetFileName(cesta);
            this.LocalPath = Path.Combine(Path.GetTempPath(), FileName);
            this.RemotePath = Path.GetFullPath(cesta);
            this.Sheet = sheet;
            this.Row = --row;
            this.Column = --column;
        }
        private XSSFWorkbook xssfwb;      
        public void ReadExcelFile()
        {
            if (File.Exists(LocalPath))
                File.Delete(LocalPath);
            try
            {
                File.Copy(RemotePath, LocalPath);
            }
            catch
            {
                MessageBox.Show(new Form() { TopMost = true }, "Soubor -" + FileName + "- zadaný v konfiguračním souboru na zadané cestě neexistuje.");
            }           
            using (FileStream file = new FileStream(LocalPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                xssfwb = new XSSFWorkbook(file);
                if (File.Exists(LocalPath))
                    File.Delete(LocalPath);
            }
        }
        public string FillCell()
        {
            ISheet sheet = null;
            object cell = null;
            sheet = xssfwb.GetSheet(Sheet);
            try
            {
                if (sheet == null)
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                MessageBox.Show(new Form() { TopMost = true }, "List -" + Sheet + "- v souboru -" + FileName + "- neexistuje.");
                return (string)null;
            }
            try
            {
                var sheetRow = sheet.GetRow(Row);
                var cellType = sheetRow.GetCell(Column);         
                cell = new object();
                if (cellType.CellType == CellType.String)
                    cell = sheetRow.GetCell(Column).StringCellValue;
                else if (cellType.CellType == CellType.Formula)
                    cell = sheetRow.GetCell(Column).NumericCellValue;
                else if (cellType.CellType == CellType.Numeric)
                cell = sheetRow.GetCell(Column).NumericCellValue;               
            }
            catch
            {
                MessageBox.Show(new Form() { TopMost = true }, "V zadané buňce v souboru -" + FileName + "- nejsou data, nebo nejsou typu -text/číslo-");
                return (string)null;
            }
            return cell.ToString();
        }
    }   
}
