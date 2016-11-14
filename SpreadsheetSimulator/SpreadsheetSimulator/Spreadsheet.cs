using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpreadsheetSimulator
{
    public struct CellCoordinates
    {
        private int x, y;
        private string name;
        public int X
        {
            get
            {
                return x;
            }
        }

        public int Y
        {
            get
            {
                return y;
            }
        }

        public string Name
        {
            get
            {
                return name;
            }
        }

        public CellCoordinates(int X, int Y, string name)
        {
            x = X;
            y = Y;
            this.name = name;
        }
    }

    public partial class Spreadsheet : Form
    {
        //all spreadsheet cells
        private List<CellCoordinates> cellsCoordinates = new List<CellCoordinates>();

        private const string alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public Spreadsheet()
        {
            InitializeComponent();
        }
        
        private void Spreadsheet_Load(object sender, EventArgs e)
        {

            SetSpreadsheetSize(50, 50);

            //Add columns/rows when scrolling finished
            dataGridView1.Scroll += DataGridView_Scroll;


            dataGridView1.CellBeginEdit += dataGridView_CellBeginEdit;
            dataGridView1.CellEndEdit += dataGridView_CellEndEdit;
   
        }
        
        
            public void SetSpreadsheetSize(int rowCount, int columnCount)
            {

                for (int i = dataGridView1.ColumnCount; i <= columnCount; i++)
                {
                    AddColumns();
                }
                for (int i = dataGridView1.RowCount; i <= rowCount; i++)
                {
                    AddRows();
                }

            }

            private void DataGridView_Scroll(object sender, ScrollEventArgs e)
            {

                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll && e.NewValue >= dataGridView1.Rows.Count - GetDisplayedRowsCount())
                    AddRows();

                if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll && dataGridView1.FirstDisplayedScrollingColumnIndex >= dataGridView1.Columns.Count - GetDisplayedColumnsCount())
                    AddColumns();

            }

            private void AddRows()
            {

                DataGridViewRow row = new DataGridViewRow();
                int rowHeader = dataGridView1.RowCount;
                row.HeaderCell.Value = rowHeader.ToString();
                dataGridView1.Rows.Add(row);
            
                //save new cells when added row
                for (int i = 0; i <= dataGridView1.ColumnCount - 1; i++)
                {
                    CellCoordinates coordinates = new CellCoordinates(i, rowHeader, GetColumnHeader(i) + (rowHeader + 1));
                    cellsCoordinates.Add(coordinates);
                }
            }

            private void AddColumns()
            {
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();

                string columnHeader = GetColumnHeader(dataGridView1.ColumnCount);
                column.Name = columnHeader;
                dataGridView1.Columns.Add(column);

                //save new cells when added column
                for (int i = 0; i <= dataGridView1.RowCount; i++)
                {
                    CellCoordinates coordinates = new CellCoordinates(dataGridView1.ColumnCount - 1, i, columnHeader + (i + 1).ToString());
                    cellsCoordinates.Add(coordinates);
                }


            }

            public string GetColumnHeader(int ColumnIndex)
            {

                string columnHeader = "";

                do
                {
                    columnHeader = alpha[ColumnIndex % 26] + columnHeader;
                    ColumnIndex = ColumnIndex / 26 - 1;

                } while (ColumnIndex != -1);

                return columnHeader;
            }

            public int GetDisplayedRowsCount()
            {
                int count = dataGridView1.Rows[dataGridView1.FirstDisplayedScrollingRowIndex].Height;
                count = dataGridView1.Height / count;
                return count;
            }

            public int GetDisplayedColumnsCount()
            {
                int count = dataGridView1.Columns[dataGridView1.FirstDisplayedScrollingColumnIndex].Width;
                count = dataGridView1.Width / count + 1;
                return count;
            }

            private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
            {

                string EditText = (string)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value ?? null;
                Cell cell;

                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag != null && dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag is Cell)
                {
                    cell = (Cell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag;
                    if (cell.expression != EditText)
                        ExpressionChange(EditText, ref cell);
                }
                else
                {
                    if (EditText == null)
                        return;

                    cell = new Cell(cellsCoordinates.Find(x => x.X == e.ColumnIndex && x.Y == e.RowIndex));
                    ExpressionChange(EditText, ref cell);
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag = cell;
                }

                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = cell.display;
            }

            private void dataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag != null && dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag is Cell)
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = ((Cell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Tag).expression;
            }

            private List<CellCoordinates> GetExpressionUsedCells(string ValidExpression)
            {
                List<CellCoordinates> UsedCells = new List<CellCoordinates>();
                string cellName = null;

                for (int i = 0; i < ValidExpression.Length; i++)
                {
                    if (!Char.IsNumber(ValidExpression[i]))
                    {
                        while (i < ValidExpression.Length && ValidExpression[i] != '+' && ValidExpression[i] != '-' && ValidExpression[i] != '*' && ValidExpression[i] != '/')

                        {
                            cellName = cellName + ValidExpression[i];
                            i++;
                        }

                        if (cellName != null)
                        {
                            CellCoordinates cell = cellsCoordinates.Find(x => x.Name == cellName);
                            if (cell.Name == null)
                                cell = new CellCoordinates(0, 0, cellName);

                            UsedCells.Add(cell);
                        }
                        cellName = null;
                    }
                }

                return UsedCells;
            }

            private void ExpressionChange(string Expression, ref Cell cell)
            {

                if (Expression == null)
                {
                    cell.display = null;
                    cell.expression = null;
                    CellClearUsedCells(ref cell);
                }
                else
                {
                    if (cell.expression != Expression)
                        CellClearUsedCells(ref cell);

                    switch (Expression[0])
                    {
                        case '=':

                            string checkExpression = Expression.Substring(1);
                            checkExpression = checkExpression.Replace(" ", string.Empty);

                            if (cell.expression != Expression)
                            {
                                cell.ExpressionValid = ExpressionValid(checkExpression);
                                cell.CircularReference = CircularReference(cell, checkExpression);

                                if (cell.ExpressionValid && !cell.CircularReference)
                                {
                                    List<CellCoordinates> usedCells = GetExpressionUsedCells(checkExpression);
                                    foreach (var item in usedCells)
                                    {
                                        CellAddFieldUsedCell(ref cell, item.Name);
                                    }
                                }

                            }

                            if (!cell.ExpressionValid)
                            {
                                cell.display = "#Not valid Expression";
                                break;
                            }

                            if (cell.CircularReference)
                            {
                                cell.display = "#Don't use circular references";
                                break;
                            }



                            //get expression value
                            foreach (var item in cell.usedСells)
                            {
                                string value = GetCellValue(item.Name);

                                if (value == null && Regex.IsMatch(checkExpression, "[+,-,*,/]"))
                                    value = "0";

                                //check if return error
                                if (value != null && value[0] == '#')
                                {
                                    cell.display = value;
                                    goto Out;
                                }

                                //check if used mathematical operations to string
                                double d;
                                if (Regex.IsMatch(checkExpression, "[+,-,*,/]") && !Double.TryParse(value, out d))
                                {
                                    cell.display = "#Not use mathematical operations to string";
                                    goto Out;
                                }

                                checkExpression = checkExpression.Replace(item.Name, value);
                            }

                            if (!Regex.IsMatch(checkExpression, "[+,-,*,/]"))
                                cell.display = checkExpression;
                            else
                                cell.display = Convert.ToDouble(new DataTable().Compute(checkExpression, null)).ToString();

                            Out:
                            break;

                        case '\'':
                            cell.display = Expression.Substring(1);
                            break;
                        default:
                            double number;
                            if (!double.TryParse(Expression.Replace('.', ','), out number))
                                cell.display = "#For the input line set prefix \'";
                            else
                                cell.display = Expression.Replace(',', '.');

                            break;
                    }
                    cell.expression = Expression;
                }

                dataGridView1.Rows[cell.Coordinates.Y].Cells[cell.Coordinates.X].Value = cell.display;

                //change cells that used this cell
                foreach (var item in cell.usedToСells.ToList())
                {
                    Cell findCell = GetCell(item.Name);
                    if (findCell != null)
                        ExpressionChange(findCell.expression, ref findCell);
                }

            }

            private bool CircularReference(Cell cell, string expression)
            {
                //if cell uses itself
                if (Regex.IsMatch(expression, "(^" + cell.Coordinates.Name + "[+,-,*,/])|([+,-,*,/]" + cell.Coordinates.Name + "[+,-,*,/])|([+,-,*,/]" + cell.Coordinates.Name + "$)|(^" + cell.Coordinates.Name + "$)"))
                    return true;

                //all cells that used this cell
                List<CellCoordinates> all = new List<CellCoordinates>();
                all.AddRange(cell.usedToСells);
                int Count = all.Count;

                for (int i = 0; i < Count; i++)
                {
                    Cell c = GetCell(all[i].Name);

                    if (c.Coordinates.Name == cell.Coordinates.Name)
                        return true;

                    all.AddRange(c.usedToСells);
                    Count = all.Count;
                }

                //check if use circularReference
                foreach (var item in all)
                {
                    if (Regex.IsMatch(expression, "(^" + item.Name + "[+,-,*,/])|([+,-,*,/]" + item.Name + "[+,-,*,/])|([+,-,*,/]" + item.Name + "$)|(^" + item.Name + "$)"))
                        return true;
                }

                return false;
            }

            private bool ExpressionValid(string expression)
            {
                string checkExpression = "^(((([A-Z]{1,}(([1-9])|([1-9][0-9]{1,})))|(([0-9]{1,})|([0-9]{1,}[/,,.][0-9]{1,})))([+,-,*,/](([A-Z]{1,}(([1-9])|([1-9][0-9]{1,})))|(([0-9]{1,})|([0-9]{1,}[/,,.][0-9]{1,})))){1,})|([A-Z]{1,}(([1-9])|([1-9][0-9]{1,}))))$";
                return Regex.IsMatch(expression, checkExpression);
            }

            private Cell GetCell(string CellName)
            {
                CellCoordinates GetCellCoordinates = cellsCoordinates.Find(x => x.Name == CellName);

                //check if cell exist
                if (GetCellCoordinates.Name != null)
                {
                    //get class cell from tag
                    Cell cell;
                    if (dataGridView1.Rows[GetCellCoordinates.Y].Cells[GetCellCoordinates.X].Tag != null && dataGridView1.Rows[GetCellCoordinates.Y].Cells[GetCellCoordinates.X].Tag is Cell)
                        cell = (Cell)dataGridView1.Rows[GetCellCoordinates.Y].Cells[GetCellCoordinates.X].Tag;

                    //or create new 
                    else
                    {
                        cell = new Cell(GetCellCoordinates);
                        dataGridView1.Rows[GetCellCoordinates.Y].Cells[GetCellCoordinates.X].Tag = cell;
                    }

                    return cell;
                }

                return null;
            }

            public string GetCellValue(string CellName)
            {
                string result = null;
                Cell cell = GetCell(CellName);

                if (cell != null)
                {
                    if (cell.display != null && cell.display[0] == '#' && cell.display != "")
                        result = "#Cell " + CellName + " return exception";

                    else
                        result = cell.display;
                }
                else
                {
                    result = "#Cell " + CellName + " not find";
                }

                return result;
            }

            private void CellAddFieldUsedCell(ref Cell cell, string usedCellName)
            {
                CellCoordinates UsedCell = cell.usedСells.Find(x => x.Name == usedCellName);
                //check if UsedCell already added
                if (UsedCell.Name == null)
                {
                    CellCoordinates FindUsedCell = cellsCoordinates.Find(x => x.Name == usedCellName);
                    cell.usedСells.Add(FindUsedCell);

                    //find used cell and add usedToCell
                    Cell findCell = GetCell(usedCellName);
                    if (findCell != null)
                        CellAddFieldUsedToCell(ref findCell, cell.Coordinates.Name);
                }
            }

            private void CellClearUsedCells(ref Cell cell)
            {
                foreach (var item in cell.usedСells)
                {
                    Cell findCell = GetCell(item.Name);
                    if (findCell != null)
                        CellRemoveUsedToCells(ref findCell, cell.Coordinates.Name);
                }
                cell.usedСells.Clear();
            }

            private void CellAddFieldUsedToCell(ref Cell cell, string usedToCellName)
            {
                Cell UsedToCell = GetCell(usedToCellName);

                // check if UsedToCell already added
                if (UsedToCell.Coordinates.Name != null)
                    cell.usedToСells.Add(UsedToCell.Coordinates);

            }

            private void CellRemoveUsedToCells(ref Cell cell, string usedToCellName)
            {
                Cell UsedToCell = GetCell(usedToCellName);
                if (UsedToCell.Coordinates.Name != null)
                    cell.usedToСells.Remove(UsedToCell.Coordinates);
            }

            private class Cell
            {
                private CellCoordinates coordinates;

                public bool ExpressionValid;
                public bool CircularReference;

                public string expression;
                public string display;

                //cells used to this cell
                public List<CellCoordinates> usedToСells;

                //this cell used cells
                public List<CellCoordinates> usedСells;

                public CellCoordinates Coordinates
                {
                    get
                    {
                        return coordinates;
                    }
                }

                public Cell(CellCoordinates coordinates)
                {

                    this.coordinates = coordinates;
                    this.expression = null;
                    this.display = null;
                    usedToСells = new List<CellCoordinates>();
                    usedСells = new List<CellCoordinates>();
                }

            }

    }
}
