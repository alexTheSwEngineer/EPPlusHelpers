using CellWriters;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CellWriter.EPPlusHelpers.Excell
{
    public class SheetWriter : ISheetWriter
    {
        private static object NullObject = null;
        private IEnumerable<Action<ExcelRange>> Modifiers { get; set; }
        public bool AllowMultipleWriterHanldes { get; private set; }
        public bool ShouldFormatEmptyCells { get; private set; }
        public bool ShouldPadRows { get; private set; }
        public int Column { get; private set; }
        public int Row { get; private set; }
        public ExcelWorksheet Sheet { get; private set; }
        public ExcelRange Cells {
            get
            {
                return Sheet.Cells;
            }
        }

        public SheetWriter(ExcelWorksheet sheet, bool multiHandle = false,bool padRows=false, bool formatEmpyCells=false)
        {
            AllowMultipleWriterHanldes = multiHandle;
            ShouldFormatEmptyCells = formatEmpyCells;
            ShouldPadRows = padRows;
            Sheet = sheet;
            Modifiers = new List<Action<ExcelRange>>();

            UpdatePointers();
        }

        #region ISheetWriterApi

        public SheetWriterMonad AsMonad()
        {
            return new SheetWriterMonad(this);
        }

        public void Write(params object[] values)
        {
            if (AllowMultipleWriterHanldes)
            {
                UpdatePointers();
            }

            //Empty cell
            if (values == null || values.Length == 0)
            {
                WriteCellFormated("EMP",ShouldFormatEmptyCells);//TODO remove this
                return;
            }

            //Non empty cell
            foreach (var value in values)
            {
                WriteCellFormated(value,true);
            }
        }


        public void WriteLine(params object[] values)
        {
            if (values.Empty())
            {
                UpdatePointers();
            }
            else
            {
                Write(values); //this automatically updates pointers
            }

            if (ShouldPadRows)
            {
                PadRow();  
            }

            Row++;
            Column = 1;

            //Move internal pointers to new row
            if (AllowMultipleWriterHanldes)
            {
                WriteCellFormated(null, false);
                Column--;
            }
        }


        public void With(IEnumerable<Action<ExcelRange>> modifiers, Action<ISheetWriter> action)
        {
            Modifiers = modifiers;
            action(this);
            Modifiers = new List<Action<ExcelRange>>();
        }

        #endregion

        private void WriteCellFormated(object value,bool useModifiers)
        {
            var cell = Cells[Row, Column];

            if (useModifiers)
            {
                foreach (var modifier in Modifiers)
                {
                    modifier(cell);
                }
            }

            cell.Value = value;
            Column++;
        }

        private void PadRow()
        {
            var rowSize = Sheet.Dimension.End.Column;
            if (rowSize > Column)
            {
                var paddingLength = rowSize - Column+1;
                for (int i = 0; i < paddingLength; i++)
                {
                    Write();
                }
            }
        }

        private void UpdatePointers()
        {
            Row = EndRow();
            Column = EndColumn();
        }

        private int EndColumn()
        {
            var column = 1;
            var cell = Cells[Row, column];
            while (cell != null && cell.Value != null)
            {
                column++;
                cell = Cells[Row, column];
            }
            return column;
        }

        private int EndRow()
        {
            return Sheet.Dimension?.End?.Row ?? 1;
        }
    }
}