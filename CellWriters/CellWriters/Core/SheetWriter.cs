using CellWriters;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CellWriter.EPPlusHelpers.Excell
{
    public class SheetWriter : ISheetWriter
    {
        private static object NullObject = null;


        /// <summary>
        /// This must be set to true if multiple sheet writers instances would be instantiated for the same <see cref="ExcelWorksheet"/>.
        /// If only one instance of a <see cref="SheetWriter"/> exist per <see cref="ExcelWorksheet"/> then you should set <see cref="AllowMultipleWriterHanldes"/>
        /// to false for performance reasons. It will work corectly either way in the case of a single <see cref="SheetWriter"/> .
        /// </summary>
        public bool AllowMultipleWriterHanldes { get; private set; }

        /// <summary>
        /// Indicates whether parametarless calls to <see cref="Write(object[])"/> and <see cref="WriteLine(object[])"/>
        /// will apply modifiers to the cells, or just move the pointer to the next cell/row.
        /// </summary>
        public bool ApplyModifiersToEmptyCells { get; private set; }

        /// <summary>
        /// Indicates whether parametarless calls to <see cref="WriteLine(object[])"/>
        /// will insert pading if the current row does not spann to the maximum number of cells used in the previous rows.
        /// Padding cells will be modified according to <see cref="ApplyModifiersToEmptyCells"/>
        /// </summary>
        public bool ShouldPadRows { get; private set; }

        /// <summary>
        /// The column index of the cell that is next in line to get written to.
        /// </summary>
        public int Column { get; private set; }

        /// <summary>
        /// The Row index of the cell that is next in line to get written to.
        /// </summary>
        public int Row { get; private set; }

        
        private ExcelWorksheet Sheet { get; set; }
        private ExcelRange Cells {
            get
            {
                return Sheet.Cells;
            }
        }
        private IEnumerable<Action<ExcelRange>> CurrentModifiers { get; set; }

        public SheetWriter(ExcelWorksheet sheet,bool multiHandle = false,bool padRows=false, bool formatEmpyCells=false)
        {
            AllowMultipleWriterHanldes = multiHandle;
            ApplyModifiersToEmptyCells = formatEmpyCells;
            ShouldPadRows = padRows;
            Sheet = sheet;
            CurrentModifiers = new List<Action<ExcelRange>>();

            UpdatePointers();
        }

        #region ISheetWriterApi

        /// <summary>
        /// Creates a <see cref="SheetWriterMonad"/> for this sheet writer.
        /// </summary>
        /// <returns></returns>
        public SheetWriterMonad AsMonad()
        {
            return new SheetWriterMonad(this);
        }

        /// <summary>
        /// Sets up <see cref="Write(object[])"/> and <see cref="WriteLine(object[])"/> to apply all <paramref name="modifiers"/>
        /// to the cells being written to. Executes <paramref name="action"/>. The modifiers are applied for this call only.
        /// Subsequent calls will not have nay of the previous modifiers.
        /// </summary>
        /// <param name="modifiers"></param>
        /// <param name="action"></param>
        public void With(IEnumerable<Action<ExcelRange>> modifiers, Action<ISheetWriter> action)
        {
            CurrentModifiers = modifiers;
            action(this);
            CurrentModifiers = new List<Action<ExcelRange>>();
        }

        /// <summary>
        /// Writes each of the objects in a separate cell in the same row. 
        /// If called with no parametars, it will write a null value to the cell and move the pointer to the next one, 
        /// such cells will be formated if <see cref="ApplyModifiersToEmptyCells"/> is set to true.
        /// </summary>
        /// <param name="values"></param>
        public void Write(params object[] values)
        {
            if (AllowMultipleWriterHanldes)
            {
                UpdatePointers();
            }

            //Empty cell
            if (values == null || values.Length == 0)
            {
                WriteCellFormated("EMP",ApplyModifiersToEmptyCells);//TODO remove this
                return;
            }

            //Non empty cell
            foreach (var value in values)
            {
                WriteCellFormated(value,true);
            }
        }

        /// <summary>
        /// Writes all the objects in a separate cell and moves the pointer to the next row.
        /// </summary>
        /// <param name="values"></param>
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
        
        #endregion

        private void WriteCellFormated(object value,bool useModifiers)
        {
            var cell = Cells[Row, Column];

            if (useModifiers)
            {
                foreach (var modifier in CurrentModifiers)
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