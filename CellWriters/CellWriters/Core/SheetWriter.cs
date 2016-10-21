using CellWriters.Exstensions;
using OfficeOpenXml;
using System;

namespace CellWriters.Core
{
    public class SheetWriter : ISheetWriter
    {
        /// <summary>
        /// This must be set to true if multiple sheet writers instances would be instantiated for the same <see cref="ExcelWorksheet"/>.
        /// If only one instance of a <see cref="SheetWriter"/> exist per <see cref="ExcelWorksheet"/> then you should set <see cref="AllowMultipleWriterHanldes"/>
        /// to false for performance reasons. It will work correctly either way in the case of a single <see cref="SheetWriter"/> .
        /// </summary>
        public bool AllowMultipleWriterHanldes { get; private set; }

        /// <summary>
        /// Indicates whether parameterless calls to <see cref="Write(object[])"/> and <see cref="WriteLine(object[])"/>
        /// will apply modifiers to the cells, or just move the pointer to the next cell/row.
        /// </summary>
        public bool ApplyModifiersToEmptyCells { get; private set; }

        /// <summary>
        /// Indicates whether parameterless calls to <see cref="WriteLine(object[])"/>
        /// will insert padding if the current row does not span to the maximum number of cells used in the previous rows.
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

        public ISettings Settings
        {
            get { return _settings; }
            internal set { _settings = value ?? Core.Settings.Empty; }
        }

        private ExcelWorksheet Sheet { get; set; }
        private ISettings _settings;
        private ISettings _currentSettings;
        private ExcelRange Cells {get { return Sheet.Cells;}}

        public SheetWriter(ExcelWorksheet sheet,bool multiHandle = false,bool padRows=false, bool formatEmpyCells=false)
        {
            AllowMultipleWriterHanldes = multiHandle;
            ApplyModifiersToEmptyCells = formatEmpyCells;
            ShouldPadRows = padRows;
            Sheet = sheet;
            Settings = Core.Settings.Empty;
            _currentSettings = Core.Settings.Empty;

            UpdatePointers();
        }

        #region ISheetWriterApi

        /// <summary>
        /// Creates a <see cref="SheetWriterSetup"/> for this sheet writer.
        /// </summary>
        /// <returns></returns>
        public SheetWriterSetup SetUp()
        {
            return new SheetWriterSetup(this);
        }

        /// <summary>
        /// Executes <paramref name="action"/> and applies only <paramref name="tempSettings"/> to all <see cref="Write(object[])"/>/<see cref="WriteLine(object[])"/> calls in <paramref name="action"/>.
        /// This settings override any existing settings of this SheetWriter.
        /// </summary>
        /// <param name="tempSettings"></param>
        /// <param name="action"></param>
        public ISheetWriter WithOverrideSettings(ISettings tempSettings, Action<ISheetWriter> action)
        {
            var preservedSetting = Settings;
            Settings = tempSettings;
            action(this);
            Settings = preservedSetting;
            return this;
        }

        /// <summary>
        /// Executes <paramref name="action"/> and applies both <paramref name="addedSettings"/> and any preexisting <see cref="Settings"/> to all  <see cref="Write(object[])"/>/<see cref="WriteLine(object[])"/>  calls in <paramref name="action"/>.
        /// <paramref name="addedSettings"/> are merged with preexisting <see cref="Settings"/>.
        /// </summary>
        /// <param name="addedSettings"></param>
        /// <param name="action"></param>
        public ISheetWriter WithAddedSettings(ISettings addedSettings, Action<ISheetWriter> action)
        {
            var preservedSetting = Settings;
            Settings = Settings.With(addedSettings);
            action(this);
            _currentSettings = preservedSetting;
            return this;
        }

        /// <summary>
        /// Writes each of the objects in a separate cell in the same row. 
        /// If called with no parameters, it will write a null value to the cell and move the pointer to the next one. 
        /// Settings will be applied to such (empty) cells if <see cref="ApplyModifiersToEmptyCells"/> is set to true.
        /// </summary>
        /// <param name="values"></param>
        public ISheetWriter Write(params object[] values)
        {
            if (AllowMultipleWriterHanldes)
            {
                UpdatePointers();
            }

            //Empty cell
            if (values.Empty())
            {
                WriteCellFormated("EMP",ApplyModifiersToEmptyCells);//TODO remove this
                return this; 
            }

            //Non empty cell
            foreach (var value in values)
            {
                WriteCellFormated(value,true);
            }

            return this;
        }

        /// <summary>
        /// Writes all the objects in a separate cell and moves the pointer to the next row.
        /// </summary>
        /// <param name="values"></param>
        public ISheetWriter WriteLine(params object[] values)
        {
            if (values.Any())
            {
                Write(values); //this automatically updates pointers
            }
            else
            {
                UpdatePointers();
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

            return this;
        }
        
        #endregion

        private void WriteCellFormated(object value,bool useModifiers)
        {
            var cell = Cells[Row, Column];

            if (useModifiers)
            {
                Settings.ApplyTo(cell);
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