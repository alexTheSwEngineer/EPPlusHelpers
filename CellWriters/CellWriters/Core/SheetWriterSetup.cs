using CellWriter.EPPlusHelpers.Excell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace CellWriters.Core
{
    public class SheetWriterSetup : ISettings
    {
        public SheetWriter Writer { get; private set; }

        public SheetWriterSetup(SheetWriter sheetWriter)
        {
            Writer = sheetWriter;
        }

        public IEnumerable<Action<ExcelRange>> Modifiers
        {
            get { return Writer.Settings.Modifiers;}
        }

        public ISettings ApplyTo(ExcelRange cell)
        {
            Writer.Settings.ApplyTo(cell);
            return this;
        }

        public SheetWriterSetup Clear()
        {
            Writer.Settings = Settings.Empty;
            return this;
        }

        public ISettings With(ISettings other)
        {
            Writer.Settings = Writer.Settings.With(other);
            return this;
        }

        public ISettings With(Action<ExcelRange> modifier)
        {
            Writer.Settings = Writer.Settings.With(modifier);
            return this;
        }
    }
}
