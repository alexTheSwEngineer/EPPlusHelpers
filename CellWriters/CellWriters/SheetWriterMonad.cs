using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CellWriter.EPPlusHelpers.Excell
{
    public class SheetWriterMonad
    {
        public ISheetWriter Writer { get; private set; }
        private List<Action<ExcelRange>> Modifiers { get; set; }

        public SheetWriterMonad(SheetWriterMonad other)
        {
            Writer = other.Writer;
            Modifiers = other.Modifiers.ToList();
        }
        public SheetWriterMonad(ISheetWriter writer)
        {
            Writer = writer;
            Modifiers = new List<Action<ExcelRange>>();
        }

        public SheetWriterMonad With(Action<ExcelRange> modifier)
        {
            return new SheetWriterMonad(this).Add(modifier);
        }

        public void Write(Action<ISheetWriter> writeAction)
        {
            Writer.With(Modifiers, writeAction);
        }

        private SheetWriterMonad Add(Action<ExcelRange> modifier)
        {
            Modifiers.Add(modifier);
            return this;
        }
    }

}