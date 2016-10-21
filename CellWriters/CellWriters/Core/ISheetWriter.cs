using System;

namespace CellWriters.Core
{
    public interface ISheetWriter
    {
        /// <summary>
        /// Writes all objects to a separate cell. Moves the cursor to the next one. Subsequent calls will write in the same row.
        /// DefaultSettings are applied.
        /// </summary>
        /// <param name="values"></param>
        ISheetWriter Write(params object[] values);

        /// <summary>
        /// Writes all objects to a separate cell in the same row as the cursor and then moves it to the next row.
        /// DefaultSettings are applied.
        /// </summary>
        /// <param name="values"></param>
        ISheetWriter WriteLine(params object[] values);

        /// <summary>
        /// Applies only this settings to all write/writeLine calls within the action.
        /// This settings override the default settings for this SheetWriter, for this WithTempSettings call.
        /// </summary>
        /// <param name="cellModifier"></param>
        /// <param name="action"></param>
        ISheetWriter WithOverrideSettings(ISettings settings, Action<ISheetWriter> action);

        /// <summary>
        /// Applies the <param name="settings"/> together with the existing settings to all write/writeLine calls within the <paramref name="action"/>.
        /// </summary>
        /// <param name="cellModifier"></param>
        /// <param name="action"></param>
        ISheetWriter WithAddedSettings(ISettings settings, Action<ISheetWriter> action);

        /// <summary>
        /// Creates a <see cref="SheetWriterSetup"/>monad for this ISheetWriter
        /// </summary>
        /// <returns></returns>
        SheetWriterSetup SetUp();
    }
}