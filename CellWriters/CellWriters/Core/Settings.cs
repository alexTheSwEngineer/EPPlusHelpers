using CellWriters.Exstensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CellWriters.Core
{
    public class Settings : ISettings
    {
        public List<Action<ExcelRange>> Modifiers { get; set; }

        IEnumerable<Action<ExcelRange>> ISettings.Modifiers{   get { return Modifiers.AsEnumerable(); }
        }

        public Settings(Action<ExcelRange> modifier)
        {
            Modifiers = new List<Action<ExcelRange>> { modifier};
        }

        public Settings(IEnumerable<Action<ExcelRange>> modifiers)
        {
            Modifiers = modifiers.ToList();
        }

        public Settings(Settings other) : this(other.Modifiers) { }

        public Settings()
        {            
            Modifiers = new List<Action<ExcelRange>>();
        }

        public ISettings ApplyTo(ExcelRange cell)
        {
            foreach (var modifier in Modifiers)
            {
                modifier(cell);
            }

            return this;
        }

        /// <summary>
        /// Adds a modifier with the highest priority. 
        /// </summary>
        /// <param name="modifier"></param>
        /// <returns></returns>
        public ISettings With(Action<ExcelRange> modifier)
        {
            return new Settings(this).Add(modifier.ToEnumerable());
        }

        /// <summary>
        /// Adds  modifiers with the highest priority. 
        /// </summary>
        /// <param name="modifier"></param>
        /// <returns></returns>
        public ISettings With(IEnumerable<Action<ExcelRange>> modifiers)
        {
            return new Settings(this).Add(modifiers);
        }

        /// <summary>
        /// Adds all modifiers from the other settings.
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public ISettings With(ISettings other)
        {
            return new Settings(this).Add(other.Modifiers);
        }

        private Settings Add(IEnumerable<Action<ExcelRange>> modifiers)
        {
            Modifiers.AddRange(modifiers);
            return this;
        }

        public static Settings Empty
        {
            get
            {
                return new Settings();
            }
        }
    }

}