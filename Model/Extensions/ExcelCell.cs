// -----------------------------------------------------------------------
// <copyright file="Cell.cs" company="Trivadis AG">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace ExcelReader.Extensions
{
    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public class ExcelCell
    {
        public int Column { get; set; }

        public int Row { get; set; }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != typeof(ExcelCell)) return false;
            return Equals((ExcelCell)obj);
        }

        public bool Equals(ExcelCell other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return other.Column == Column && other.Row == Row;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Column * 397) ^ Row;
            }
        }
    }
}