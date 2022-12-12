using System;
using System.Collections.Generic;

namespace DCOM.Global
{
    public class StringEqualNullEmptyComparer : IEqualityComparer<string>
    {
        /// <inheritdoc />
        public bool Equals(string x, string y)
        {
            if (string.IsNullOrEmpty(x) && string.IsNullOrEmpty(y)) return true;
            if (ReferenceEquals(null, x) && !ReferenceEquals(null, y)) return false;
            if (ReferenceEquals(null, y) && !ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(x, y)) return true;
            return string.Equals(x, y);
        }

        /// <inheritdoc cref="string.Equals(string, string, StringComparison)"/>
        public bool Equals(string x, string y, StringComparison comparisonType)
        {
            if (string.IsNullOrEmpty(x) && string.IsNullOrEmpty(y)) return true;
            if (ReferenceEquals(null, x) && !ReferenceEquals(null, y)) return false;
            if (ReferenceEquals(null, y) && !ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(x, y)) return true;
            return string.Equals(x, y, comparisonType);
        }

        /// <inheritdoc />
        public int GetHashCode(string obj)
        {
            unchecked
            {
                int hash = 13;
                hash = (hash * 7) ^ (obj.GetHashCode());
                return hash;
            }
        }
    }
}