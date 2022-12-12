using System;
using COMAdmin;
using DCOM.Global;

namespace WinDevOps
{
    /// <summary>
    /// Contains a list of the protocols to be used by DCOM
    /// </summary>
    public class DCOMProtocol : IEquatable<DCOMProtocol>, IComparable<DCOMProtocol>
    {
        /// <summary>
        /// The name of the protocol. This property is returned when the Name property method is called on an object of this collection.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// The order in which to try the protocol.
        /// </summary>
        public int Order { get; }

        /// <summary>
        /// The code specifying the RPC protocol sequence. The supported protocol codes include the following:
        /// ncacn_ip_tcp, ncacn_http, ncacn_spx.
        ///
        /// This property is returned when the Key property method is called on an object of this collection.
        /// </summary>
        public string ProtocolCode { get; }

        internal DCOMProtocol(COMAdminCatalogObject catalogObject)
        {
            Name = (string)catalogObject.Value["Name"];
            Order = (int)catalogObject.Value["Order"];
            ProtocolCode = (string)catalogObject.Value["ProtocolCode"];
        }


        /// <inheritdoc />
        public bool Equals(DCOMProtocol other)
        {
            StringEqualNullEmptyComparer sC = new StringEqualNullEmptyComparer(); //comparer to handle string, null, empties, foreign language
            return sC.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase) &&
                   sC.Equals(ProtocolCode, other.ProtocolCode, StringComparison.OrdinalIgnoreCase) &&
                   Order == other.Order;
        }

        /// <inheritdoc />
        public int CompareTo(DCOMProtocol other)
        {
            if (Order > other.Order) return 1;
            if (Order < other.Order) return -1;
            return 0;
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 13;
                hash = (hash * 7) ^ Name.GetHashCode();
                hash = (hash * 7) ^ Order.GetHashCode();
                hash = (hash * 7) ^ ProtocolCode.GetHashCode();
                return hash;
            }
        }
    }
}