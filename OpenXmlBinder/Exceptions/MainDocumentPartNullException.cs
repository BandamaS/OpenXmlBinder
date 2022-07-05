using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlBinder.Exceptions
{
    internal class MainDocumentPartNullException : Exception
    {
        public MainDocumentPartNullException()
        {
        }

        public MainDocumentPartNullException(string? message) : base(message)
        {
        }

        public MainDocumentPartNullException(string? message, Exception? innerException) : base(message, innerException)
        {
        }

        protected MainDocumentPartNullException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
