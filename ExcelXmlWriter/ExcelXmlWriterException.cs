using System;

namespace ExcelXmlWriter
{
    /// <summary>
    /// Base exception for all ExcelXmlWriter Library exceptions.
    /// </summary>
    public class ExcelXmlWriterException : Exception
    {
        #region Construction/Finalization
        /// <summary>
        /// Initializes the empty instance of this exception.
        /// </summary>
        public ExcelXmlWriterException()
        {
        }

        /// <summary>
        /// Initializes the instance of this class with exception message.
        /// </summary>
        /// <param name="message">Exception message text.</param>
        public ExcelXmlWriterException(String message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes the instance of this class with exception message and inner exception reference.
        /// </summary>
        /// <param name="message">Exception message text.</param>
        /// <param name="innerException">Exception message text.</param>
        public ExcelXmlWriterException(String message, Exception innerException)
            : base(message, innerException)
        {
        }
        #endregion
    }
}