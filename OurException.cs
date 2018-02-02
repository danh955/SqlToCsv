// <copyright file="OurException.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
namespace SqlToCsv
{
    using System;

    /// <summary>
    /// Our exception class.
    /// </summary>
    public class OurException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the OurException class.
        /// </summary>
        /// <param name="message">The error message.</param>
        public OurException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the OurException class.
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="innerException">The inner exception object.</param>
        public OurException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}