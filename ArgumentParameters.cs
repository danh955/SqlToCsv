// <copyright file="ArgumentParameters.cs" ="">
//     Copyright (c) . All rights reserved.
// </copyright>
namespace SqlToCsv
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// The class will take a string and parse out the arguments and parameters
    /// into two collections.  The arguments must come before the parameters.
    /// Optionally it can ignore case.  If there are more then one parameters,
    /// it can ether pick the last value or append them together with a separator
    /// character.  It will try its best to parse out what’s in the string and
    /// not give any error message.
    /// </summary>
    internal class ArgumentParameters : Dictionary<string, string>
    {
        /// <summary>
        /// Initializes a new instance of the ArgumentParameters class.
        /// This will take a string and parse out the parameters and arguments.
        /// </summary>
        /// <param name="text">The text to parse</param>
        /// <param name="ignoreCase">True to ignore case character</param>
        /// <param name="valueSeparator">
        ///     Character to use to separate the duplicate values.
        ///     null=use last value
        /// </param>
        /// <param name="parameterFileKeyList">
        ///     List of keys for importing parameters.
        ///     If there is a key that matches the fileKeyList, it will read the file for more parameters.
        ///     If null, no parameter file.
        /// </param>
        public ArgumentParameters(string text, bool ignoreCase, string valueSeparator, string[] parameterFileKeyList)
            : base(ignoreCase ? StringComparer.CurrentCultureIgnoreCase : StringComparer.CurrentCulture)
        {
            this.Text = text;
            this.Arguments = new List<string>();

            int idx = 0;
            int textLength = text.Length;

            while (idx < textLength)
            {
                idx = SkipOverWhiteSpace(text, idx);

                // Do we have a parameter starter?
                if (IsParameterStarter(text[idx]))
                {
                    // Yes, its a key.
                    idx++;
                    int leftIdx = idx;

                    // Look for a parameter / value separator.
                    while (idx < textLength)
                    {
                        if (IsParameterValueSeparator(text[idx])
                            || IsParameterStarter(text[idx])
                            || char.IsWhiteSpace(text[idx]))
                        {
                            break;
                        }

                        idx++;
                    }

                    // We got the key.
                    string key = text.Substring(leftIdx, idx - leftIdx);
                    string value;

                    idx = SkipOverWhiteSpace(text, idx);

                    // Did we bump into the next parameter?
                    if ((idx >= textLength) || IsParameterStarter(text[idx]))
                    {
                        // Yes.
                        value = string.Empty;
                    }
                    else
                    {
                        // No.
                        // Did we find the separator?
                        if (IsParameterValueSeparator(text[idx]))
                        {
                            // Yes, skip over it.
                            idx++;
                            idx = SkipOverWhiteSpace(text, idx);
                        }

                        value = GetArgument(text, ref idx);
                    }

                    this.AddKeyValue(key, value, ignoreCase, valueSeparator, parameterFileKeyList);
                }
                else
                {
                    // No, its an argument.
                    this.Arguments.Add(GetArgument(text, ref idx));
                }

                idx = SkipOverWhiteSpace(text, idx);
            }
        }

        /// <summary>
        /// Gets a list of arguments.
        /// </summary>
        public IList<string> Arguments { get; private set; }

        /// <summary>
        /// Gets a text that was parsed.
        /// </summary>
        public string Text { get; private set; }

        /// <summary>
        /// Get an Argument.
        /// </summary>
        /// <param name="text">Text to scan over</param>
        /// <param name="idx">Starting index</param>
        /// <returns>Index past the end of the Argument or length of string</returns>
        private static string GetArgument(string text, ref int idx)
        {
            int textLength = text.Length;

            // Are we at the end of the string?
            if (idx >= textLength)
            {
                // Yes.
                return string.Empty;
            }

            int startIdx = idx;
            char quotationChar = text[idx];

            // Is this the start of a Quotation?
            if ((quotationChar == '"') || (quotationChar == '\''))
            {
                // Yes.
                startIdx++;
                idx++;

                int diffCount = 0;

                while (idx < textLength)
                {
                    // Did we find a quote character?
                    if (text[idx] == quotationChar)
                    {
                        // Yes.
                        idx++;

                        // Are there two quotes in a row?
                        if ((idx >= textLength) || (text[idx] != quotationChar))
                        {
                            // No, we are at the end of the Quotation.
                            diffCount = -1;
                            break;
                        }
                    }

                    idx++;
                }

                string quote = quotationChar.ToString();
                return text.Substring(startIdx, idx - startIdx + diffCount).Replace(quote + quote, quote);
            }
            else
            {
                // No, skip until we find a separator characters.
                while (idx < textLength)
                {
                    if (char.IsSeparator(text, idx))
                    {
                        break;
                    }

                    idx++;
                }

                return text.Substring(startIdx, idx - startIdx);
            }
        }

        /// <summary>
        /// This tests for the parameter starter character.
        /// </summary>
        /// <param name="value">Value to test.</param>
        /// <returns>True if a starter character.</returns>
        private static bool IsParameterStarter(char value)
        {
            return (value == '-') || (value == '/');
        }

        /// <summary>
        /// This tests for the parameter value separator character.
        /// </summary>
        /// <param name="value">Value to test.</param>
        /// <returns>True if a separator character.</returns>
        private static bool IsParameterValueSeparator(char value)
        {
            return (value == ':') || (value == '=');
        }

        /// <summary>
        /// Skip over the white space characters.
        /// </summary>
        /// <param name="text">Text to skip over</param>
        /// <param name="idx">Starting index</param>
        /// <returns>Index of non white space character or length of string</returns>
        private static int SkipOverWhiteSpace(string text, int idx)
        {
            int textLength = text.Length;
            while (idx < textLength)
            {
                if (char.IsWhiteSpace(text, idx) == false)
                {
                    break;
                }

                idx++;
            }

            return idx;
        }

        /// <summary>
        /// Add Key Value
        /// </summary>
        /// <param name="key">Key name to add.</param>
        /// <param name="value">Value to add.</param>
        /// <param name="ignoreCase">True if to ignore case.</param>
        /// <param name="valueSeparator">Value separator.</param>
        /// <param name="parameterFileKeyList">Parameter file key list.</param>
        private void AddKeyValue(string key, string value, bool ignoreCase, string valueSeparator, string[] parameterFileKeyList)
        {
            // Is this a duplicate key?
            if (this.ContainsKey(key))
            {
                // Yes.
                // Do we have a value separator?
                if (valueSeparator != null)
                {
                    // Yes, append the new value.
                    this[key] = this[key]
                                         + valueSeparator
                                         + value.Replace(valueSeparator, string.Empty);
                }
                else
                {
                    // No, replace it with the new value.
                    this[key] = value;
                }
            }
            else
            {
                // No.
                bool found = false;

                // Do we have a parameter file key list?
                if (parameterFileKeyList != null)
                {
                    // Yes.
                    foreach (string parameterFileKey in parameterFileKeyList)
                    {
                        // Did we find it?
                        if (string.Compare(key, parameterFileKey, ignoreCase) == 0)
                        {
                            // Yes.
                            found = true;
                            break;
                        }
                    }
                }

                // Did we find a parameter file key?
                if (found)
                {
                    // Yes.
                    this.ProcessParameterFile(value, ignoreCase, valueSeparator, parameterFileKeyList);
                }
                else
                {
                    // No.
                    this.Add(key, value);
                }
            }
        }

        /// <summary>
        /// Process Parameter File
        /// </summary>
        /// <param name="fileName">Parameter file to process.</param>
        /// <param name="ignoreCase">Ignore case.</param>
        /// <param name="valueSeparator">Value separator.</param>
        /// <param name="parameterFileKeyList">Parameter file key list.</param>
        private void ProcessParameterFile(string fileName, bool ignoreCase, string valueSeparator, string[] parameterFileKeyList)
        {
            using (System.IO.StreamReader reader = new System.IO.StreamReader(fileName))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    int idx = SkipOverWhiteSpace(line, 0);

                    // Is this a comment?
                    if ((idx < line.Length) && (line[idx] != '#'))
                    {
                        // No.
                        int keyIdx = idx;

                        // Find other end of key.
                        while ((idx < line.Length)
                            && (char.IsWhiteSpace(line[idx]) == false)
                            && (line[idx] != '='))
                        {
                            idx++;
                        }

                        string key = line.Substring(keyIdx, idx - keyIdx);
                        idx = SkipOverWhiteSpace(line, idx);

                        // Do we have a value?
                        if ((idx < line.Length) && (line[idx] == '='))
                        {
                            // Yes.
                            string value = line.Substring(idx + 1);
                            this.AddKeyValue(key, value.Trim(), ignoreCase, valueSeparator, parameterFileKeyList);
                        }
                        else
                        {
                            // Are we at the end?
                            if (idx < line.Length)
                            {
                                // Yes.
                                this.AddKeyValue(key, string.Empty, ignoreCase, valueSeparator, parameterFileKeyList);
                            }
                            else
                            {
                                // No, just use the rest of the line as the value.
                                string value = line.Substring(idx);
                                this.AddKeyValue(key, value.Trim(), ignoreCase, valueSeparator, parameterFileKeyList);
                            }
                        }
                    }
                }
            }
        }
    }
}