// <copyright file="ArgumentParameters.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>

namespace SqlToCsv
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Diagnostics;
    using System.IO;
    using System.Net.Mail;
    using System.Reflection;
    using System.Text;

    /// <summary>
    /// Program class.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// Send the file as an attachment with this file name.
        /// </summary>
        private static string attachmentFileName;

        /// <summary>
        /// Zero of more email address to send the CSV file to.
        /// </summary>
        private static string bcc = null;

        /// <summary>
        /// Text that goes into the body of the email.
        /// </summary>
        private static string body = null;

        /// <summary>
        /// Zero or more email address to send the CSV file to.
        /// </summary>
        private static string cc = null;

        /// <summary>
        /// The connection string.
        /// </summary>
        private static string connectionString = null;

        /// <summary>
        /// True if to create the database table before importing the file.
        /// </summary>
        private static bool createTable = false;

        /// <summary>
        /// The database to query.
        /// </summary>
        private static string database = null;

        /// <summary>
        /// True if to send the results via email.
        /// </summary>
        private static bool doEmail = false;

        /// <summary>
        /// How to encode the output text file.
        /// </summary>
        private static Encoding encoding = Encoding.ASCII;

        /// <summary>
        /// This is the one character field separator in the CSV file. (\t = tab)
        /// </summary>
        private static char fieldSeparator = ',';

        /// <summary>
        /// The file to import or export.
        /// </summary>
        private static string file = null;

        /// <summary>
        /// From email address.
        /// </summary>
        private static string from = null;

        /// <summary>
        /// True if to include the header in the output file.
        /// </summary>
        private static bool includeHeader = true;

        /// <summary>
        /// True if this is an export operation.
        /// </summary>
        private static bool isExport = false;

        /// <summary>
        /// True if this is a import operation.
        /// </summary>
        private static bool isImport = false;

        /// <summary>
        /// The parameters.
        /// </summary>
        private static ArgumentParameters parameters;

        /// <summary>
        /// The password for the login.
        /// </summary>
        private static string password = null;

        /// <summary>
        /// The SQL query to export.
        /// </summary>
        private static string query = null;

        /// <summary>
        /// Replace oldText with newText.  (\t=tab, \n=lf, \r=cr, \:=:, \\=\)
        /// </summary>
        private static List<KeyValuePair<string, string>> replace = null;

        /// <summary>
        /// This will replace all control character with text except \r\n\t.
        /// </summary>
        private static string replaceControlCharacter = null;

        /// <summary>
        /// The server to get the data from.  Default "(local)"
        /// </summary>
        private static string server = "(local)";

        /// <summary>
        /// Is the SMTP server to send the email though.
        /// </summary>
        private static string smtpServer = null;

        /// <summary>
        /// Text of the subject line in the email.
        /// </summary>
        private static string subject = null;

        /// <summary>
        /// The table name to import.
        /// </summary>
        private static string table = null;

        /// <summary>
        /// The time in seconds to wait for the command to execute. The default is 30 seconds.
        /// </summary>
        private static int timeout = int.MinValue;

        /// <summary>
        /// One or more email address to send the CSV file to.
        /// </summary>
        private static string to = null;

        /// <summary>
        /// Trim the white space from both end of the data values.
        /// </summary>
        private static bool trimWhiteSpace = false;

        /// <summary>
        /// User name to login.
        /// </summary>
        private static string user = null;

        /// <summary>
        /// Add to replace
        /// </summary>
        /// <param name="inText">String to do replacement with.</param>
        private static void AddToReplace(string inText)
        {
            StringBuilder text = new StringBuilder();
            string oldValue = null;
            int idx = 0;
            while (idx < inText.Length)
            {
                char inChar = inText[idx];
                if (inChar == '\\')
                {
                    idx++;
                    if (idx < inText.Length)
                    {
                        inChar = inText[idx];

                        // \t=tab, \n=lf, \r=cr, \:=:, \\=\
                        switch (char.ToLower(inChar))
                        {
                            case 't':
                                {
                                    inChar = '\t';
                                    break;
                                }

                            case 'n':
                                {
                                    inChar = '\n';
                                    break;
                                }

                            case 'r':
                                {
                                    inChar = '\r';
                                    break;
                                }

                            case ':':
                                {
                                    inChar = ':';
                                    break;
                                }

                            case '\\':
                                {
                                    inChar = '\\';
                                    break;
                                }

                            default:
                                {
                                    text.Append('\\');
                                    break;
                                }
                        }
                    }

                    text.Append(inChar);
                }
                else if ((oldValue != null) || (inChar != ':'))
                {
                    text.Append(inChar);
                }
                else
                {
                    oldValue = text.ToString();
                    text.Length = 0;
                }

                idx++;
            }

            string newValue;
            if (oldValue == null)
            {
                oldValue = text.ToString();
                newValue = string.Empty;
            }
            else
            {
                newValue = text.ToString();
            }

            if (replace == null)
            {
                replace = new List<KeyValuePair<string, string>>();
            }

            replace.Add(new KeyValuePair<string, string>(oldValue, newValue));
        }

        /// <summary>
        /// ASCII only characters
        /// </summary>
        /// <param name="text">A string to convert.</param>
        /// <param name="replace">Text to replace control characters.</param>
        /// <returns>A string that has ASCII only characters.</returns>
        private static string AsciiOnly(string text, string replace)
        {
            if (replace != null)
            {
                for (int idx = 0; idx < text.Length; idx++)
                {
                    char c = text[idx];

                    // ASCII characters only (32 to 126 and 13, 10, 9)
                    if (((c < 32) && (c != 13) && (c != 10) && (c != 9)) || (c > 126))
                    {
                        text = text.Substring(0, idx) + replace + text.Substring(idx + 1);
                    }
                }
            }

            return text;
        }

        /// <summary>
        /// Occurs every time that the number of rows specified by the NotifyAfter
        /// property have been processed.
        /// </summary>
        /// <param name="sender">The senders object.</param>
        /// <param name="e">The event arguments.</param>
        private static void BulkCopy_SqlRowsCopied(object sender, SqlRowsCopiedEventArgs e)
        {
            Console.WriteLine("{0:#,##0} rows uploaded", e.RowsCopied);
        }

        /// <summary>
        /// Create the table.
        /// This will look at the CSV file and find the width of each column.
        /// It will get the column names from the first row of the file.
        /// </summary>
        /// <returns>true if created.</returns>
        private static bool CreateTable()
        {
            Debug.Assert(file != null, "The file parameter must not be null.");
            Debug.Assert(table != null, "The table parameter must not be null.");

            try
            {
                string[] fields;
                int[] counts;

                // Get all the column names and get the max width of all the values.
                using (StreamReader reader = new StreamReader(file))
                {
                    fields = CSV.Import(reader, fieldSeparator, trimWhiteSpace, null);
                    counts = new int[fields.Length];

                    KeyValuePair<string, string>[] replaceArray = replace == null ? null : replace.ToArray();

                    string[] values;
                    while ((values = CSV.Import(reader, fieldSeparator, trimWhiteSpace, replaceArray)) != null)
                    {
                        for (int idx = 0; (idx < values.Length) && (idx < fields.Length); idx++)
                        {
                            if (counts[idx] < values[idx].Length)
                            {
                                counts[idx] = values[idx].Length;
                            }
                        }
                    }
                }

                StringBuilder sql = new StringBuilder();

                // Create the SQL statement to create the table.
                sql.AppendFormat("create table [{0}] (\n", table);
                for (int idx = 0; idx < fields.Length; idx++)
                {
                    if (idx > 0)
                    {
                        sql.Append(",\n");
                    }

                    sql.AppendFormat("[{0}] nvarchar({1})", fields[idx], counts[idx] <= 0 ? 1 : counts[idx]);
                }

                sql.Append("\n)");

                // Create the table.
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    SqlCommand command = new SqlCommand(sql.ToString(), connection);
                    command.ExecuteNonQuery();
                }

                Console.WriteLine("Table created.");
            }
            catch (FileNotFoundException)
            {
                return ErrorMessage("{0} not found.", file);
            }
            catch (SqlException caught)
            {
                return ErrorMessage(caught.Message);
            }

            return true;
        }

        /// <summary>
        /// Send error message to the console.
        /// </summary>
        /// <param name="message">A composite format string.</param>
        /// <param name="args">An object array that contains zero or more objects to format.</param>
        /// <returns>A Boolean value of false.</returns>
        private static bool ErrorMessage(string message, params object[] args)
        {
            Debug.Assert(message != null, "The message parameter must not be null.");

            return ErrorMessage(string.Format(message, args));
        }

        /// <summary>
        /// Send error message to the console.
        /// </summary>
        /// <param name="message">String of the message to send to the console.</param>
        /// <returns>A false value.</returns>
        private static bool ErrorMessage(string message)
        {
            Debug.Assert(message != null, "The message parameter must not be null.");

            Console.WriteLine("\nError: {0}\n", message);
            return false;
        }

        /// <summary>
        /// Export SQL to CSV Email
        /// </summary>
        private static void ExportSqltoCsvEmail()
        {
            Debug.Assert(query != null, "The query parameter must not be null.");
            Debug.Assert(connectionString != null, "The connectionString must not be null.");
            Debug.Assert(from != null, "The from must not be null.");
            Debug.Assert(to != null, "The to must not be null.");
            Debug.Assert(subject != null, "The subject must not be null.");
            Debug.Assert(smtpServer != null, "The smtpServer must not be null.");

            Console.WriteLine("            Query: " + query);
            if (file != null)
            {
                Console.WriteLine("      Output file: " + file);
            }

            Console.WriteLine("Connection string: " + connectionString);
            Console.WriteLine("  Field separator: " + fieldSeparator.ToString());
            Console.WriteLine(" Trim white space: " + trimWhiteSpace.ToString());

            // Get the CSV text.
            string csvText;
            using (TextWriter writer = new StringWriter())
            {
                ExportSqltoCsvTextWriter(writer);
                csvText = writer.ToString();
            }

            // Are we writing it out to a file?
            if (file != null)
            {
                // Yes.
                using (StreamWriter writer = new StreamWriter(file, false, encoding))
                {
                    writer.Write(csvText);
                }
            }

            MailMessage mail = new MailMessage();
            mail.From = GetMailAddress(from, "From", true);
            SetMailAddressList(mail.To, to, "To", true);
            SetMailAddressList(mail.CC, cc, "CC", false);
            SetMailAddressList(mail.Bcc, bcc, "BCC", false);
            mail.Subject = subject;
            mail.IsBodyHtml = false;

            // Will it be attached?
            if (attachmentFileName != null)
            {
                // Yes.
                UnicodeEncoding encoding = new UnicodeEncoding();
                byte[] csvBytes = encoding.GetBytes(csvText);
                Stream stream = new MemoryStream(csvBytes);
                Attachment attachment = new Attachment(stream, attachmentFileName);
                mail.Attachments.Add(attachment);

                mail.Body = body == null ? string.Empty : body;
            }
            else
            {
                if (body == null)
                {
                    mail.Body = csvText;
                }
                else
                {
                    mail.Body = body + "\n\n" + csvText;
                }
            }

            SmtpClient smtp = new SmtpClient(smtpServer);
            smtp.Send(mail);
        }

        /// <summary>
        /// Export SQL to CSV file.
        /// </summary>
        private static void ExportSqltoCsvFile()
        {
            Debug.Assert(query != null, "The query parameter must not be null.");
            Debug.Assert(file != null, "The file parameter must not be null.");
            Debug.Assert(connectionString != null, "The connectionString parameter must not be null.");

            Console.WriteLine("            Query: " + query);
            Console.WriteLine("      Output file: " + file);
            Console.WriteLine("Connection string: " + connectionString);
            Console.WriteLine("  Field separator: " + fieldSeparator.ToString());
            Console.WriteLine(" Trim white space: " + trimWhiteSpace.ToString());

            using (TextWriter writer = new StreamWriter(file, false, encoding))
            {
                ExportSqltoCsvTextWriter(writer);
            }
        }

        /// <summary>
        /// Export SQL to a CSV TextWriter.
        /// </summary>
        /// <param name="writer">Text writer object.</param>
        private static void ExportSqltoCsvTextWriter(TextWriter writer)
        {
            Debug.Assert(writer != null, "The writer parameter must not be null.");
            Debug.Assert(query != null, "The query parameter must not be null.");
            Debug.Assert(connectionString != null, "The connectionString parameter must not be null.");

            DateTime startTime = DateTime.Now;
            int exportCount = 0;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    // Do we have a timeout value?
                    if (timeout >= 0)
                    {
                        // Yes.
                        command.CommandTimeout = timeout;
                    }

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        object[] dataValues = new object[reader.FieldCount];

                        DataTable table = reader.GetSchemaTable();
                        DataRowCollection rows = table.Rows;
                        for (int idx = 0; idx < rows.Count; idx++)
                        {
                            dataValues[idx] = rows[idx]["ColumnName"];
                        }

                        // Should we export the header?
                        if (includeHeader)
                        {
                            // Yes.
                            string text = CSV.Export(dataValues, fieldSeparator);
                            writer.WriteLine(text);
                        }

                        KeyValuePair<string, string>[] replaceArray = replace == null ? null : replace.ToArray();

                        // Export data.
                        while (reader.Read())
                        {
                            reader.GetProviderSpecificValues(dataValues);
                            string text = CSV.Export(dataValues, fieldSeparator, trimWhiteSpace, replaceArray);
                            writer.WriteLine(AsciiOnly(text, replaceControlCharacter));
                            exportCount++;
                        }
                    }
                }
            }

            Console.WriteLine("Exported {0:#,##0} rows in{1}.\n", exportCount, GetDuration(DateTime.Now - startTime));
        }

        /// <summary>
        /// Get Duration text.
        /// </summary>
        /// <param name="span">A time span to convert.</param>
        /// <returns>String of duration.</returns>
        private static string GetDuration(TimeSpan span)
        {
            StringBuilder text = new StringBuilder();
            if (span.TotalDays >= 1)
            {
                text.AppendFormat(" {0:#,##0} day{1}", span.TotalDays, span.TotalDays == 1 ? string.Empty : "s");
            }

            if (span.Hours > 0)
            {
                text.AppendFormat(" {0} hour{1}", span.Hours, span.Hours == 1 ? string.Empty : "s");
            }

            if (span.Minutes > 0)
            {
                text.AppendFormat(" {0} minute{1}", span.Minutes, span.Minutes == 1 ? string.Empty : "s");
            }

            if ((span.TotalSeconds < 60) && (span.Milliseconds != 0))
            {
                text.AppendFormat(" {0:0.0##} seconds", ((double)span.TotalMilliseconds) / 1000);
            }
            else
            {
                if (span.Seconds > 0)
                {
                    text.AppendFormat(" {0} second{1}", span.Seconds, span.Seconds == 1 ? string.Empty : "s");
                }
            }

            return text.ToString();
        }

        /// <summary>
        /// Parse the email address out of the string.
        /// </summary>
        /// <param name="emailAddress">A single email address.</param>
        /// <param name="fieldName">Field name to display if there is an error.</param>
        /// <param name="required">True if email address is required.</param>
        /// <returns>Mail address object otherwise null if not found</returns>
        private static MailAddress GetMailAddress(string emailAddress, string fieldName, bool required)
        {
            if ((emailAddress == null) || (emailAddress.Trim() == string.Empty))
            {
                if (required)
                {
                    throw new OurException(string.Format("\"{0}\" requires an email address", fieldName));
                }
                else
                {
                    return null;
                }
            }

            // Extract the name if any.
            string name = null;
            emailAddress = emailAddress.Trim();

            if (emailAddress.EndsWith(")"))
            {
                int leftIdx = emailAddress.LastIndexOf('(');
                if (leftIdx > 0)
                {
                    name = emailAddress.Substring(leftIdx + 1, emailAddress.Length - leftIdx - 2).Trim();
                    emailAddress = emailAddress.Substring(0, leftIdx).Trim();
                }
            }

            // Is there an email address?
            if (emailAddress == string.Empty)
            {
                // No.
                return null;
            }

            try
            {
                // Do we have a name?
                if (name == null)
                {
                    // No.
                    return new MailAddress(emailAddress);
                }
                else
                {
                    // Yes.
                    return new MailAddress(emailAddress, name);
                }
            }
            catch (FormatException)
            {
                throw new OurException(string.Format("\"{0}\" has an invalid email address of \"{1}\"", fieldName, emailAddress));
            }
        }

        /// <summary>
        /// Import CSV file to SQL.
        /// </summary>
        private static void ImportCsvFileToSql()
        {
            Debug.Assert(file != null, "The file parameter must not be null.");
            Debug.Assert(table != null, "The table parameter must not be null.");
            Debug.Assert(connectionString != null, "The connectionString parameter must not be null.");

            Console.WriteLine("             File: " + file);
            Console.WriteLine("     Output table: " + table);
            Console.WriteLine("Connection string: " + connectionString);
            Console.WriteLine("  Field separator: " + fieldSeparator.ToString());
            Console.WriteLine(" Trim white space: " + trimWhiteSpace.ToString());

            // Create the table?
            if (createTable)
            {
                // Yes.
                if (CreateTable() == false)
                {
                    return;
                }
            }

            KeyValuePair<string, string>[] replaceArray = replace == null ? null : replace.ToArray();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    using (StreamReader reader = new StreamReader(file))
                    {
                        using (ImportDataReader importReader
                            = new ImportDataReader(reader, fieldSeparator, trimWhiteSpace, replaceArray))
                        {
                            bulkCopy.DestinationTableName = table;
                            bulkCopy.BulkCopyTimeout = 1000;
                            bulkCopy.NotifyAfter = 100;
                            bulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(BulkCopy_SqlRowsCopied);
                            bulkCopy.WriteToServer(importReader);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Load the arguments and parameters.
        /// </summary>
        /// <returns>True if successful</returns>
        private static bool LoadArgumentParameters()
        {
            string[] parameterFileKeyList = { "PF", "paramFile", "parameterFile" };
            parameters = new ArgumentParameters(System.Environment.CommandLine, true, null, parameterFileKeyList);

            // Anything to work with?
            if ((parameters.Arguments.Count <= 1) && (parameters.Count == 0))
            {
                // No.
                // Display the help and exit.
                Console.WriteLine(Properties.Resources.Help);
                return false;
            }

            // Do we have the first argument?
            if (parameters.Arguments.Count > 1)
            {
                // Yes.
                // Is this an import?
                if (parameters.ContainsKey("Import"))
                {
                    // Yes.
                    file = parameters.Arguments[1];
                }
                else
                {
                    // No.
                    query = parameters.Arguments[1];
                }
            }

            // Do we have the second argument?
            if (parameters.Arguments.Count > 2)
            {
                // Yes.
                // Is this an import?
                if (parameters.ContainsKey("Import"))
                {
                    // Yes.
                    table = parameters.Arguments[2];
                }
                else
                {
                    // No.
                    file = parameters.Arguments[2];
                }
            }

            // Do we have too many arguments?
            if (parameters.Arguments.Count > 3)
            {
                // Yes.
                return ErrorMessage("Too many arguments.");
            }

            // Loop and get all the parameters.
            foreach (KeyValuePair<string, string> item in parameters)
            {
                switch (item.Key.Trim().ToLower())
                {
                    case "attachment":  // Attachment
                    case "attachmentfile":  // Attachment file
                    case "attachmentfilename":  // Attachment file name
                        {
                            attachmentFileName = item.Value;
                            break;
                        }

                    case "bcc":  // BCC
                        {
                            bcc = item.Value;
                            break;
                        }

                    case "body":  // Body
                        {
                            body = item.Value;
                            break;
                        }

                    case "cc":  // CC
                        {
                            cc = item.Value;
                            break;
                        }

                    case "ct":
                    case "create":
                    case "createtable":  // create table
                        {
                            createTable = true;
                            break;
                        }

                    case "d":
                    case "database":
                        {
                            database = item.Value;
                            break;
                        }

                    case "e":
                    case "encoding":
                        {
                            const string EncodingErrorMessage =
                                "Invalid encoding, must be \"ASCII\", \"Unicode\","
                                + "\"UTF7\", \"UTF8\", \"UTF16\" or \"UTF32\"";

                            string encodingText = item.Value;
                            if (encodingText == null)
                            {
                                return ErrorMessage(EncodingErrorMessage);
                            }

                            switch (encodingText.Trim().ToLower())
                            {
                                // ASCII
                                case "ascii":
                                    encoding = Encoding.ASCII;
                                    break;

                                // Unicode
                                case "unicode":
                                    encoding = Encoding.Unicode;
                                    break;

                                // UTF7
                                case "utf7":
                                    encoding = Encoding.UTF7;
                                    break;

                                // UTF8
                                case "utf8":
                                    encoding = Encoding.UTF8;
                                    break;

                                // UTF16
                                case "utf16":
                                    encoding = Encoding.Unicode;
                                    break;

                                // UTF32
                                case "utf32":
                                    encoding = Encoding.UTF32;
                                    break;

                                default:
                                    return ErrorMessage(EncodingErrorMessage);
                            }

                            break;
                        }

                    case "eh":
                    case "excludeheader":  // ExcludeHeader
                        {
                            includeHeader = false;
                            break;
                        }

                    case "export":
                        {
                            isExport = true;
                            break;
                        }

                    case "fs":
                    case "fieldseparator":  // field separator
                        {
                            string text = item.Value.Trim();
                            if (text == @"\t")
                            {
                                fieldSeparator = (char)9;
                                break;
                            }

                            if (text.Length != 1)
                            {
                                return ErrorMessage("The field separator can only be one character or \\t");
                            }

                            fieldSeparator = text[0];
                            break;
                        }

                    case "f":
                    case "file":
                        {
                            file = item.Value;
                            break;
                        }

                    case "from":  // From
                        {
                            from = item.Value;
                            break;
                        }

                    case "h":
                    case "help":
                    case "?":
                        {
                            // Display the help and exit.
                            Console.WriteLine(Properties.Resources.Help);
                            return false;
                        }

                    case "import":
                        {
                            isImport = true;
                            break;
                        }

                    case "lf":  // Replace LF
                        {
                            AddToReplace("\n:" + item.Value);
                            break;
                        }

                    case "p":
                    case "password":
                        {
                            password = item.Value;
                            break;
                        }

                    case "q":
                    case "query":
                        {
                            query = item.Value;
                            break;
                        }

                    case "r":
                    case "replace":  // Replace
                        {
                            AddToReplace(item.Value);
                            break;
                        }

                    case "rcc":  // Replace Control Character
                    case "replacecontrolcharacter":
                        {
                            replaceControlCharacter = item.Value;
                            break;
                        }

                    case "cr":  // Replace CR
                        {
                            AddToReplace("\r:" + item.Value);
                            break;
                        }

                    case "s":
                    case "server":
                        {
                            server = item.Value;
                            break;
                        }

                    case "smtpserver":  // SmtpServer
                        {
                            smtpServer = item.Value;
                            break;
                        }

                    case "subject":  // Subject
                        {
                            subject = item.Value;
                            break;
                        }

                    case "t":
                    case "timeout":
                        {
                            if (int.TryParse(item.Value, out timeout) == false)
                            {
                                return ErrorMessage("Invalid timeout value.");
                            }

                            if (timeout < 0)
                            {
                                return ErrorMessage("Timeout must be a positive value.");
                            }

                            break;
                        }

                    case "table":
                        {
                            table = item.Value;
                            break;
                        }

                    case "to":  // To
                        {
                            to = item.Value;
                            break;
                        }

                    case "tws":
                    case "trimwhitespace":  // trim white space
                        {
                            trimWhiteSpace = true;
                            break;
                        }

                    case "u":
                    case "user":
                        {
                            user = item.Value;
                            break;
                        }

                    default:
                        {
                            return ErrorMessage("Invalid parameter: {0}", item.Key);
                        }
                }
            }

            StringBuilder newConnectionString = new StringBuilder();
            newConnectionString.AppendFormat("Server={0}", server);

            // Do we have a database?
            if (database != null)
            {
                // Yes.
                newConnectionString.AppendFormat(";Database={0}", database);
            }

            // Do we have the user name?
            if (user != null)
            {
                // Yes.
                newConnectionString.AppendFormat(";UID={0}", user);
                if (password != null)
                {
                    newConnectionString.AppendFormat(";Password={0}", password);
                }
            }
            else
            {
                newConnectionString.Append(";Trusted_Connection=SSPI");
            }

            connectionString = newConnectionString.ToString();

            // If any of the email stuff is set, its to be emailed.
            doEmail = from != null || to != null || cc != null || bcc != null || subject != null
                || body != null || attachmentFileName != null || smtpServer != null;

            // Is this an import operation?
            if (isImport)
            {
                // Yes.
                // Export too?
                if (isExport)
                {
                    // Yes.
                    return ErrorMessage("Only one \"Import\" or \"Export\" parameter can be set.");
                }

                if (doEmail)
                {
                    return ErrorMessage("Can't import and email.");
                }

                // Have a file path?
                if (file == null)
                {
                    // No.
                    return ErrorMessage("\"File\" parameter/argument is required.");
                }

                // Have a table name?
                if (table == null)
                {
                    // No.
                    return ErrorMessage("\"Table\" parameter/argument is required.");
                }
            }
            else
            {
                // No, export.
                isExport = true;

                // Have a query?
                if (query == null)
                {
                    // No.
                    return ErrorMessage("\"Query\" parameter/argument is required.");
                }

                // Are we sending an email?
                if (doEmail)
                {
                    // Yes.
                    if (from == null)
                    {
                        return ErrorMessage("Must have a \"From\" email address");
                    }

                    if (to == null)
                    {
                        return ErrorMessage("Must have a \"To\" email address");
                    }

                    if (from == null)
                    {
                        return ErrorMessage("Must have a \"Subject\"");
                    }

                    if (smtpServer == null)
                    {
                        return ErrorMessage("Must have a \"SmtpServer\"");
                    }
                }
                else
                {
                    // No.
                    // Have a file path?
                    if (file == null)
                    {
                        // No.
                        return ErrorMessage("\"File\" parameter/argument is required.");
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// Start of main program.
        /// </summary>
        /// <param name="args">Program arguments.</param>
        private static void Main(string[] args)
        {
            try
            {
                string version = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                Console.WriteLine("SqlToCsv - SQL to Comma Separated Value (CSV) file."
                                    + "  Version: " + version);
                Console.WriteLine("             User: " + System.Environment.UserDomainName + "\\" + System.Environment.UserName);

                if (LoadArgumentParameters())
                {
                    if (isImport)
                    {
                        ImportCsvFileToSql();
                    }
                    else if (isExport)
                    {
                        if (doEmail)
                        {
                            ExportSqltoCsvEmail();
                        }
                        else
                        {
                            ExportSqltoCsvFile();
                        }
                    }
                }
            }
            catch (Exception caught)
            {
                if (caught is DirectoryNotFoundException
                    || caught is FileNotFoundException
                    || caught is OurException
                    || caught is SqlException)
                {
                    ErrorMessage(caught.Message);
                }
                else
                {
                    ErrorMessage(caught.ToString());
                }
            }
        }

        /// <summary>
        /// Set Mail Address List
        /// </summary>
        /// <param name="addressList">Mail address collection.</param>
        /// <param name="emailAddresses">Semicolon separated list of email address to add.</param>
        /// <param name="fieldName">Field name to display if there is an error.</param>
        /// <param name="required">True if email address is required.</param>
        private static void SetMailAddressList(MailAddressCollection addressList, string emailAddresses, string fieldName, bool required)
        {
            if ((emailAddresses == null) || (emailAddresses.Trim() == string.Empty))
            {
                if (required)
                {
                    throw new OurException(string.Format("\"{0}\" requires an email address", fieldName));
                }
                else
                {
                    return;
                }
            }

            string[] list = emailAddresses.Split(';');

            foreach (string item in list)
            {
                // Do we have an email address?
                if ((item != null) && (item.Trim() != string.Empty))
                {
                    // Yes.
                    MailAddress mailAddress = GetMailAddress(item, fieldName, false);

                    // Is it valid?
                    if (mailAddress != null)
                    {
                        // Yes.
                        addressList.Add(mailAddress);
                    }
                }
            }

            // Do we have nothing and its required to have something?
            if (addressList.Count <= 0 && required)
            {
                // Yes.
                throw new OurException(string.Format("\"{0}\" requires an email address", fieldName));
            }
        }

        /// <summary>
        /// Convert a yes/no string to Boolean
        /// </summary>
        /// <param name="value">A string that has a yes/true/1 as a true value.</param>
        /// <returns>Boolean value.</returns>
        private static bool YesNoToBoolean(string value)
        {
            switch (value.Trim().ToLower())
            {
                case "yes":
                case "true":
                case "1":
                    return true;
                default:
                    return false;
            }
        }
    }
}