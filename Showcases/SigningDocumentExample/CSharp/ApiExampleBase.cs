// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using NUnit.Framework;

namespace SigningDocumentExample
{
    public class ApiExampleBase
    {
        private readonly string dirPath = MyDir + @"\Artifacts\";

        [SetUp]
        public void SetUp()
        {
            if (!Directory.Exists(dirPath))
                //Create new empty directory
                Directory.CreateDirectory(dirPath);
        }

        [TearDown]
        public void TearDown()
        {
            //Delete all dirs and files from directory
            Directory.Delete(dirPath, true);
        }

        /// <summary>
        /// Returns the assembly directory correctly even if the assembly is shadow-copied.
        /// </summary>
        private static string GetAssemblyDir(Assembly assembly)
        {
            // CodeBase is a full URI, such as file:///x:\blahblah.
            Uri uri = new Uri(assembly.CodeBase);
            return Path.GetDirectoryName(uri.LocalPath) + Path.DirectorySeparatorChar;
        }

        /// <summary>
        /// Gets the path to the currently running executable.
        /// </summary>
        internal static string AssemblyDir
        {
            get { return gAssemblyDir; }
        }

        /// <summary>
        /// Gets the path to the documents used by the code examples. Ends with a back slash.
        /// </summary>
        internal static string MyDir
        {
            get { return gMyDir; }
        }

        static ApiExampleBase()
        {
            gAssemblyDir = GetAssemblyDir(Assembly.GetExecutingAssembly());
            gMyDir = new Uri(new Uri(gAssemblyDir), @"../../../Data/").LocalPath;
        }

        private static readonly string gAssemblyDir;
        private static readonly string gMyDir;
    }
}