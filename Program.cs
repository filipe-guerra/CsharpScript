using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;


namespace BuscaPastas
{
    class Program
    {
        static void Main(string[] args)
        {
            var directories = CustomSearcher.GetDirectories(@"C:\Users\Netlex\Documents\test\");
            /*foreach (var dir in directories)
            {
                Console.WriteLine(dir);
            }*/

            foreach (string path in directories)
            {
                if (File.Exists(path))
                {   // This path is a file
                    RecursiveFileProcessor.ProcessFile(path);
                }
                else if (Directory.Exists(path))
                {   // This path is a directory
                    RecursiveFileProcessor.ProcessDirectory(path);
                }
                else
                {
                    Console.WriteLine("{0} is not a valid file or directory.", path);
                }
            }

            Console.ReadKey();
        }
    }

    public class CustomSearcher
    {
        public static List<string> GetDirectories(string path, string searchPattern = "*",
            SearchOption searchOption = SearchOption.TopDirectoryOnly)
        {
            if (searchOption == SearchOption.TopDirectoryOnly)
                return Directory.GetDirectories(path, searchPattern).ToList();

            var directories = new List<string>(GetDirectories(path, searchPattern));

            for (var i = 0; i < directories.Count; i++)
                directories.AddRange(GetDirectories(directories[i], searchPattern));

            return directories;
        }

        private static List<string> GetDirectories(string path, string searchPattern)
        {
            try
            {
                return Directory.GetDirectories(path, searchPattern).ToList();
            }
            catch (UnauthorizedAccessException)
            {
                return new List<string>();
            }
        }
    }

    public class RecursiveFileProcessor
    {
        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public static void ProcessDirectory(string targetDirectory)
        {
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
                ProcessFile(fileName);

            // Recurse into subdirectories of this directory.
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            foreach (string subdirectory in subdirectoryEntries)
                ProcessDirectory(subdirectory);
        }

        // Insert logic for processing found files here.
        public static void ProcessFile(string path)
        {
            Console.WriteLine("Processed file '{0}'.", path);
        }
    }
}
