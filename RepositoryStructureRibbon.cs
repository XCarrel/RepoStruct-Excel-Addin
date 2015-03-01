using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;

/*
 * Repository Structure Excel Add-In
 * Version 1.0
 * Author: X. Carrel
 * February 1, 2015
 * 
 * This add-in allows to verify that repositories in the file system comply with an XML definition.
 * 
 * An Excel file contains the list of repositories we want to manage in a worksheet.
 * Place this file in a directory, along with the XML definition of the repository structure.
 * This directory is also the root of all repositories
 * 
 * Open the Excel file, go to the top of your list and click "Check Repositories".
 * Non-compliance errors with the model - if any - will be added as a comment to the cell that contains the name of the repo
 * 
 * XML guideline:
 * - Root element must be named "repoRoot"
 * - A "directory" node must have an attribute "name" that identifies an expected subdirectory
 * - A "directory" node can have an attribute "allowotherfiles". If it is set to "no", then the content of the directory must
 *   match exactly what the XML describes, i.e: you cannot put extra stuff is that directory
 * - A "file" node must have an attribute "name" that identifies an expected file
 * - The "name" attributes of both files and dirs can contain wildcards
 * 
 * Example:

<repoRoot>
	<directory name="Module description" allowotherfiles="no">
		<file name="Description.*" />
		<file name="Readme.txt" />
	</directory>
	<directory name="Exercises">
		<directory name="Solutions"></directory>
		<directory name="Data"></directory>
	</directory>
</repoRoot>

 * if you have a cell that contains "C# course" in your workbook, you select it and click "check repositories",
 * the repository is compliant if:
 * - There is a directory named "C# course" in the same folder where the excel file is
 * - This directory contains two subdirs ("Module description" and "Exercises"), nothing more, nothing less
 * - The "Module description" subdir contains only:
 *     - one file named "Readme.txt"
 *     - one or many files named "Description". For example: "Description.docx" and "Description.pdf"
 * - The "Exercises" subdir contains at least two subdirs ("Solutions" and "Data"). But there can be other files 
 *   and directories in there too
 */
namespace RepoStruct
{
    public partial class RepositoryStructureRibbon
    {
        string validationReport; // result string: contains all the error messages
        const string XMLDefinitionOfRepositoryStructure = "Structure d'un repository de module.xml";

        private bool directoryIsOK(XmlNode n, string path, bool allowOtherFiles, bool isRoot)
        {
            bool result = true; // optimistic start: it'll all be fine!!

            if (isRoot) validationReport = ""; // not a recursive call: reset error messages

            if (!Directory.Exists(path))
            {
                validationReport += (Environment.NewLine + path + ": ce dossier n'existe pas");
                return false;
            }

            // These lists are used to check if there are too many files or directories in the directory
            List<string> allFiles = Directory.GetFiles(path).ToList();
            List<string> allDirs = Directory.GetDirectories(path).ToList();

            foreach (XmlNode cn in n.ChildNodes) // let's do this...
                switch(cn.Name)
                {
                    case "directory":
                        string[] dirs = Directory.GetDirectories(path, cn.Attributes[0].Value); // list the directories of the current path that match the definition. There could be more than one, wildcards being allowed
                        if (dirs.Length > 0) // subdir found
                        {
                            bool aof = !((cn.Attributes.Count > 1) && (cn.Attributes[1].Name == "allowotherfiles") && (cn.Attributes[1].Value == "no")); // evaluates to true if we allow the subdir to have extra files
                            foreach (string sdir in dirs) // scan all found subdirs recursively
                            {
                                result &= directoryIsOK(cn, sdir, aof, false); // aggregate subdir result with current
                                if (!allowOtherFiles) // Remove it from the list
                                    allDirs.RemoveAt(allDirs.IndexOf(sdir));
                            }
                        }
                        else // subdir not found
                        {
                            validationReport += (Environment.NewLine + "Dossier manquant: " + cn.Attributes[0].Value);
                            result = false;
                        }

                        break;
                    case "file":
                        string[] files = Directory.GetFiles(path, cn.Attributes[0].Value);// list the files of the current path that match the definition. There could be more than one, wildcards being allowed
                        if (files.Length > 0) // file found
                            if (!allowOtherFiles) // Remove it from the list
                                allFiles.RemoveAt(allFiles.IndexOf(files[0]));
                            else ; // nothing
                        else // file not found
                        {
                            validationReport += (Environment.NewLine + "Fichier manquant: " + cn.Attributes[0].Value);
                            result = false;
                        }
                        break;
                }

            if (!allowOtherFiles) // check the leftovers in the lists. They should both be empty
            {
                if (allFiles.Count > 0)
                {
                    foreach (string f in allFiles) validationReport += (Environment.NewLine + "Fichiers excédentaire: " + f);
                    result = false;
                }
                if (allDirs.Count > 0)
                {
                    foreach (string f in allDirs) validationReport += (Environment.NewLine + "Dossier excédentaire:" + f);
                    result = false;
                }
            }
            return result;
        }

        private void cmdCheckRepo_Click(object sender, RibbonControlEventArgs e)
        {
            XmlDocument xDoc = new XmlDocument();
            string rootPath = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
            try
            {
                xDoc.Load(rootPath + @"\" + XMLDefinitionOfRepositoryStructure); // get the XML definition of the repo.
                if (xDoc.FirstChild.Name != "repoRoot") throw new Exception("Noeud root incorrect");
                Excel.Worksheet F = Globals.ThisAddIn.Application.ActiveSheet;

                int line = Globals.ThisAddIn.Application.Selection.Row;
                int column = Globals.ThisAddIn.Application.Selection.Column;

                bool everythingOK = true; // let's start optimistic

                while (F.Cells[line, column].Value != null) // step down the column until we hit an empty cell
                {
                    if (F.Cells[line, column].Comment != null) F.Cells[line, column].Comment.Delete(); // remove cell comment if any
                    if (!directoryIsOK(xDoc.FirstChild, rootPath + @"\" + F.Cells[line, column].Value, false, true))
                    {
                        validationReport = DateTime.Now.ToString("d") + ": ### Repository pas OK ###" + validationReport;
                        everythingOK = false;
                        F.Cells[line, column].AddComment(validationReport); // add messages as cell comment
                    }
                    line++;
                }
                if (everythingOK)
                    MessageBox.Show("Tous les repository sont OK !!");
                else
                    MessageBox.Show("Il y a des erreurs, voir commentaires");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Le fichier de description du repository est pourrave!!! (" + ex.Message+")");
            }
        }
    }
}
