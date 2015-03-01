This add-in allows to verify that repositories in the file system comply with an XML definition.
 
 An Excel file contains the list of repositories we want to manage in a worksheet.
 Place this file in a directory, along with the XML definition of the repository structure.
 This directory is also the root of all repositories
 
 Open the Excel file, go to the top of your list and click "Check Repositories".
 Non-compliance errors with the model - if any - will be added as a comment to the cell that contains the name of the repo
 
 XML guideline:
 - Root element must be named "repoRoot"
 - A "directory" node must have an attribute "name" that identifies an expected subdirectory
 - A "directory" node can have an attribute "allowotherfiles". If it is set to "no", then the content of the directory must
   match exactly what the XML describes, i.e: you cannot put extra stuff is that directory
 - A "file" node must have an attribute "name" that identifies an expected file
 - The "name" attributes of both files and dirs can contain wildcards
 
 Example:

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

 if you have a cell that contains "C# course" in your workbook, you select it and click "check repositories",
 the repository is compliant if:
 - There is a directory named "C# course" in the same folder where the excel file is
 - This directory contains two subdirs ("Module description" and "Exercises"), nothing more, nothing less
 - The "Module description" subdir contains only:
     - one file named "Readme.txt"
     - one or many files named "Description". For example: "Description.docx" and "Description.pdf"
 - The "Exercises" subdir contains at least two subdirs ("Solutions" and "Data"). But there can be other files 
   and directories in there too
