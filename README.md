# HCML
HCML is a manuscript preparation template created by John Atkins for HarperCollins Publishers.

NOTE: This template is proprietary and is made available here only as a work sample.

# About the Project
## Purpose
HCML was created as a suite of Microsoft Word-based tools to assist editors at the publishing house to convert author-submitted manuscripts to a consistent format that can be converted automatically to a standard XML using XSLT.

##Technology
* Microsoft Word 2007 or higher for Windows
* Microsoft Word 2016 or higher for Mac
* Uses standard VBA for automation within Word
* Uses XML to modify the standard Word Ribbon

#Usage
The suite consists of two Word templates:
* HCML.dotm: Contains a set of custom Word styles to be applied to paragraphs and characters. These provide semantic structure to the content, and are used downstream in the process to convert content to XML, and to produce both print and electronic formats from the XML. Also incuded in this tool are various routines such as interactive forms and input boxes to help the use apply consistent and proper structure to the manuscript. 
* HCCleanup.dotm: This is an extended toolset to provide some automated routines for editors to help clean up common typos, check the manuscript for structural errors, and to generate standard reports based on the content.
* InstallHCML.docm: This is an installation routine to automatically install the templates in their proper locations for use.

#My Contributions
I took over this project very early in its development. Following is a list of functionality that I was responsible for:
* Designed the content architecture, based on diverse content requirements (fiction, business, culinary, high-design)
* Integrated some legacy "cleanup" routines, converting to the new content model, and refactored code to improve accuracy
* Created new routines to 
  * validate document structure
  * automatically fix errors that may have been introduced by the Word software
  * auto-generate content such as copyright pages based on various book imprints
  * generate tables of contents 
  * generate lists of styles used in documents
  * create favorite lists of styles to use
  * identify and style special characters in the document
  * easily create custom keyboard shortcuts for accessing the functionality
  * identify and convert fractions to a standard format
  * check formatting of endnotes and footnotes and fix any problems automatically
  * automation to link endnotes and footnotes using Word's Notes feature
  * generate art lists based on standard styling of art througout the manuscript
