RInsightUIConsoleVb
===================

## Overview
This repository provides a framework for developing user interface specifications for dialogs.
It achieves this by traversing the tree of transformations specified in the dialog's JSON file. 
It then identifies which values need to be set at the top-level (i.e. in the main panel of the dialog) and which values only need to be updated when specific conditions are met (e.g. visibility triggered by a checkbox being checked).

The framework creates a tree of UI elements. 
It outputs this tree to two output files in the `tmp` folder on the desktop: a JSON file, and a more human-readable `txt` file.
## Usage
If you wish to use this framework, then for each dialog you want to process, you need to create a subfolder in the `DialogDefinitions` folder. 
Within this subfolder, you need to put a JSON file containing the list of transformations for that dialog. 
These transformations define how the dialog settings update the R script associated with the dialog. 
For more details about these transformations, please see the [RInsight](https://github.com/lloyddewit/RInsightF461) library.
You also need to add a metadata JSON file which contains extra information about each UI element in the dialog. 
See the examples in this repository for more details.