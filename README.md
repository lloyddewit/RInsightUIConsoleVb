RInsightUIConsoleVb
==============

## Overview
This repository is a prototype to create a user interface specification for a dialog.
It does this by traversing the tree of transformations specified in the dialog's JSON file. It then determines which values need to be set at the top-level (i.e. in the main panel of the dialog) and which values only need to be set when a condition is met (e.g. only need to be made visible when a checkbox is checked).