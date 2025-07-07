RInsightUIConsoleVb
===================

## Overview
This repository provides a framework for developing user interface specifications for dialogs.
It achieves this by traversing the tree of transformations specified in the dialog's JSON file. 
It then identifies which values need to be set at the top-level (i.e. in the main panel of the dialog) and which values only need to be updated when specific conditions are met (e.g. visibility triggered by a checkbox being checked).