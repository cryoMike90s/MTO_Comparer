# Two xlsx files logic comparision script
> Script was made due shortages of time in present job to spend so much time to compare output file from deployed
> modeling software (Autocad Plant 3d). Task of this script is to compare two version of one file,
> (second one is newer) with utilized index set on chosen columns and give output with marked information which rows
> were removed, added or changed.


## Table of Contents
* [General Info](#general-information)
* [Technologies Used](#technologies-used)
* [Features](#features)
* [Setup](#setup)
* [Usage](#usage)
* [Project Status](#project-status)
* [Room for Improvement](#room-for-improvement)
* [Acknowledgements](#acknowledgements)

<!-- * [License](#license) -->


## General Information
- By usage libraries such as pandas, numpy and os there is 
  possible to check if/where the fresh version differs from previous one.
- Main reason for creation this script was to cut significantly time which now is spending
  by designers in my present workplace to compare such files and allow them to focus on tasks which needs engineering
  or just drink another coffee in a canteen.


## Technologies Used
- Python - version 3.9.2

## Libraries
- numpy - version 1.20.3
- pandas - version 1.2.4
- openpyxl - version 3.0.7


## Setup
Only one requirement is to have adequate Excel files which normally comes from 3d modeling software and install 
listed libraries.


## Project Status
Project is: _in progress_ 


## Room for Improvement
To do:
- Split the actual final file DataFrame for smaller with selection "Sheet name" as split index
- Merge changed data with the rest (rows which did not change in newer version and which were rejected
  during processing)


## Acknowledgements
- This project was inspired by website: https://pbpython.com

<!-- Optional -->
<!-- ## License -->
<!-- This project is open source and available under the [... License](). -->

<!-- You don't have to include all sections - just the one's relevant to your project -->