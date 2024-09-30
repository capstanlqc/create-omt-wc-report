# Create OmegaT wordcount report

## What it does

This script obtains the wordcount statistics from a number of omegat projects and creates a report with the word counts for each target language (assuming the source language is always English). 

This can be helpful when the source files (or the text extracted) are different in each omegat project.

## Business logic 

- All omegat project folders inside the input directory that match the `project_filter` are collected for processing. 
- From each of them, the stats file in JSON format is read and the wordcount statistics is extracted from it. 
- If there are subsets defined as the `file_filter` configuration, the wordcount of all files included in that subset is calculated and added to the report in a separate row.
- Results are saved, either as one report spreadsheet file including one sheet per language or as one spreadsheet per language.

## Input

The path to the folder containing all omegat project folders, one project per language. 

## Output

The output will include one worksheet (as one sheet in the same workbook or as a separate file) per language. 

For each language, the worksheet contains one columne for each these:

- the `total` number of words (including repeated segments) 
- the total number of `unique` words (not including repeated segments)
- the number of `remaining` words (including repetitions)
- the number of `unique remaining` words (not including repetitions)

The report includes one row for each for these: 

- project: the whole number of words in the project 
- each file in the project (if configuration option `get_stats_per_file` is enabled)
- for subsets (if any values in the configuration optoin `file_filter` are found)

## Configuration

You can modify the behaviour of the script with the following configuration options: 

- `project_filter`: list of substrings that the path to each omegat project must contain to be included in the report
- `get_stats_per_file`: boolean value that indicates whether an individual row for each file in the project must be included in the report (if `get_stats_per_file` is false, any files outside of the defined subsets will not be included in the report)
- `file_filter`: list of substrings that will be used to find subsets of files in each project, each of which will have its own row in the report

For example, the following configuration:

```json
{
    "project_filter": ["ar-JO", "ar-PS", "uz-KG", "ru-YY"],
    "get_stats_per_file": false,
    "file_filter": ["assessments", "questionnaires"],
    "console_translate": false,
    "one_sheet_per_language": true
}
```

will produce a report like this for each project found for locales "ar-JO", "ar-PS", "uz-KG" or "ru-YY":

|                | total     | remaining | unique | unique-remaining |
|----------------|-----------|--------|------------------|----|
| **project**        | 7495      | 52     | 3024             | 52 |
| *questionnaires* | 7134      | 52      | 2669             | 52  |
| *assessments*    | 361       | 0      | 355              | 0  |


## How to use it

Run `python app.py --help` to see how to run the script and which arguments it requires, e.g. 

```bash
python app.py -p /path/to/parent/directory -r /path/to/wordcounts.xlsx
```

## Disclaimers

The intended purpose is to analyze a directory containing at its first level a number of omegat project folders for different target languages. 

``` 
tree /path/to/parent/directory

├── pisa_2025stg_translation_ar-JO_verification-review
├── pisa_2025stg_translation_ar-PS_verification-review
├── pisa_2025stg_translation_uz-KG_verification-review
└── pisa_2025stg_translation_ru-YY_verification-review
``` 

The report contains one sheet per target language, so if several projects for the same target language are processed, the latest will always overwrite the previous one. To avoi this, an index sheet could be addd that lists all projects and points to the sheet for each project (as in Excel exports from OmegaT), so that sheets can be named with numbers or something else (not necessarily the omegat project name).

## TODO

Potential extensions of this script: 

- Input folder contains a list of OMT packages, and the script unpacks them and runs OmegaT on them to generate/update the stats file.
- Input folder (at the root level) is an omegat project folder, in which case the script analyses that project only instead of looking for omegat project children folders at the first level. 
- In case any projects do not already included a `project_stats.json` file (or in case it's not updated), the script runs omegat on each project (if option `console_translate` is enabled) to genera the stats file.