# ReplayAnalysis

This is a repository for a python tool created by me to read from the json files outputted from redraskal's R6-Dissect replay file parser (https://github.com/redraskal/r6-dissect).
Meant to provide the same level of in-depth statistics for teams without the extra headache or time involved with manually extracting the information from VODs.

It takes in the json, and will parse through it to output an excel file with per round and overall match statistics, modelled after those used in Pro League by Siege.GG. 

To use, simply clone the repo and run ReplayAnalyzer.py and it will prompt you for the path to the json file you would like to analyze, and then it will output the xlsx file into a folder called "Output" in the same directory.
Dependencies: json, xlsxwriter, os

End of Match stats page
![image](https://github.com/Zander-9909/ReplayAnalysis/assets/71144499/88589f3e-d34d-42ee-864c-929926b40741)

Round by round stat page
![image](https://github.com/Zander-9909/ReplayAnalysis/assets/71144499/22eb7aea-f534-4bcb-bb99-938b85b11c63)

Hope this is helpful!
Griffin Taylor
Gamehead for R6
UOttawa Esports