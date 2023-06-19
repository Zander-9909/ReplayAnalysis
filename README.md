# ReplayAnalysis

This is a repository for a python tool created by me to read from the json files outputted from redraskal's R6-Dissect replay file parser (https://github.com/redraskal/r6-dissect).
Meant to provide the same level of in-depth statistics for teams without the extra headache or time involved with manually extracting the information from VODs. The KOST formula is the one that [Reaper_en reverse engineered](https://www.youtube.com/watch?v=faoQZK2875Q) and that the Overwolf app R6-Analyst uses. As well, it uses the (as far as I know) accepted trade time of 10s.

It takes in the json, and will parse through it to output an excel file with per round and overall match statistics, modelled after those used in Pro League by Siege.GG. 

To use, simply clone the repo and run ReplayAnalyzer.py and it will prompt you for the path to the json file you would like to analyze, and then it will output the xlsx file into a folder called "Output" in the same directory.
Dependencies: json, xlsxwriter, os

End of Match stats page
![image](https://github.com/Zander-9909/ReplayAnalysis/blob/main/Screenshots/Overall%20Match%20View.png)

Round by round stat page
![image](https://github.com/Zander-9909/ReplayAnalysis/blob/main/Screenshots/Round%20stat%20view.png)

Hope this is helpful!

Griffin Taylor

Gamehead for R6

UOttawa Esports
