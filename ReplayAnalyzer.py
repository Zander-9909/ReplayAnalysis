import json
import xlsxwriter
import os

def split_dictionary(input_dict, chunk_size):
    res = []
    new_dict = {}
    for k, v in input_dict.items():
        if len(new_dict) < chunk_size:
            new_dict[k] = v
        else:
            res.append(new_dict)
            new_dict = {k: v}
    res.append(new_dict)
    return res

def writeHeader(sheet,chkRound):
    header = ["Team Name","Team ID","Username", "Rating","Kills", "Deaths","K/D", "Headshot %","KOST","Entry Diff.", "KPR", "SRV","Trade Diff.","Clutch", "Plant", "Defuse"]
    if(chkRound):
        header = ["Username","Team ID", "Operator","Kills", "Died", "Plant", "Defuse","Death Traded","KOST Point","# of Headshots","OK","OD", "Clutch"]
    row,column = 0,0
    for entry in header:
        sheet.write(row,column, entry)
        column += 1

def writeRoundStats(sheet, roundDict,teamNames):
    row = 1
    column = 1
    writeHeader(sheet,True)

    #[teamID, operator, kills, died, planted, defused, traded, KOST Point, # of headshots,OK, OD,clutch]
    roundStats = roundDict.get("stats")
    for player in roundStats:
        i = 0
        column = 1
        sheet.write(row,0, player)
        for data in roundStats.get(player):
            if(i >=3 and i<=7):
                if(type(data)==int and data >0):
                    data = True
                elif(type(data)==int and data <= 0):
                    data = False
            sheet.write(row,column, data)
            i+=1
            column +=1
        row+=1
    sheet.write(12, 0,"Site: ")
    sheet.write(12,1,roundDict.get("site"))
    sheet.write(13,0,"Winner: ")
    sheet.write(13,1,roundDict.get("winner"))
    sheet.write(14,0,"Win Type: ")
    sheet.write(14,1,roundDict.get("winCondition"))

def writeEndGame(sheet,gameInfo, overallStats,siteStatistics):
    totalRounds = gameInfo.get("Team0Score") + gameInfo.get("Team1Score")
    row = 1
#["Team Name","Username", "Kills", "Deaths","K/D", "Headshot %","KOST","Entry Diff.", "KPR", "SRV",Trade Diff.,"Clutch", "Plant", "Defuse"]
#overallStats: Team Name, Kills, Deaths, KOST Points, OK, OD, Clutch, Plant, Defuse,Deaths Traded,Headshots
    for player in overallStats:
        
        statArray = overallStats.get(player)
        teamID = 0
        if(statArray[0] == gameInfo.get("team0Name")):teamID = 0
        else: teamID = 1
        kills = statArray[1]
        deaths = statArray[2]
        numHeadshots = statArray[10]
        if(deaths == 0): kd = round(kills,2)
        else: kd = round(kills/deaths,2)
        OK = statArray[4]
        OD = statArray[5]
        if(kills != 0):
            headshots = round(numHeadshots/kills,2)
        else: headshots = kills
        kost = statArray[3]/totalRounds
        kostPerc = round(kost*100,2)
        entryDiff = OK - OD
        kpr = round(kills/totalRounds,2)
        srv = round(1-deaths/totalRounds,2)
        tradeDiff = statArray[9] - statArray[2]
        clutches = statArray[6]
        plants = statArray[7]
        defuses = statArray[8]
        rating = round(0.037 + 0.0004*OK - 0.005*OD + 0.714*kpr + 0.492*srv + 0.471*kost + 0.026*clutches + 0.015*plants + 0.019*defuses,2)

        sheet.write(row,0, statArray[0])#Team Name
        sheet.write(row,1, teamID)# TeamID
        sheet.write(row,2, player)# Username
        sheet.write(row,3,rating)# Rating
        sheet.write(row,4, statArray[1])#Kills
        sheet.write(row,5, statArray[2])#Deaths
        sheet.write(row,6, kd) #K/D
        sheet.write(row,7, headshots) #Headshot Percentage
        sheet.write(row,8, kostPerc)#KOST %
        sheet.write(row,9, entryDiff)#entry +/-
        sheet.write(row,10, kpr)#KPR
        sheet.write(row,11, srv)#Survival rate
        sheet.write(row,12, tradeDiff)#Trade Diff (Traded Deaths - Deaths)
        sheet.write(row,13, clutches)#Clutches
        sheet.write(row,14, plants)#Plants
        sheet.write(row,15, defuses)#Defuses
        row+=1
    sheet.write(12, 0,"Map: ")
    sheet.write(12,1,gameInfo.get("Map"))
    sheet.write(13,0,"Team 0 Score: ")
    sheet.write(13,1,gameInfo.get("Team0Score"))
    sheet.write(14,0,"Team 1 Score: ")
    sheet.write(14,1,gameInfo.get("Team1Score"))
    #Write site win/loss statistics
    sheet.write(12,3,"Site Name")
    sheet.write(12,4,"T0 Att Plays")
    sheet.write(12,5,"T0 Att W/L")
    sheet.write(12,6,"T0 Def Plays")
    sheet.write(12,7,"T0 Def W/L")
    sheet.write(12,8,"T1 Att Plays")
    sheet.write(12,9,"T1 Att W/L")
    sheet.write(12,10,"T1 Def Plays")
    sheet.write(12,11,"T1 Def W/L")
    siteRow = 13
    for site in siteStatistics:
        sheet.write(siteRow,3,site)
        siteDic = siteStatistics.get(site)
        wins = siteDic.get("wins")
        plays = siteDic.get("plays")
        siteCol = 4
        for i in range (len(plays)):
            sheet.write(siteRow,siteCol,plays[i])
            siteCol +=1
            if(plays[i]!= 0):
                sheet.write(siteRow,siteCol,"{:.0%}".format(wins[i]/plays[i]))
            else: 
                sheet.write(siteRow,siteCol,"{:.0%}".format(wins[i]/1))
            siteCol +=1
        siteRow+=1



def writeToExcel(gameInfo,perRoundStats,fullGameStats,siteStatistics):#Pass full game info as a list. Each entry is a dictionary
    bookName = "Output/"+gameInfo.get("Map")+"_"+ str(gameInfo.get("Team0Score"))+"-"+ str(gameInfo.get("Team1Score")) + "_"+gameInfo.get("Date")+".xlsx"
    workbook = xlsxwriter.Workbook(bookName)
    teamNamesList = []
    teamNamesList.append(gameInfo.get("team0Name"))
    teamNamesList.append(gameInfo.get("team1Name"))
    fullGameSheet = workbook.add_worksheet("Full Game Stats")#Final Game Scoreboard
    writeHeader(fullGameSheet,False)
    #Write Full Game Stats
    writeEndGame(fullGameSheet,gameInfo, fullGameStats,siteStatistics)

    inc = 1
    for roundE in perRoundStats:
        roundSheet = workbook.add_worksheet("Round "+str(inc))
        writeHeader(roundSheet,True)
        writeRoundStats(roundSheet,roundE,"")
        inc += 1
    workbook.close()

def addToOverallStats(currOverall, roundStats):
#["Team Name","Username", "Kills", "Deaths","K/D", "Headshot %","KOST","Entry Diff.", "KPR", "SRV",Trade Diff.,"Clutch", "Plant", "Defuse"]
#overallStats: Team Name, Kills, Deaths, KOST Points, OK, OD, Clutch, Plant, Defuse,Deaths Traded,Headshots
#roundStats:[teamID, operator, kills, died, planted, defused, traded, KOST Point, # of headshots,OK, OD,1vX]
    for player in roundStats:
        currStats = currOverall.get(player)
        toAddStats = roundStats.get(player)
        currStats[1] += toAddStats[2]#add Kills
        if(toAddStats[3]) : currStats[2]+=1#add Death
        if(toAddStats[4]) : currStats[7]+=1#add Plant
        if(toAddStats[5]) : currStats[8]+=1#add Defuse
        if(toAddStats[6]) : currStats[9]+=1#add Traded Death
        if(toAddStats[7]) : currStats[3]+=1#add KOST point
        currStats[4]+= toAddStats[9]#add OK
        currStats[5]+= toAddStats[10]#add OD
        currStats[6]+= toAddStats[11]#add Clutch
        currStats[10]+= toAddStats[8]#add Headshots
        currOverall.update({player:currStats})
    return currOverall

        

def getRoundStats(roundDict): #get player stats for each round
    playerList = roundDict.get("players")
    teamInfo = roundDict.get("teams")
    outDict = {}#Output Dictionary
    for player in playerList:
        operator = player.get("operator").get("name")
        playerStat = [0]*12
        playerStat[0] = player.get("teamIndex")
        playerStat[1] = operator
        playerStat[6] =False
        outDict.update({player.get("username"):playerStat})

    matchFeedback = roundDict.get("matchFeedback") #list of dictionaries of stuff that happens in the round
    roundKills = []
    killTimes = []
    roundPlant = []
    roundDefuse = []

    for event in matchFeedback:#Get info about each important event
        eventType = event.get("type").get("name")
        if(eventType =="Kill"):#Add all kills to "roundKills"
            killer = event.get("username")
            killed = event.get("target")
            if(outDict.get(killer)[0] != outDict.get(killed)[0]):
                killInfo = [killer,killed,event.get("headshot"),False]
                killTimes.append([killer,killed,event.get("timeInSeconds")])
                roundKills.append(killInfo)#append list of info about each kill (killer, killed, headshot,TK (true or false))
            else:
                killInfo = [killer,killed,event.get("headshot"),True]
                roundKills.append(killInfo)#append list of info about each kill (killer, killed, headshot)
        elif(eventType == "DefuserPlantComplete"):
            plantInfo = [event.get("username")]
            roundPlant = plantInfo
        elif(eventType == "DefuserDisableComplete"):
            defuseInfo = [event.get("username")]
            roundDefuse = defuseInfo

    #Format for array of each player's stats
    #[teamID, operator, kills, died, planted, defused, traded, KOST Point, # of headshots,OK, OD,1vX]
    first = True
    for killEvent in roundKills:#adding kill info to each player's stat array
        killer = killEvent[0]
        killed = killEvent[1]
        killerStats = outDict.get(killer)
        killedStats = outDict.get(killed)
        if(first):#adding OK or OD
            if(not killEvent[3]):
                killerStats[9] +=1
            killedStats[10] +=1
            first = False

        if(not killEvent[3]):
            killerStats[2] +=1
            killerStats[7] += 1 #Adding KOST Point
            if(killEvent[2]):#headshots
                killerStats[8]+=1
        killedStats[3] +=1
        outDict.update({killer : killerStats})
        outDict.update({killed : killedStats})
    #adding Plant/Defuse
    if(len(roundPlant)> 0):
        planterStats = outDict.get(roundPlant[0])
        planterStats[4] += 1
        planterStats[7] += 1 
        outDict.update({roundPlant[0]:planterStats})
    
    if(len(roundDefuse)> 0):
        defuserStats = outDict.get(roundDefuse[0])
        defuserStats[5] += 1
        defuserStats[7] += 1 
        outDict.update({roundDefuse[0]:defuserStats})
    
    #Calculating Trades (Traded within 10s)
    index = 0
    for killEvent in killTimes:
        killer = killEvent[0]
        killed = killEvent[1]
        killerTeam = outDict.get(killer)[0]
        killTime = killEvent[2]

        for j in range(0,index): #Checking every kill BEFORE this one, to see if this one was a trade
            possTraded = killTimes[j][1]
            possTradeTime = killTimes[j][2]
            tradedTeam = outDict.get(killTimes[j][1])[0]

            if(tradedTeam == killerTeam):#if a previous killed player had the same team as the current killer
                if(possTradeTime - killTime <= 10):#if possTradeTime - killTime <= 10
                    #Successful Trade
                    tradedStats = outDict.get(possTraded)
                    tradedStats[6] = True
                    tradedStats[7] += 1
                    outDict.update({possTraded : tradedStats})
        index += 1
    #Calculating 1vX Stat (if the team that won only had 1 person alive at the end)
    playersAlive = []
    winnerTeamID = 0
    if(teamInfo[1].get("won")):
        winnerTeamID = 1
    for player in outDict:#counting players that did not die
        if(outDict.get(player)[0]==winnerTeamID and outDict.get(player)[3]==0): #If a player was alive at the end and they won
            playersAlive.append(player)
    if(len(playersAlive)==1):#If there was only one person alive at the end
        stats = outDict.get(str(playersAlive[0]))
        stats[11] += 1
        outDict.update({playersAlive[0]:stats})
    return outDict


# Opening JSON file
filed = input("Please input the path to the json file.\n")
f = open(filed)

#making output Directory
directory = "Output/"

if not os.path.exists(directory):
    os.makedirs(directory)

# returns JSON object as 
# a dictionary
data = json.load(f)
game_data = split_dictionary(data, 1)[0]
score_board = split_dictionary(data, 1)[1]
rounds = game_data.get("rounds")#list of rounds. Each is a dictionary
players = rounds[0].get("players")#list of players
#5 row array. Row index 0-1 are # win on Att or Def for Team 0, 2-3 are # win on Att or Def for Team 1
siteStatistics = {}

#[teamID, operator, kills, died, planted, defused, traded, KOST Point, # of headshots,OK, OD,clutch]
rawRoundStats = []
Team0Score = 0
Team1Score = 0
teamInfo1 = rounds[0].get("teams")
overallStats = {}
for player in players:
    newArray = [0]*11
    playerName = player.get("username")
    playerTeamID = player.get("teamIndex")
    newArray[0] = teamInfo1[playerTeamID].get("name")
    overallStats.update({playerName:newArray})
for roundA in rounds:
    roundInfo = {}
    roundInfo.update({"site": roundA.get("site")})#Site
    
    teamInfo = roundA.get("teams")
    roundInfo.update({"team0Name": teamInfo[0].get("name")})#Team 0 Name
    roundInfo.update({"team1Name": teamInfo[1].get("name")})#Team 1 Name
    roundInfo.update({"team0Side": teamInfo[0].get("role")})#Team 0 side
    roundInfo.update({"team1Side": teamInfo[1].get("role")})#Team 1 side
    if "winCondition" in teamInfo[0]:#Getting Round winner and Win Condition
        roundInfo.update({"winner": teamInfo[0].get("name")})#Team 0 Name
        roundInfo.update({"winCondition": teamInfo[0].get("winCondition")})#Team 0 Name
        Team0Score += 1 
    else:
        roundInfo.update({"winner": teamInfo[1].get("name")})#Team 0 Name
        roundInfo.update({"winCondition": teamInfo[1].get("winCondition")})#Team 0 Name
        Team1Score += 1 
    #Add round statistics to siteStatistics array
    
    if roundInfo.get("site") in siteStatistics:
        currSiteDic = siteStatistics.get(roundInfo.get("site"))
        currSiteArray = currSiteDic.get("wins")
        currSitePlays = currSiteDic.get("plays")
        #Site played previously
        if(roundInfo.get("winner")==roundInfo.get("team0Name")):
            #Team 0 Won
            if(roundInfo.get("team0Side") == "Attack"):
                #Att win on that site
                currSiteArray[0] +=1 #add attack win
                currSitePlays[0] +=1
                currSitePlays[3] +=1
            else: 
                #Def win on that site
                currSiteArray[1] +=1 #add Def win
                currSitePlays[1] +=1
                currSitePlays[2] +=1
        else:
            #Team 1 Won
            if(roundInfo.get("team1Side") == "Attack"):
                #Att win on that site
                currSiteArray[2] +=1 #add attack win
                currSitePlays[2] +=1
                currSitePlays[1] +=1
            else: 
                #Def win on that site
                currSiteArray[3] +=1 #add Def win
                currSitePlays[3] +=1
                currSitePlays[0] +=1
        currSiteDic.update({"plays":currSitePlays})  
        currSiteDic.update({"wins":currSiteArray})  
        siteStatistics.update({roundInfo.get("site"):currSiteDic})
    else:
        arrPlays = [0,0,0,0]
        arrWins = [0,0,0,0]
        if(roundInfo.get("winner")==roundInfo.get("team0Name")):
            #Team 0 Won
            if(roundInfo.get("team0Side") == "Attack"):
                #Att win on that site
                arrWins[0] +=1 #add attack win
                arrPlays[0] +=1 #add attack win
                arrPlays[3] +=1 
            else: 
                #Def win on that site
                arrWins[1] +=1 #add attack win
                arrPlays[1] +=1 #add attack win
                arrPlays[2] +=1 
        else:
            #Team 1 Won
            if(roundInfo.get("team1Side") == "Attack"):
                #Att win on that site
                arrWins[2] +=1 #add attack win
                arrPlays[2] +=1 #add attack win
                arrPlays[1] +=1 
            else: 
                #Def win on that site
                arrWins[3] +=1 #add attack win
                arrPlays[3] +=1 #add attack win
                arrPlays[0] +=1 
        newDic = {}
        newDic.update({"wins":arrWins})
        newDic.update({"plays":arrPlays})
        siteStatistics.update({roundA.get("site"):newDic})  
    roundStats = getRoundStats(roundA)
    roundInfo.update({"stats":roundStats})
    overallStats = addToOverallStats(overallStats, roundStats)
    rawRoundStats.append(roundInfo)
FullGameInfo =({})
FullGameInfo.update({"team0Name": teamInfo[0].get("name")})#Team 0 Name
FullGameInfo.update({"team1Name": teamInfo[1].get("name")})#Team 1 Name
FullGameInfo.update({"Map":rounds[0].get("map").get("name")})
FullGameInfo.update({"Team0Score":Team0Score})
FullGameInfo.update({"Team1Score":Team1Score})
date = rounds[0].get("timestamp").split('T')[0]
FullGameInfo.update({"Date":date})
writeToExcel(FullGameInfo,rawRoundStats,overallStats,siteStatistics)

# Closing file
f.close()
