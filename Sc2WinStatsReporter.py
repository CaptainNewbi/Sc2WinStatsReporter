import sc2reader
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import os
import threading
import sys
import pythoncom
from win32com.shell import shell, shellcon

def init():
    global outputFile, mostRecentReplayTime, mostRecentReplay, iWon, pathToAccounts, accountList, accountPaths, myRace
    global opponentsRace, window, file
    global matchUp1Wins, matchUp2Wins, matchUp3Wins, matchUp1Total, matchUp2Total, matchUp3Total
    global matchUp1Output, matchUp2Output, matchUp3Output

    pathToAccounts        = filedialog.askdirectory(title="SElECT YOUR STARCRAFT 2 FOLDER: i.e. C:/Users/Newbi/Documents/StarCraft II")
    outputFile            = filedialog.askopenfilename(title="SELECT A TEXT FILE TO WRITE OUTPUT TO.")
    myRace                = "invalid"
    myName                = "invalid"
    opponentsRace         = "invalid"
    mostRecentReplay      = "invalid"
    accountList           = []
    accountPaths          = []
    iWon                  = False

    errorLogging("Starcraft2Directory: ", pathToAccounts, True)
    errorLogging("OutputTextFile: ",outputFile, False)
    
    matchUp1Wins          = 0
    matchUp2Wins          = 0
    matchUp3Wins          = 0
    matchUp1Total         = 0
    matchUp2Total         = 0
    matchUp3Total         = 0
    mostRecentReplayTime  = 0

    matchUp1Output = "XvP: 0/0"
    matchUp2Output = "XvZ: 0/0"
    matchUp3Output = "XvT: 0/0"
    getUserInfo()

    # Clear output file
    file = open(outputFile, "w")
    file.write(matchUp1Output + " \n" + matchUp2Output + " \n" + matchUp3Output)
    file.close()

def errorLogging(string1, string2, makeNewFile):
    global pathToAccounts
    if makeNewFile:
        f = open(pathToAccounts + "/NewbiLog.txt", "w+")
    else:
        f = open(pathToAccounts + "/NewbiLog.txt", "a")
    f.write(string1 + " " + string2 + "\n")
    f.close()

def shortcutTarget (shortcutfile):

    link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
    link.QueryInterface(pythoncom.IID_IPersistFile).Load (shortcutfile)
    targetPath, _ = link.GetPath (shell.SLGP_UNCPRIORITY)
    return targetPath

def getUserInfo():
    global pathToAccounts, accountList, accountPaths
    includedExtensions = ["lnk"]
    dirList        = [fd for fd in os.listdir(pathToAccounts) 
                      if any(fd.endswith(ext) for ext in includedExtensions)]
    lnkList        = [pathToAccounts + "/" + fd for fd in dirList]
    for x in range(0,len(dirList)):
        accountList.append(dirList[x].split("_")[0])
        accountPaths.append(shortcutTarget(lnkList[x]))
    accountPaths = [fd + "/Replays/Multiplayer" for fd in accountPaths]
    for acct in accountPaths:
        errorLogging("PathToReplays: ", acct, False)

def getMostRecentReplay(programStarted):
    global accountPaths, mostRecentReplayTime, mostRecentReplay, window
    errorLogging("Program ", "Running", False)
    included_extensions = ["SC2Replay"]
    for path in accountPaths:
        dir_list    = [path + "/" + fd for fd in os.listdir(path) 
                       if any(fd.endswith(ext) for ext in included_extensions)]
        if(len(dir_list) > 0):
            replayPath = max(dir_list, key=os.path.getmtime)
            replayTime = os.path.getmtime(replayPath)
            if ( replayTime > mostRecentReplayTime):
                mostRecentReplayTime = os.path.getmtime(replayPath)
                mostRecentReplay     = replayPath
                errorLogging("mostRecentReplay: ", mostRecentReplay, False)
                errorLogging("mostRecentReplayTime: ", str(mostRecentReplayTime), False)
                analyzeReplay(programStarted)

    if programStarted:
        window.update()
        window.after(1000,getMostRecentReplay, True)

def analyzeReplay(programStarted):    
    global replay, myRace, opponentsRace, iWon, accountList, mostRecentReplay
    errorLogging("analyzeReplay: ", "Yay!", False)
    replay      = sc2reader.load_replay(mostRecentReplay, load_level=4)
    errorLogging("player1Race: ", replay.people[0].play_race, False)
    errorLogging("player2Race: ", replay.people[1].play_race, False)
    if replay.is_ladder and programStarted:
        errorLogging("is_ladder: ", "True", False)
        errorLogging("winner: ", replay.winner.players[0].name, False)
        player1Race = replay.people[0].play_race
        player2Race = replay.people[1].play_race
        iWon = False
        for myName in accountList:
            if myName == replay.people[0].name:
                myRace        = replay.people[0].play_race
                opponentsRace = replay.people[1].play_race
                if myName == replay.winner.players[0].name:
                    iWon = True
                break
            elif myName == replay.people[1].name:
                myRace        = replay.people[1].play_race
                opponentsRace = replay.people[0].play_race
                if myName == replay.winner.players[0].name:
                    iWon = True
                break
        updateTextFile()
            # Perhaps in the future we can add a check to determine if it was a ladder game
            # and we can exclude non ladder games or make additional stats for ranked/unranked etc
            # isLadder = replay.is_ladder

def updateTextFile():
    global myRace, opponentsRace, iWon
    global matchUp1Wins, matchUp2Wins, matchUp3Wins, matchUp1Total, matchUp2Total, matchUp3Total
    global matchUp1Output, matchUp2Output, matchUp3Output, file

    matchUp1 = myRace[0] + "vP: "
    matchUp2 = myRace[0] + "vZ: "
    matchUp3 = myRace[0] + "vT: "
    
    if (opponentsRace == 'Protoss'):
        if (iWon):
            matchUp1Wins += 1
        matchUp1Total += 1
        matchUp1Output = matchUp1 + str(matchUp1Wins) + "/" + str(matchUp1Total)
    elif (opponentsRace == 'Zerg'):
        if (iWon):
            matchUp2Wins += 1
        matchUp2Total += 1
        matchUp2Output = matchUp2 + str(matchUp2Wins) + "/" + str(matchUp2Total)
    elif (opponentsRace == 'Terran'):
        if (iWon):
            matchUp3Wins += 1
        matchUp3Total += 1
        matchUp3Output = matchUp3 + str(matchUp3Wins) + "/" + str(matchUp3Total)

    file = open(outputFile, "w")
    file.write(matchUp1Output + "\n" + matchUp2Output + "\n" + matchUp3Output )
    file.close()

def startProgram():
    errorLogging("Program ", "started", False)
    theThread = threading.Thread(target=getMostRecentReplay(True))
    theThread.daemon = True
    theThread.start()

def exitProgram(window):
    global file
    file.close()
    window.destroy()
    window.quit()
    
def Sc2WinStatsReporterGui():
    global window
    # Create a window
    window = tk.Tk()
    window.title("Sc2WinTextFileUpdater!")
    window.geometry('300x150')

    # Create ttk widgets
    style = ttk.Style()
    style.map('generalButton', background = [('pressed', 'blue')])

    label = ttk.Label(master = window, 
                      text   ='Created by: CaptainNewbi')
    label.pack()

    # ttk button
    startButton = ttk.Button(master  = window, 
                             text    = 'Start Program', 
                             command = lambda: startProgram())
    
    startButton.place(x     = 40,
                      y     = 50, 
                      width = 100, 
                      height= 50)

    exitButton = ttk.Button(master  = window, 
                            text    = 'Exit Program', 
                            command = lambda: exitProgram(window))
    exitButton.place(x     = 160,
                     y     = 50, 
                     width = 100, 
                     height= 50)

def main():
    init() 
    getMostRecentReplay(False)
    Sc2WinStatsReporterGui()

if __name__ == '__main__':
    main()
    # run
    window.mainloop()
