import sc2reader                                      # Reads Replays
import tkinter as tk                                  # Makes the GUI
from tkinter import filedialog                        # Prompts user for information
from tkinter import ttk                               # Better GUI options
import os                                             # Parsing directory information
import threading                                      # Threads so that program doesnt freeze
import sys                                            # Probably dont need this one
import pythoncom                                      # Parse window Link files
from win32com.shell import shell, shellcon            # Parse window Link files

# Initialize main global variables and output text file
def init():
    global outputFile, mostRecentReplayTime, mostRecentReplay, iWon, pathToAccounts, accountList, accountPaths, myRace
    global opponentsRace, window, programStarted
    global matchUp1Wins, matchUp2Wins, matchUp3Wins, matchUp1Total, matchUp2Total, matchUp3Total
    global matchUp1Output, matchUp2Output, matchUp3Output

    pathToAccounts        = filedialog.askdirectory(title="SElECT YOUR STARCRAFT 2 FOLDER: i.e. C:/Users/Newbi/Documents/StarCraft II")
    outputFile            = filedialog.askopenfilename(title="Text file to write to.")
    myRace                = "invalid"
    myName                = "invalid"
    opponentsRace         = "invalid"
    mostRecentReplay      = "invalid"
    accountList           = []
    accountPaths          = []
    iWon                  = False
    programStarted        = False

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
    f = open(outputFile, "w")
    f.write(matchUp1Output + " \n" + matchUp2Output + " \n" + matchUp3Output)
    f.close()

# Parse window link files
# Input: path including lnk file
# Output: Target Path of the lnk file
def shortcutTarget (shortcutfile):

    link = pythoncom.CoCreateInstance(shell.CLSID_ShellLink, None, pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink)
    link.QueryInterface(pythoncom.IID_IPersistFile).Load (shortcutfile)
    targetPath, _ = link.GetPath (shell.SLGP_UNCPRIORITY)
    return targetPath

# Figure outs who the players ID and all of their accounts and creates a list of paths to the replay folders
# This is so that we can "monitor" those folders for the most recent replay
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

# Figures out if a new replay is in any of the account replay folders
def getMostRecentReplay():
    global accountPaths, mostRecentReplayTime, mostRecentReplay, window, programStarted
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
                if programStarted:
                    analyzeReplay()

    if programStarted:
        window.update()
        window.after(1000,getMostRecentReplay)

# Parse the replay for the information desired
def analyzeReplay():    
    global replay, myRace, opponentsRace, iWon, accountList, mostRecentReplay
    replay      = sc2reader.load_replay(mostRecentReplay, load_level=4)
    if replay.is_ladder:
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

# Update the text file
def updateTextFile():
    global myRace, opponentsRace, iWon
    global matchUp1Wins, matchUp2Wins, matchUp3Wins, matchUp1Total, matchUp2Total, matchUp3Total
    global matchUp1Output, matchUp2Output, matchUp3Output

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

    f = open(outputFile, "w")
    f.write(matchUp1Output + "\n" + matchUp2Output + "\n" + matchUp3Output )
    f.close()

# Start program on its own thread
def startProgram():
    global programStarted
    programStarted = True
    theThread = threading.Thread(target=getMostRecentReplay())
    theThread.daemon = True
    theThread.start()

# Close program
def exitProgram(window):
    window.destroy()
    window.quit()

# Main Gui Interface
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

# Just because main functions
def main():
    init() 
    getMostRecentReplay()
    Sc2WinStatsReporterGui()

# Ok, I dont know what I am doing anymore, I think this is what we should do
if __name__ == '__main__':
    main()
    # run
    window.mainloop()
