import sc2reader                                      # Reads Replays
import tkinter as tk                                  # Makes the GUI
from tkinter import filedialog                        # Prompts user for information
from tkinter import ttk                               # Better GUI options
import os                                             # Parsing directory information
import threading                                      # Threads so that program doesnt freeze
import sys                                            # Probably dont need this one
import pythoncom                                      # Parse window Link files
from win32com.shell import shell, shellcon            # Parse window Link files
import time

# Initialize main global variables
def init():
    global mostRecentReplayTime, accountList, accountPaths
    global opponentsRace, window
    global matchUp1Wins, matchUp2Wins, matchUp3Wins
    global matchUp1Total, matchUp2Total, matchUp3Total
    global matchUp1Output, matchUp2Output, matchUp3Output

    myName                = "invalid"
    opponentsRace         = "invalid"

    accountList           = []
    accountPaths          = []

    matchUp1Wins          = 0
    matchUp2Wins          = 0
    matchUp3Wins          = 0
    matchUp1Total         = 0
    matchUp2Total         = 0
    matchUp3Total         = 0
    
    mostRecentReplayTime      = time.time()

    matchUp1Output        = "XvP: 0/0"
    matchUp2Output        = "XvZ: 0/0"
    matchUp3Output        = "XvT: 0/0"
    
    promptUserForInfo()
    getUserInfo()

# Clear output file
def clearOutputFile():
    f = open(outputFile, "w")
    f.write(matchUp1Output + " \n" + matchUp2Output + " \n" + matchUp3Output)
    f.close()

# Ask for starcraft 2 directory and text file to output to
def promptUserForInfo():
    global pathToAccounts, outputFile
    pathToAccounts = filedialog.askdirectory(title="SElECT YOUR STARCRAFT 2 FOLDER: i.e. C:/Users/Newbi/Documents/StarCraft II")
    outputFile     = filedialog.askopenfilename(title="Text file to write to.")

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
    dirList = [fd for fd in os.listdir(pathToAccounts) 
                if any(fd.endswith(ext) for ext in includedExtensions)]
    lnkList = [pathToAccounts + "/" + fd for fd in dirList]
    for x in range(0,len(dirList)):
        accountList.append(dirList[x].split("_")[0])
        accountPaths.append(shortcutTarget(lnkList[x]))
    accountPaths = [fd + "/Replays/Multiplayer" for fd in accountPaths]

# Figures out if a new replay is in any of the account replay folders
def getMostRecentReplay():
    global mostRecentReplayTime, window
    replayInfo          = scanReplayDirectory()
    replayTime          = next(replayInfo)
    mostRecentReplay    = next(replayInfo)
    
    #Is this a new replay?
    if ( replayTime > mostRecentReplayTime):
        mostRecentReplayTime = replayTime
        analyzeReplay(mostRecentReplay)

    # Continue scanning for new replays
    window.update()
    window.after(1000,getMostRecentReplay)

def scanReplayDirectory():
    global accountPaths
    included_extensions = ["SC2Replay"]
    replayTime = 0
    replayPath = ""
    for path in accountPaths:
        dir_list = [path + "/" + fd for fd in os.listdir(path) 
                       if any(fd.endswith(ext) for ext in included_extensions)]
        if(len(dir_list) > 0):
            replayPath = max(dir_list, key=os.path.getmtime)
            replayTime = os.path.getmtime(replayPath)
    yield replayTime
    yield replayPath
    
# Parse the replay for the information desired
def analyzeReplay(mostRecentReplay):    
    global replay, opponentsRace, accountList
    global matchUp1, matchUp2, matchUp3
    replay = sc2reader.load_replay(mostRecentReplay, load_level=4)
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
        if ('myRace' in locals()):
            matchUp1 = myRace[0] + "vP: "
            matchUp2 = myRace[0] + "vZ: "
            matchUp3 = myRace[0] + "vT: "
        updateTextFile(iWon)

# Load Previous session from the output text file
def loadTextFile():
    global outputFile, matchUp1, matchUp2, matchUp3, mostRecentReplayTime
    global matchUp1Wins, matchUp2Wins, matchUp3Wins, matchUp1Total, matchUp2Total, matchUp3Total
    
    f       = open(outputFile, "r")
    line1   = f.readline()
    line2   = f.readline()
    line3   = f.readline()
    f.close()
    
    # Update the saved most recent replay so that we can continue from where we left off
    # Their could be an issue in the future if we use current time everytime
    # If no such issue is thought of, remove these and use init to current time
    replayInfo              = scanReplayDirectory()
    mostRecentReplayTime    = next(replayInfo)
    
    # For now, attempt to at least minimally try to verify the text file 
    if ((line1[1:4] == "vP:" ) or (line1[1:4] == "vZ:" ) or (line1[1:4] == "vT:" )):
        matchUp1        = line1[:4]
        matchUp2        = line2[:4]
        matchUp3        = line3[:4]

        matchUp1Wins    = int(line1[5])
        matchUp2Wins    = int(line2[5])
        matchUp3Wins    = int(line3[5])
        
        matchUp1Total   = int(line1[7])
        matchUp2Total   = int(line2[7])
        matchUp3Total   = int(line3[7])
        return True
    else:
        return False

# Update the text file
def updateTextFile(iWon):
    global opponentsRace, matchUp1, matchUp2, matchUp3
    global matchUp1Wins, matchUp2Wins, matchUp3Wins, matchUp1Total, matchUp2Total, matchUp3Total
    global matchUp1Output, matchUp2Output, matchUp3Output
    if (opponentsRace == 'Protoss'):
        if (iWon):
            matchUp1Wins += 1
        matchUp1Total += 1
    elif (opponentsRace == 'Zerg'):
        if (iWon):
            matchUp2Wins += 1
        matchUp2Total += 1
    elif (opponentsRace == 'Terran'):
        if (iWon):
            matchUp3Wins += 1
        matchUp3Total += 1

    matchUp1Output = matchUp1 + str(matchUp1Wins) + "/" + str(matchUp1Total)
    matchUp2Output = matchUp2 + str(matchUp2Wins) + "/" + str(matchUp2Total)
    matchUp3Output = matchUp3 + str(matchUp3Wins) + "/" + str(matchUp3Total)
    f = open(outputFile, "w")
    f.write(matchUp1Output + "\n" + matchUp2Output + "\n" + matchUp3Output )
    f.close()

# Start program on its own thread
def startProgram():
    theThread = threading.Thread(target=getMostRecentReplay())
    theThread.daemon = True
    theThread.start()

# Close program
def exitProgram(window):
    window.destroy()
    window.quit()

def startButtonCallBack():
    init()
    clearOutputFile()
    startProgram()

def exitButtonCallBack(window):
    exitProgram(window)
    
def previousSessionButtonCallBack():
    init()
    fileLoaded = loadTextFile()
    if not fileLoaded:
        raise ValueError("File not in supported format.")
    startProgram()

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
                             command = lambda: startButtonCallBack())
    
    startButton.place(x     = 40,
                      y     = 30, 
                      width = 100, 
                      height= 50)

    exitButton = ttk.Button(master  = window, 
                            text    = 'Exit Program', 
                            command = lambda: exitButtonCallBack(window))
    exitButton.place(x     = 160,
                     y     = 30, 
                     width = 100, 
                     height= 50)

    loadTextFileButton = ttk.Button(master  = window, 
                           text    = 'Start From Previous Session', 
                           command = lambda: previousSessionButtonCallBack())
    loadTextFileButton.place(x     = 70,
                             y     = 90, 
                             width = 160, 
                             height= 50)

# The most basic attempt at error loging
def errorLogging(string1, string2, makeNewFile):
    global pathToAccounts
    if makeNewFile:
        f = open(pathToAccounts + "/NewbiLog.txt", "w+")
    else:
        f = open(pathToAccounts + "/NewbiLog.txt", "a")
    f.write(string1 + " " + string2 + "\n")
    f.close()
 
# Just because main functions
def main():
    Sc2WinStatsReporterGui()

# Ok, I dont know what I am doing anymore, I think this is what we should do
if __name__ == '__main__':
    main()
    # run
    window.mainloop()
