from tkinter import Button
from tkinter import Label
from tkinter import Tk
from tkinter import filedialog

import win32com.client


# Constants
FONT_SIZE_MIN = 6
FONT_SIZE_MAX = 35
FONT_SIZE_STEP = 2
FONT_SIZE_LABEL_TEXT = "Schriftgröße: "

wordMatrix = []
MAX_WORD_LENGTH = 30
MIN_WORD_LENGTH = 3
WORD_LENGTH_LABEL_TEXT = "Wortlänge: "
wordLength = 5
textFileName = ""
wordMatrixInitialized = False

fontSize = 25
wordFont = "Arial"
counter = None

score = 0

NEXT_WORD_BUTTON_TEXT = "Nächstes Wort (Enter)"
READ_WORD_BUTTON_TEXT = "(v)orlesen"
READ_FILE_BUTTON_TEXT = "Datei lesen"
EXIT_BUTTON_TEXT      = "Beenden"
SCORE_LABEL_TEXT      = "Punkte: "

REPLACEMENT_SYMBOLS = [".", ",", "\"", "?", "!", ":", ";", "-", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "(", ")", "\r", "\n", "„"]

# Ein Fenster erstellen
fenster = Tk()
stimme = win32com.client.Dispatch("SAPI.SpVoice")
#vcs = stimme.GetVoices()
#print(vcs.Item (1) .GetAttribute ("Name")) # speaker name

# Die folgende Funktion soll ausgeführt werden, wenn
# der Benutzer den Button anklickt

def initWordMatrix() :
    global wordMatrixInitialized

    for i in range(MAX_WORD_LENGTH) :
        wordMatrix.append([])
    wordMatrixInitialized = True

# Read the text from the given file name
def readFile(fileName) : 
    text = ""
    print(fileName)
    fobj = open(fileName)
    for line in fobj:
        text = text + " " + line
    fobj.close()
    return text
    
# Strip all unnecessary characters from the text,
# like ., ,, :, (, ), ? etc.
def cleanText(text) : 
    text.strip()
    for s in REPLACEMENT_SYMBOLS :
        text = text.replace(s, "")

    return text

# Insert word accordig to length into row.
# A word needs to have at least 3 letters
def putToWordMatrix(word) :
    length = len(word.strip())

    if length >= MIN_WORD_LENGTH and length <= MAX_WORD_LENGTH :
        # check for existence and for spelling 
        found = False
        for w in wordMatrix[length] :
            if w.lower() == word.lower() :
                if w.isupper and word.islower :
                    wordMatrix[length].remove(w)
                    wordMatrix[length].append(word)
                found = True

        if not found :
            wordMatrix[length].append(word)

# split the text in single words and send word to
# insertion method
def transform(text) :
    words = text.split(" ")

    for word in words :
        putToWordMatrix(word)

# Create the word matrix from given file. 
# The words of the file are inserted into the word matrix 
# All words of a certain length in a single row
def loadText(fileName) : 
    initWordMatrix()
    text = readFile(fileName)
    text = cleanText(text)
    transform(text)

def change_fontsize(newFontSize) :
    global fontSize

    # do the change
    fontSize = newFontSize
    fontsize_label.config(text=FONT_SIZE_LABEL_TEXT + str(fontSize))

# Methods for Button action
# Actions on Font Size
def decrease_fontsize() : 
    # check if a change is allowed
    if fontSize <= FONT_SIZE_MIN : return 

    change_fontsize(fontSize - FONT_SIZE_STEP)

def increase_fontsize() : 
    # check if a change is allowed
    if fontSize >= FONT_SIZE_MAX : return 

    change_fontsize(fontSize + FONT_SIZE_STEP)

# Actions on word length
def increase_word_length() : 
    global wordLength
    wordLength = wordLength + 1

    # do the change in the GUI
    word_length_label.config(text=WORD_LENGTH_LABEL_TEXT + str(wordLength))

def decrease_word_length() : 
    global wordLength
    if wordLength <= 3 : return 

    wordLength = wordLength - 1

    # do the change in the GUI
    word_length_label.config(text=WORD_LENGTH_LABEL_TEXT + str(wordLength))

def getCurrentWord() : 
    if not wordMatrixInitialized : return ""

    return wordMatrix[wordLength][counter]

def getNextWord() :
    global counter

    if counter == None : 
        counter = 0
    else :
        counter = counter + 1
        if counter >= len(wordMatrix[wordLength]) :
            counter = 0

    word = getCurrentWord()

    return word

def next_word(event=None) :
    global score
    
    word_label.config(text=getNextWord(), font=(wordFont, fontSize))
    score = score + wordLength
    score_label.config(text=SCORE_LABEL_TEXT + str(score), font=("Arial", 12))

def browse_button():
    global textFileName
    filename = filedialog.askopenfilename()
    print(filename)
    loadText( filename )

def readWord() :
    stimme.Speak( getCurrentWord() )



# Den Fenstertitle erstellen
fenster.title("Lese Training")

# Label und Buttons erstellen

nextWord_Button = Button(fenster, text=NEXT_WORD_BUTTON_TEXT, command=next_word)
fenster.bind('<Return>', next_word)

readWord_Button = Button(fenster, text=READ_WORD_BUTTON_TEXT, command=readWord)
fenster.bind('v', readWord)

readFile_Button = Button(fenster, text=READ_FILE_BUTTON_TEXT, command=browse_button)
fenster.bind('r', browse_button)

exit_button = Button(fenster, text=EXIT_BUTTON_TEXT, command=fenster.quit)

fontsize_increase_button = Button(fenster, text="+", command=increase_fontsize)
fontsize_decrease_button = Button(fenster, text="-", command=decrease_fontsize)

change_font_button = Button(fenster, text="<Font>")

word_label = Label(fenster, text=getCurrentWord(), font=(wordFont, fontSize))
fontsize_label = Label(fenster, text=FONT_SIZE_LABEL_TEXT + str(fontSize))
font_label = Label(fenster, text="Schriftart")

word_length_label = Label(fenster, text=WORD_LENGTH_LABEL_TEXT + str(wordLength))
word_length_increase_button = Button(fenster, text="+", command=increase_word_length)
word_length_decrease_button = Button(fenster, text="-", command=decrease_word_length)

score_label = Label(fenster, text=SCORE_LABEL_TEXT + str(score), font=("Arial", 12))

# Nun fügen wir die Komponenten unserem Fenster 
# in der gwünschten Reihenfolge hinzu.

# word to read
word_label.grid(row=3, column=1, padx = 40)

# Menu
readFile_Button.grid(row=1, column=3, pady = 20, padx=20)

word_length_label.grid(          row=2, column=3, pady = 20, padx=20)
word_length_increase_button.grid(row=2, column=4, pady = 5, padx=5)
word_length_decrease_button.grid(row=2, column=5, pady = 5, padx=5)

fontsize_label.grid(          row=3, column=3, pady = 20, padx=20)
fontsize_increase_button.grid(row=3, column=4, pady = 5, padx=5)
fontsize_decrease_button.grid(row=3, column=5, pady = 5, padx=5)

font_label.grid(        row=4, column=3, pady = 5, padx=5)
change_font_button.grid(row=4, column=4, pady = 20, padx=20)

readWord_Button.grid(row=5, column=3, pady = 20, padx=20)

nextWord_Button.grid(row=6, column=3, pady = 20, padx=20)
score_label.grid(    row=6, column=4, pady = 20, padx=20)
exit_button.grid(    row=6, column=5, pady = 20, padx = 20)
# In der Ereignisschleife auf Eingabe des Benutzers warten.
fenster.mainloop()