#-*- coding:Utf-8 -*-
import win32api, win32con
import os
import time
import ImageOps
import ImageGrab
import sys, pygame
import random
import win32com.client
import webbrowser
from numpy import *

def leftClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
    time.sleep(.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)

def mousePos(cord):
    win32api.SetCursorPos((cord[0],cord[1]))

def get_cords():
    x,y = win32api.GetCursorPos()
    return(x,y)


def screengrab():
    im = ImageGrab.grab()
    im.save(os.getcwd() + '\\Grab001.png', 'PNG')
    return im


##Recuperation des configs dans le fichier option

Config= open("option.txt","r")
x=Config.readline()
TD=int(Config.readline())
x=Config.readline()
Panorama=int(Config.readline())
x=Config.readline()
UE=int(Config.readline())
x=Config.readline()
Temps=float(Config.readline())
Config.close()

##Affichage des options sélectionnées

if TD==1:
    print"TD : A"
elif TD==2:
    print"TD : B"
    
if Panorama==1:   
    print"Panorama : Chimie"
elif Panorama==2:
    print"Panorama : Maths"
elif Panorama==3:
    print"Panorama : Physique1"
elif Panorama==4:
    print"Panorama : Physique2"

if UE==1:    
    print"UE : Panorama du vivant"
elif UE==2:
    print"UE : Planete Terre"

x_decal=0
y_decal=0

##Ouverture de la page internet de l'ENT

webbrowser.open('http://ent.univ-lemans.fr/tag.c690c78d41db7141.render.userLayoutRootNode.uP?uP_sparam=focusedTabID&focusedTabID=48&uP_sparam=mode&mode=view')
print"!!!!!!!!! NE PAS TOUCHEZ LA SOURIS !!!!!!!!!!"
time.sleep(3)

##Modification des coordonnées en fonction de la taille de la page internet

while x_decal==0:
    im = screengrab()
    for y in range (0,400):
        for x in range(0,400):
            if im.getpixel((x,y)) ==(76,130,155) and x_decal==0: ##couleur de l'onglet "recherche"
                x_decal=x
                y_decal=y
                break

mousePos((10+x_decal,67+y_decal))  ##Etudiants
leftClick()
time.sleep(Temps)
mousePos((19+x_decal,187+y_decal)) ##UFR science
leftClick()
time.sleep(Temps)
mousePos((28+x_decal,317+y_decal)) ##Licence
leftClick()
time.sleep(Temps)
mousePos((37+x_decal,330+y_decal)) ##L1
leftClick()
time.sleep(Temps)
mousePos((46+x_decal,355+y_decal)) ##L1 MPCE2I
leftClick()
time.sleep(Temps)
mousePos((55+x_decal,368+y_decal)) ##Sem 1
leftClick()
time.sleep(Temps)

mousePos((64+x_decal,409+y_decal)) ##UE
leftClick()
time.sleep(Temps)

if UE==1:
    mousePos((88+x_decal,423+y_decal))
    leftClick()
    time.sleep(Temps)
elif UE==2:   
    mousePos((88+x_decal,435+y_decal))
    leftClick()
    time.sleep(Temps)

mousePos((77+x_decal,397+y_decal)) ##TP
leftClick()
time.sleep(Temps)

mousePos((64+x_decal,381+y_decal)) ##TD MPCE2I
leftClick()
time.sleep(Temps)

if TD==1:
    mousePos((88+x_decal,409+y_decal))
    leftClick()
    time.sleep(Temps)
elif TD==2:
    mousePos((88+x_decal,422+y_decal))
    leftClick()
    time.sleep(Temps)


mousePos((73+x_decal,394+y_decal)) ##Groupes "Panorama des Sciences"
leftClick()
time.sleep(Temps)



if Panorama==1:
    mousePos((97+x_decal,408+y_decal))
    leftClick()
    time.sleep(Temps)
elif Panorama==2:   
    mousePos((97+x_decal,422+y_decal))
    leftClick()
    time.sleep(Temps)
elif Panorama==3:   
    mousePos((97+x_decal,434+y_decal))
    leftClick()
    time.sleep(Temps)
elif Panorama==4:   
    mousePos((97+x_decal,448+y_decal))
    leftClick()
    time.sleep(Temps)



print"Emplois du Temps affiche !"
print"By Cayon Etienne "

