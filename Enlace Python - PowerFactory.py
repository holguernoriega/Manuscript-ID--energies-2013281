# -*- coding: utf-8 -*-
"""
Created on Fri Nov  5 20:17:11 2021

@author: Rommel Gallegos; Jose Rivera; Manuel Álvarez; Holguer Noriega
"""
#%% LIBRARY

# DigSILENT-PowerFactory PATH:
import os
os.environ['PATH']=r'C:\Program Files\DIgSILENT\PowerFactory 2020 SP1A'+os.environ['PATH']

# Phyton and  DigSILENT-PowerFactory Link
import sys
sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2020 SP1A\Python\3.8')

# APP Import
import powerfactory as pf


# PANDA LIBRARY IMPORT
import pandas as pd

# TIME MANAGMENT AND PROCESS TIME

import time

# RANDOM LIBRARY IMPORT
import random as rd

# ARRAY LIBRARY IMPORT
import numpy as np

# LIBRARY FOR GRAPHS
#import matplotlib.pyplot as plt

# PROYECT ACTIVATION
app = pf.GetApplication()                          # CALL DigSILENT-PowerFactory

app.Show()                                         # SHOW DigSILENT-PowerFactory SPACE

user = app.GetCurrentUser()
project = app.ActivateProject('simulate39bus')     # FIND PROJECT BY NAME
prj = app.GetActiveProject()                       # ACTIVATE PROJECT ONCE  DigSILENT-PowerFactory IS CHARGED

# ACCES PROJECT DIRECTORIES
# PV PANELS  GENERATION DIRECTORY
pv_dict = {}  
pvs = app.GetCalcRelevantObjects('*.ElmPvsys')     # GET THE LIST OF PV PANELS GENERATORS
for i in pvs:                                      # GET EACH ELEMENT FROM THE DIRECTORY
    pv_dict[i.loc_name] = i

# WIND TURBINE GENERATION DIRECTORY
genst_dict = {}  
gs = app.GetCalcRelevantObjects('*.ElmGenstat')    # GET THE LIST OF WT GENERATORS
for i in gs:                                       # GET EACH ELEMENT FROM THE DIRECTORY
    genst_dict[i.loc_name] = i

# DIRECTORY OF LOADS
load_dict = {}  
loads = app.GetCalcRelevantObjects('*.ElmLod')     # DIRECTORY OF LOADS
for i in loads:                                    # GET EACH ELEMENT FROM THE DIRECTORY
    load_dict[i.loc_name] = i

# DIRECTORY OF SYNC GENERATORS
gen_dict = {}  
gens = app.GetCalcRelevantObjects('*.ElmSym')      # DIRECTORY OF SYNC GENERATORS
for i in gens:                                     # GET EACH ELEMENT FROM THE DIRECTORY
    gen_dict[i.loc_name] = i


# NUMBERS OF BARS WITH PV PANEL AND WT GENERATORS , len(pvs)=len(gs)
n_barras=len(pvs)


# FUNCTION FOR THE INPUT OF PARTICLES
def ingreso_particulas(): #GET BACK THE NUMBER ENTERED BY KEYBOARD, VERIFY IF IT IS INTEGER
    flagParticulas = False
    while not flagParticulas:
        num_p = input('Enter the number of particles: ')
        if num_p.isnumeric():
            flagParticulas = True
        else:
            print("The number entered is not a integer. Try again")
    return int(num_p)


# Function for entering iterations.
def ingreso_iteraciones():
    flagIteraciones = False
    while not flagIteraciones:
        num_p = input('Enter the number of iterations: ')
        if num_p.isnumeric():
            flagIteraciones = True
        else:
            print("The number entered is not a integer. Try again")
    return int(num_p)


# A number is even?
def es_par(numero):
    if numero%2 == 0:
        par = True
    else:
        par = False
    return par


# Function for reading files .txt and get the index
def obtenerDatos(nombre,dfT,i_p,i_ite):
    cod=str(i_p)+str(i_ite)
    tiempo='tiempo'+cod
    freq='freq'+cod

    df=pd.read_csv(nombre,delimiter = "\t",skiprows=1)
    df.columns=[tiempo,freq]
    df[freq] = df[freq].astype(float)
    #In the case the software export the results with  ",", it changes to "."
    #df[freq]=df[freq].str.replace(",",".").astype(float)

    indice=60/df[freq].mean()

    dfT=pd.concat([dfT,df],axis=1)
    return indice,dfT


# Function to get the best local particles
def obtenerBL(dc,strParticula,i_ite):
    listaTemp=['P0',0,0]
    
    for i in range(1,i_ite):
        indice=dc[strParticula][i]['indice']
        
        if(indice>listaTemp[2]):
            listaTemp[0]=strParticula
            listaTemp[1]=i
            listaTemp[2]=indice
    return dc[listaTemp[0]][listaTemp[1]]['config']


# Function to get the best Global Particle Swarm
def obtenerBG(dc,p,ite):
    listaG=['P0',0,0]

    for i in range(1,p+1):
        for j in range(1,ite+1):
            s1="P"+str(i)
            indice=dc[s1][j]['indice']
            if(indice>listaG[2]):
                listaG[0]=s1
                listaG[1]=j
                listaG[2]=indice
    return dc[listaG[0]][listaG[1]]['config']


# Function to get the best Global Optimal Iteration
def obtenerBT(dc,p,ite):
    listaG=['P0',0,0,0]
    for i in range(1,p+1):
        listaTemp=['P0',0,0,0]
        
        for j in range(1,ite+1):
            s1="P"+str(i)
            indice=dc[s1][j]['indice']
            
            if(indice>listaTemp[1]):
                listaTemp[0]=s1
                listaTemp[1]=j
                listaTemp[2]=indice
                listaTemp[3]=dc[listaTemp[0]][listaTemp[1]]['config']
            if(indice>listaG[2]):
                listaG[0]=s1
                listaG[1]=j
                listaG[2]=indice
                listaG[3]=dc[listaG[0]][listaG[1]]['config']

    return listaG


# Events
def shcfolder(p,i):
    shc_folder = app.GetFromStudyCase('IntEvt')  # Abrir la carpeta para crear eventos
    cod=str(p)+str(i)
    shc_folder.CreateObject('EvtSwitch', 'Evento_switch_on'+cod)  # Event: open load
    shc_folder.CreateObject('EvtSwitch', 'Evento_switch_off'+cod)  # Event: close load

    shc_folder.CreateObject('EvtSwitch', '1Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 1
    shc_folder.CreateObject('EvtSwitch', '1Evento_switch_GenSt'+cod)  # Event switch StaGen 1

    shc_folder.CreateObject('EvtSwitch', '2Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 2
    shc_folder.CreateObject('EvtSwitch', '2Evento_switch_GenSt'+cod)  # Event switch StaGen 2

    shc_folder.CreateObject('EvtSwitch', '3Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 3
    shc_folder.CreateObject('EvtSwitch', '3Evento_switch_GenSt'+cod)  # Event switch StaGen 3

    shc_folder.CreateObject('EvtSwitch', '4Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 4
    shc_folder.CreateObject('EvtSwitch', '4Evento_switch_GenSt'+cod)  # Event switch StaGen 4

    shc_folder.CreateObject('EvtSwitch', '5Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 5
    shc_folder.CreateObject('EvtSwitch', '5Evento_switch_GenSt'+cod)  # Event switch StaGen 5

    shc_folder.CreateObject('EvtSwitch', '6Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 6
    shc_folder.CreateObject('EvtSwitch', '6Evento_switch_GenSt'+cod)  # Event switch StaGen 6

    shc_folder.CreateObject('EvtSwitch', '7Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 7
    shc_folder.CreateObject('EvtSwitch', '7Evento_switch_GenSt'+cod)  # Event switch StaGen 7

    shc_folder.CreateObject('EvtSwitch', '8Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 8
    shc_folder.CreateObject('EvtSwitch', '8Evento_switch_GenSt'+cod)  # Event switch StaGen 8

    shc_folder.CreateObject('EvtSwitch', '9Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 9
    shc_folder.CreateObject('EvtSwitch', '9Evento_switch_GenSt'+cod)  # Event switch StaGen 9

    shc_folder.CreateObject('EvtSwitch', '10Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 10
    shc_folder.CreateObject('EvtSwitch', '10Evento_switch_GenSt'+cod)  # Event switch StaGen 10

    shc_folder.CreateObject('EvtSwitch', '11Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 11
    shc_folder.CreateObject('EvtSwitch', '11Evento_switch_GenSt'+cod)  # Event switch StaGen 11

    shc_folder.CreateObject('EvtSwitch', '12Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 12
    shc_folder.CreateObject('EvtSwitch', '12Evento_switch_GenSt'+cod)  # Event switch StaGen 12

    shc_folder.CreateObject('EvtSwitch', '13Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 13
    shc_folder.CreateObject('EvtSwitch', '13Evento_switch_GenSt'+cod)  # Event switch StaGen 13

    shc_folder.CreateObject('EvtSwitch', '14Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 14
    shc_folder.CreateObject('EvtSwitch', '14Evento_switch_GenSt'+cod)  # Event switch StaGen 14

    shc_folder.CreateObject('EvtSwitch', '15Evento_switch_Pvsys'+cod)  # Event switch PV_Sys 15
    shc_folder.CreateObject('EvtSwitch', '15Evento_switch_GenSt'+cod)  # Event switch StaGen 15
    return shc_folder


# Print the optimal global
    listaG=obtenerBT(dicc,particulas,iteraciones)
    print(" Best global Particle Swarm: ",listaG[3])
    print("The Index is: ",listaG[2])
    print("\n")
    print("The optimal configuration is when the following generators are connected:\n")
    listaFotovoltaicos=[]
    listaEolicos=[]
    print(len(listaG[3]))
    for i in range(0,len(listaG[3])//2):

        if listaG[3][i*2]==1:
            listaFotovoltaicos.append(((i*2)//2)+1)

        if listaG[3][i*2+1]==1:
            listaEolicos.append(((i*2)//2)+1)

    print("PV PANELS GENERATORS: ",listaFotovoltaicos)
    print("WIND TURBINE GENERATORS: ",listaEolicos)
    print("\n")


iteraciones=ingreso_iteraciones()
particulas=ingreso_particulas()
                                    

# Function for the events created after the first iteration
dfT=pd.DataFrame()
dc = {}
for i_ite in range(1,iteraciones+1):
    for i_p in range(1,particulas+1):
        shc_folder=shcfolder(p=i_p,i=i_ite)
        eventos = shc_folder.GetContents()  # Get the list of events

        strParticula = "P"+str(i_p)
        # Asignation of the events after the first iteration
        if i_ite==1:
            dc[strParticula]={}
            dc[strParticula][i_ite]={}
            # Parameter adjustment  for disconnection of load
            eventos[0].time = 0.25                                    # open time a t = 0.25 s
            eventos[0].i_switch = 0                                   # Open load
            eventos[0].p_target = load_dict['Load39']                 #  Load39 is called

            #  Parameter adjustment  for connection of load
            eventos[1].time = 0.5                                     # close time a t = 0.50 s
            eventos[1].i_switch = 1                                   # close load
            eventos[1].p_target = load_dict['Load39']                 # Load39 is called
            
            nPv=1
            nGen=1
            listaTempX = []
            listaTempV = []
            for i in range(2,2*n_barras+2):

                if es_par(i):
                    var1_pv = rd.randrange(0, 2)
                    listaTempX.append(var1_pv)
                    listaTempV.append(0)
                    eventos[i].time = 0                                          # close time del sw t = 0 s
                    eventos[i].i_switch = var1_pv                                # pv 1 connect y 0 disconnect
                    eventos[i].p_target = pv_dict['PV_Sys'+str(nPv)]
                    nPv+=1

                else:
                    var1_gen = rd.randrange(0,2)

                    listaTempX.append(var1_gen)
                    listaTempV.append(0)
                    eventos[i].time = 0                                          # close time del sw t = 0 s
                    eventos[i].i_switch = var1_gen                               # wt 1 connect y 0 disconnect
                    eventos[i].p_target = genst_dict['StaGen'+str(nGen)]
                    nGen+=1
            dc[strParticula][i_ite]['config']=listaTempX
            dc[strParticula][i_ite]['vel']=listaTempV
        # PSO METHOD point start
        else:
            dc[strParticula][i_ite]={}
            # Parámtros del método PSO
            w=0.4                                       # Inertia factor
            vant=dc[strParticula][i_ite-1]['vel']       # Last particle speed
            xant=dc[strParticula][i_ite-1]['config']    # Last particle Positión
            c1=c2=2                                     # Attraction Constant
            r1=rd.randrange(1,2*n_barras+1)             # Random number
            r2=rd.randrange(1,2*n_barras+1)             # Random number
            bestL=obtenerBL(dc,strParticula,i_ite)      # Best Local particle

            # Speed calculation
            b1= [i * w for i in vant]

            con1=c1*r1
            l1=list(np.array(bestL) - np.array(xant))
            b2 = [i * con1 for i in l1]

            con2=c2*r2
            l2=list(np.array(bestG) - np.array(xant))
            b3 = [i * con2 for i in l2]

            velocidad = list(np.array(b1)+np.array(b2)+np.array(b3))  # New speed
            for i in range(len(velocidad)):
                if velocidad[i]>0:
                    velocidad[i]=1
                else:
                    velocidad[i]=0

            posicion=list(np.array((xant) + np.array(velocidad)))     # New Position
            for i in range(len(posicion)):
                posicion[i] = int(posicion[i])
                if posicion[i]>0:
                    posicion[i]=1
                else:
                    posicion[i]=0



            dc[strParticula][i_ite]['config'] = posicion
            dc[strParticula][i_ite]['vel'] = velocidad
            # Desconexion de carga
            eventos[0].time = 0.25                     # open time a t = 0.25 s
            eventos[0].i_switch = 0                    # open load
            eventos[0].p_target = load_dict['Load39']  #  Load39 is called
            # Conexion de carga
            eventos[1].time = 0.5                      # close time a t = 0.25 s
            eventos[1].i_switch = 1                    # close load
            eventos[1].p_target = load_dict['Load39']  #  Load39 is called

            nPv=1
            nGen=1
            for i in range(2,2*n_barras+2):

                if es_par(i):
                    eventos[i].time = 0                                          # close time a t = 0.25 s
                    eventos[i].i_switch = posicion[i-2]                          # pv 1 connect y 0 disconnect
                    eventos[i].p_target = pv_dict['PV_Sys'+str(nPv)]
                    nPv+=1

                else:
                    eventos[i].time = 0                                          # tiempo de cierre del sw t = 0 s
                    eventos[i].i_switch = posicion[i-2]                          # wt 1 connect y 0 disconnect
                    eventos[i].p_target = genst_dict['StaGen'+str(nGen)]
                    nGen+=1



    #%% STUDY CASE

        # Reset COMPUTATION
        #app.ResetCalculation()
        
        # Leer y modificar archivos de resultado
        elmres = app.GetFromStudyCase('All calculations.ElmRes')       # Call the result files

        # Condiciones iniciales
        ini = app.GetFromStudyCase('ComInc')                           # Initial conditions

        ini.Execute()                                                  # STUDY CASE EXECUTION
        
        # Simulacion Dinánmica
        sm = app.GetFromStudyCase('ComSim')                            # Dinamic simulation

        sm.tstop = 20                                                  # simulation time  t = 20s
        
        sm.Execute()                                                   # STUDY CASE EXECUTION


        # Export from DigSILENT-PowerFactory to .txt
        comres = app.GetFromStudyCase('ComRes')
        comres.iopt_csel = 0
        comres.iopt_tsel = 1
        comres.iopt_locn = 2
        comres.ciopt_head = 1
        comres.pResult = elmres
        comres.iopt_exp = 4
        comres.f_name = r'C:\Users\CltControl\Desktop\JG\particula' + str(i_p)+str(i_ite) + '.txt'
        comres.Execute()



        time.sleep(1)       # Configuration time
        elmres.Clear()
        nombreA='particula'+str(i_p)+str(i_ite)+".txt"

        indice,dfT=obtenerDatos(nombreA,dfT,i_p,i_ite) #Get the index and dfTotal from txt 
        dc[strParticula][i_ite]['indice']=indice
        #shc_folder.Delete()
    bestG=obtenerBG(dc,particulas,i_ite)


listaG=obtenerBT(dc,particulas,iteraciones)

resultados(dc,iteraciones)
