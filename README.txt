README
A.	INITIAL CONFIGURATION 

1.	Import and activate the electrical design in pfd format with DIgSILENT PowerFactory. The visualization of the power system schematic will be available after the import of the file and the activation. 
2.	Open the Python file with a Python IDE - could be PyCharm. Before using this IDE, be sure about the activation of the PowerFactory electrical design at least one time. 
3.	Locate the path of the Python folder which you are working. It is inside of the folder DigSilent Power Factory program files. This path must be copied and pasted in the sys.path.append line code. This line is in the beginning of the Python code. 

Example:
sys.path.append (r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP5\Python\3.8')

4.	The path copied in the step before, pasted again but in the os.environ [‘PATH’]  line code wich is located before of sys.path.append line.  This path mustn’t be contained the Python name and version. 
Example: 
os.environ['PATH'] = r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP5'+ os.environ['PATH']

B.	LOCATION OF THE TXT FILES WHICH WILL BE THE PSO PARTICLES

1.	Created a folder in your PC
2.	Copy the folder path in the comres.f_name line code. This line is in the last part of the Python code. 
Example: 
comres.f_name = r'C:\Users\hnori\OneDrive\ESPOL 2022\PROYECTO DE INV MA\Nuevo Proyecto Gallegos Rivera\Código y Simulación 1\ARCHIVOS TXT' + str( i_p) + str(i_ite) + '.txt'

3.	Be sure the fact of many computers recognize the point instead of comma for the frequency values in txt file. If this happens, please use the following instruction: 
a.	df[freq] = df[freq].astype(float) for comma uses in frequency; or
b.	df[freq]=df[freq].str.replace(",",".").astype(float) for point uses in frequency. 

C.	EXECUTION OF THE SIMULATION 

1.	Run the Python file. Wait until the indexation finishs. 
2.	The DigSilent schematic will open automatically. Here, be sure about of having nothing in Simulation Events/ Faults, before defining 
3.	Define the number of particles and iterations.

D.	INTERPRETATION OF RESULTS
1.	The results show which PV generator and Wind turbine generator should be connected in order to get the optimal result. They are the PSO results. 
2.	The frequency index is also shown. 

