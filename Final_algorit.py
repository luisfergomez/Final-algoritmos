from random import *
import time
import win32com.client
import os


while True:
    clientes=[" Rosa "," José "," Victor "," Brayan "," Edwin "," Jimena "," Esperanza "," Ana "," Anahí "," Gabriel "," Carlos "," Pablo "," Josie "," Marcos "," Eduardo "]
    indice=randrange(len(clientes))
    persona=clientes[indice]
    print(persona)

    qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    computer_name = os.getenv('COMPUTERNAME')
    qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\valores"
    queue=qinfo.Open(2,0)   # Open a ref to queue
    msg=win32com.client.Dispatch("MSMQ.MSMQMessage")
    msg.Label=persona
    msg.Send(queue)
    time.sleep(1)

    queue.Close()
