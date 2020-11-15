import win32com.client
import os

while True:
    qinfo=win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
    computer_name = os.getenv('COMPUTERNAME')
    qinfo.FormatName="direct=os:"+computer_name+"\\PRIVATE$\\valores"
    queue=qinfo.Open(1,0)   
    msg=queue.Receive()
    valores=msg.Label
    print (msg.Label)
    queue.Close()
    valor=msg.Label


    ubicacion = open("nombres.txt", "a")
    ubicacion.write(valor)
    ubicacion.close()


