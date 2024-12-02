import win32com.client as win32
outlook = win32.Dispatch('outlook.application')

import random

def mail(sender,reciever):
      mail = outlook.CreateItem(0)
      mail.To = '{}'.format(sender[1])
      mail.Subject = "It's Secret Santa time!"
      mail.Body = '''
      Hello {},
      your secret santa is {}. Date of exchange is >< and the budget is ><'
      - Computer
                      '''.format(sender[0], reciever[0])
    
      mail.Send()
    

p = [['Name_1', 'Email_1', '1'],
      ['Name_2', 'Email_2', '1'],
      ['Name_3', 'Email_3', '2'],
      ['Name_4', 'Email_4', '3']]


while True:
    random.shuffle(p)
    t = True
    for j in range(len(p)-1):
        if p[j][2] == p[j+1][2]:
            t = False
    if t == True:
        break

for i in range(len(p)):
    mail(p[i % len(p)], p[(i+1) % len(p)])
    #print(p[i % len(p)], p[(i+1) % len(p)])
