def getcolor(x):
   y=float(x)
   if y > 30:
    return '8F6538'
   elif y <=30 and y > 10:
    return '468657'
   elif y <=10 and y >6:
    return '529C64'
   elif y <=6 and y >5:
    return '63BE7B'
   elif y <=5 and y >4:
    return '7DC67D'
   elif y <=4 and y >3:
    return '98CE7F'
   elif y <=3 and y >2:
    return 'B1D580'
   elif y <=2 and y >1:
    return 'CCDD82'
   elif y <=1 and y >0:
    return 'E5E483'
   elif y <=0 and y > -1:
    return 'FFEB83'
   elif y <=-1 and y >-2:
    return 'FDD57F'
   elif y <=-2 and y >-3:
    return '63BE7B'
   elif y <=-3 and y >-4:
    return 'FCBF7B'
   elif y <=-4 and y >-5:
    return 'FA9473'
   elif y <=-5 and y >-6:
    return 'F97E6F'
   elif y <=-6 and y >-10:
    return 'F97E6F'
   elif y <=-10 and y >-30:
    return 'F80013'
   elif y <=-30:
    return 'E317F8'
   else:
    return 'FFFFFF'

