import os
import pandas as pd
import warnings
warnings.simplefilter("ignore")

myroot = '/Users/dyu/Downloads/all/president'
os.chdir(myroot)


for x in os.listdir(myroot):
   if ('~' in x ):
	   continue
   elif ('總統' in x) :
       #print('count ', count)
       df = pd.read_excel(x)

       #print('index  ', df.index[7] )
       df2 = df.assign(green=(df['Unnamed: 4']/ df['Unnamed: 6']))
       print('each city 99 ', df2.sort_values(['green'],ascending=True) )
