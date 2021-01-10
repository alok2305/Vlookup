import pandas as pd
import numpy as np
import keyword
from difflib import get_close_matches
val1=input("Enter the name of the Datasheet file With .xlsx Extension ")
val2=input("Enter the name of the reference attribute sheet file With .xlsx Extension ")
df=pd.read_excel(val1,usecols=["Lock make"])
pf=pd.read_excel(val2,usecols=['Lock make (Dist PO Sheet)'])
er =pd.read_excel(val1,usecols=['Profile'])
tr=pd.read_excel(val2,usecols=['PROFILE (Dist PO Sheet)'])
td=pd.read_excel(val2,usecols=['COLOUR (Dist PO Sheet)'])
ds=pd.read_excel(val1,usecols=['Finish code'])
orien=pd.read_excel(val2,usecols=['Orientation (Dist PO)'])
ori=pd.read_excel(val1,usecols=['Door Orientation'])
design=pd.read_excel(val2,usecols=['Design (Dist PO Sheets)'])
desi=pd.read_excel(val1,usecols=['Embossed Sheet Design Code/Plain Sheet'])
peephole=pd.read_excel(val2,usecols=['Peephole (Dist PO)'])
peep=pd.read_excel(val1,usecols=['PeepHole'])
def clearance(data):
    data.dropna(how='all',inplace=True)
    for i in range(0,len(data)):
        x=str(data.iloc[i][0]).endswith(" ")
        if(x):
            (data.iloc[i][0])=str(data.iloc[i][0]).rstrip()
            continue
        else:
            continue
    data.drop_duplicates(inplace=True)
    data.set_index([pd.Index(range(0,len(data)))],inplace=True)
    return data
def seperation(dataset,datapo):
    y=pd.Series()
    c=pd.Series()
    for i in range(0,len(dataset)):
        for j in range(0,len(datapo)):
            x=(str(pd.Series(dataset.iloc[i].values))in str(pd.Series(datapo.iloc[j].values)))
            if(x==True):
                y=y.append(pd.Series([dataset.iloc[i].values]))
        if(str(dataset.iloc[i].values)not in str(datapo.iloc[:].values)):
            c=c.append(pd.Series([dataset.iloc[i].values]))
        
    present=pd.DataFrame(y)
    notpresent=pd.DataFrame(c)
    present.set_index([pd.Index(range(0,len(y)))],inplace=True)
    notpresent.set_index([pd.Index(range(0,len(c)))],inplace=True)
    return present,notpresent
def removal(data2):
    if(len(data2)==0):
        return data2
    else:       
        for i in range(0,len(data2)):
            data2[i]=data2.iloc[i][0]
    return data2
def upperconv(datainput):
    for i in range(0,len(datainput)):
        datainput.dropna(inplace=True)
        datainput[i]=datainput[i].upper()
        continue
def setvalue(datainput):
    datainput.dropna(inplace=True)
    if(datainput.empty != True):
        upperconv(datainput)
        datainput=datainput.str.replace("*","")
        datainput=datainput.str.replace("X","")
        datainput=datainput.str.replace("x","")
        datainput.drop_duplicates(inplace=True)
        datainput.set_axis(range(0,len(datainput)),inplace=True)
        return datainput
    else:
        return datainput
def seperation2(dataset,datapo):
    c=[]
    if(dataset.empty!=True):
        for data in dataset:
            if(data not in datapo.values):
                c.append(data)
        if(len(c)==0):   
            c.append(np.nan)
            notpresent=pd.DataFrame(c) 
            return notpresent
        else:
            notpresent=pd.DataFrame(c)
            return notpresent
    else:
        c.append(np.nan)
        notpresent=pd.DataFrame(c)
        return notpresent
def check(dataup,datanotup):
    c=[]
    dataup.dropna(inplace=True)
    if(dataup.empty!=True):
        for i in range(0,len(dataup)):
            for j in range(0,len(datanotup)):
                upr= datanotup[j].upper()
                if('X' in upr):
                    specials=upr.replace('X',"")
                elif('*' in upr):
                    specials=upr.replace('*',"")
                else:
                    specials=upr
                if(dataup[i]==specials):
                    c.append(j)
                    break
            continue
        key=pd.DataFrame(c)
        return key
    else:
        c.append(np.nan)
        key=pd.DataFrame(c)
        return key
def myfunctions(value,datainput):
    for i in range(0,len(value)):
        count=0
        percentu=[]
        if(len(value[i][0])<=len(datainput[i])):
            for j in range(0,len(value[i][0])):
                if(value[i][0][j].lower()==datainput[i][j].lower()):
                    count=count+1
            percentu.append((count/len(datainput[i]))*100)
        else:
            for j in range(0,len(datainput[i])):
                     if(value[i][0][j].lower()==datainput[i][j].lower()):
                           count=count+1
            percentu.append((count/len(datainput[i]))*100)       
        percentage.append(percentu)    
def merging(datalock,datanot):
    c=[]
    datalock.dropna(inplace=True)
    if(datalock.empty!=True):
        for i in range(0,len(datalock)):
            a=datanot[datalock[i]]
            c.append(a)
        newf=pd.DataFrame(c)  
        return newf
    else:
        c.append(np.nan)
        newf=pd.DataFrame(c)
        return newf   
def closeMatches(patterns, word):
    a=(get_close_matches(word, patterns,cutoff=0.5,n=1))
    return a
def closematches(datainput,datapo):
    if __name__ == "__main__":
        eachvalues=[]
        percentage=[]
    for k in range(0,len(datainput)):
        if(len(pd.Series(datainput).dropna())!=0):
            percentu=[]
            word = datainput[k]
            patterns =datapo 
            list1=closeMatches(patterns, word)
            eachvalues.append(list1)
    return eachvalues     
df["Lock make"]=df["Lock make"].str.replace(" ","")
pf["Lock make (Dist PO Sheet)"]=pf["Lock make (Dist PO Sheet)"].str.replace(" ","")
er["Profile"]=er["Profile"].str.replace(" ","")
tr["PROFILE (Dist PO Sheet)"]=tr["PROFILE (Dist PO Sheet)"].str.replace(" ","")
td["COLOUR (Dist PO Sheet)"]=td["COLOUR (Dist PO Sheet)"].str.replace(" ","")
ds["Finish code"]=ds["Finish code"].str.replace(" ","")
orien["Orientation (Dist PO)"]=orien["Orientation (Dist PO)"].str.replace(" ","")
ori["Door Orientation"]=ori["Door Orientation"].str.replace(" ","")
design['Design (Dist PO Sheets)']=design['Design (Dist PO Sheets)'].str.replace(" ","")
desi['Embossed Sheet Design Code/Plain Sheet']=desi['Embossed Sheet Design Code/Plain Sheet'].str.replace(" ","")
peephole['Peephole (Dist PO)']=peephole['Peephole (Dist PO)'].str.replace(" ","")
peep['PeepHole']=peep['PeepHole'].str.replace(" ","")
df=clearance(df)
pf=clearance(pf)
er=clearance(er)
tr=clearance(tr)
td=clearance(td)
ds=clearance(ds)
orien=clearance(orien)
ori=clearance(ori)
design=clearance(design)
desi=clearance(desi)
peephole=clearance(peephole)
peep=clearance(peep)
df,newdf=seperation(df,pf)
er,newer=seperation(er,tr)
ds,newds=seperation(ds,td)
ori,newori=seperation(ori,orien)
desi,newdesi=seperation(desi,design)
peep,newpeep=seperation(peep,peephole)
df.columns=['Present Locks']
df["New Locks"]=newdf
df["Present Profile"]=er
df["New Profile"]=newer
df["Present Colour"]=ds
df["New Colour"]=newds
df["Present Orientation"]=ori
df["New Orientation"]=newori
df['Present Design']=desi
df['New Design']=newdesi
df["Present PeepHole"]=peep
df['New PeepHole']=newpeep
del df["Present Locks"]
del df["Present Profile"]
del df["Present Colour"]
del df["Present Orientation"]
del df['Present Design']
del df['Present PeepHole']
df["New Locks"].dropna(inplace=True)
df["New Profile"].dropna(inplace=True)
df["New Colour"].dropna(inplace=True)
df["New Orientation"].dropna(inplace=True)
df["New Design"].dropna(inplace=True)
df["New PeepHole"].dropna(inplace=True)
aa=df["New Locks"]
bb=df["New Profile"]
cc=df["New Colour"]
dd=df["New Orientation"]
ee=df["New Design"]
ff=df["New PeepHole"]
df["New Locks"]=removal(aa)
df["New Profile"]=removal(bb)
df["New Colour"]=removal(cc)
df["New Orientation"]=removal(dd)
df["New Design"]=removal(ee)
df["New PeepHole"]=removal(ff)
df.to_excel('Output1.xlsx',index=False,columns=["New Locks","New Profile","New Colour",'New Orientation','New Design','New PeepHole'],encoding='utf-8')
newlock =df["New Locks"].copy()
newprofile=df["New Profile"].copy()
newcolour=df["New Colour"].copy()
neworien=df["New Orientation"].copy()
newdesign=df["New Design"].copy()
newpeep=df["New PeepHole"].copy()
polock=pf["Lock make (Dist PO Sheet)"].copy()
poprofile=tr['PROFILE (Dist PO Sheet)'].copy()
pocolour=td["COLOUR (Dist PO Sheet)"].copy()
poori=orien["Orientation (Dist PO)"].copy()
podesi=design['Design (Dist PO Sheets)'].copy()
popeep=peephole['Peephole (Dist PO)'].copy()
newlock=setvalue(newlock)
newcolour=setvalue(newcolour)
newprofile=setvalue(newprofile)
neworien=setvalue(neworien)
newdesign=setvalue(newdesign)
newpeep=setvalue(newpeep)
polock=setvalue(polock)
poprofile=setvalue(poprofile)
pocolour=setvalue(pocolour)
poori=setvalue(poori)
podesi=setvalue(podesi)
popeep=setvalue(popeep)
slock=seperation2(newlock,polock)
sprofile=seperation2(newprofile,poprofile)
scolour=seperation2(newcolour,pocolour)
sori=seperation2(neworien,poori)
sdesign=seperation2(newdesign,podesi)
speep=seperation2(newpeep,popeep)
lockkey=check(slock[0],df["New Locks"])
profilekey=check(sprofile[0],df["New Profile"])
colourkey=check(scolour[0],df['New Colour'])
orienkey=check(sori[0],df["New Orientation"])
designkey=check(sdesign[0],df["New Design"])
peepkey=check(speep[0],df["New PeepHole"])
sortedl=merging(lockkey[0],df["New Locks"])
sortedpr=merging(profilekey[0],df["New Profile"])
sortedc=merging(colourkey[0],df["New Colour"])
sortedo=merging(orienkey[0],df["New Orientation"])
sortedd=merging(designkey[0],df["New Design"])
sortedp=merging(peepkey[0],df["New PeepHole"])
newpolock=pf["Lock make (Dist PO Sheet)"].copy()
newpoprofile=tr['PROFILE (Dist PO Sheet)'].copy()
newpocolour=td["COLOUR (Dist PO Sheet)"].copy()
newpoori=orien["Orientation (Dist PO)"].copy()
newpodesi=design['Design (Dist PO Sheets)'].copy()
newpopeep=peephole['Peephole (Dist PO)'].copy()
lowlock =sortedl[0].copy()
lowprofile=sortedpr[0].copy()
lowcolour=sortedc[0].copy()
lowori=sortedo[0].copy()
lowdesi=sortedd[0].copy()
lowpeep=sortedp[0].copy()
lowlock=list(lowlock)
lowprofile=list(lowprofile)
lowcolour=list(lowcolour)
lowdesi=list(lowdesi)
lowori=list(lowori)
lowpeep=list(lowpeep)
newpolock=list(newpolock)
newpoprofile=list(newpoprofile)
newpocolour=list(newpocolour)
newpoori=list(newpoori)
newpodesi=list(newpodesi)
newpopeep=list(newpopeep)
lastmethod=pd.DataFrame(lowlock)
lastmethod.columns=['Unmatched Locks']
lastmethod["Closest values of Unmatch Locks"]=np.nan
lastmethod["Percentage Match of Closest Locks"]=np.nan
a=closematches(lowlock,newpolock)
percentage=[]
myfunctions(a,lastmethod["Unmatched Locks"])
for  i in range(0,len(lastmethod["Unmatched Locks"].dropna())):
    lastmethod["Closest values of Unmatch Locks"][i]=a[i][0]
    lastmethod['Percentage Match of Closest Locks'][i]=percentage[i][0]
lastmethod["Unmatched Profiles"]=pd.Series(lowprofile)    
lastmethod["Closest values of Unmatch Profiles"]=np.nan
lastmethod['Percentage Match of Closest Profiles']=np.nan
a=closematches(lowprofile,newpoprofile)
percentage=[]
myfunctions(a,lastmethod["Unmatched Profiles"])
for  i in range(0,len(lastmethod["Unmatched Profiles"].dropna())):
    lastmethod["Closest values of Unmatch Profiles"][i]=a[i][0]
    lastmethod['Percentage Match of Closest Profiles'][i]=percentage[i][0]
lastmethod["Unmatched Colours"]=pd.Series(lowcolour)
lastmethod["Closest values of Unmatch Colours"]=np.nan
lastmethod["Percentage Match of Closest Colours"]=np.nan
a=closematches(lowcolour,newpocolour)
percentage=[]
myfunctions(a,lastmethod["Unmatched Colours"])
for  i in range(0,len(lastmethod["Unmatched Colours"].dropna())):
    lastmethod["Closest values of Unmatch Colours"][i]=a[i][0]
    lastmethod['Percentage Match of Closest Colours'][i]=percentage[i][0]
lastmethod["Unmatched Orientations"]=pd.Series(lowori)
lastmethod["Closest values of Unmatch Orientations"]=np.nan
lastmethod["Percentage Match of Closest Orientations"]=np.nan
a=closematches(lowori,newpoori)
percentage=[]
myfunctions(a,lastmethod["Unmatched Orientations"])
for  i in range(0,len(lastmethod["Unmatched Orientations"].dropna())):
    lastmethod["Closest values of Unmatch Orientations"][i]=a[i][0]
    lastmethod['Percentage Match of Closest Orientations'][i]=percentage[i][0]
lastmethod["Unmatched Designs"]=pd.Series(lowdesi)
lastmethod["Closest values of Unmatch Designs"]=np.nan
lastmethod["Percentage Match of Closest Designs"]=np.nan
a=closematches(lowdesi,newpodesi)  
percentage=[]
myfunctions(a,lastmethod["Unmatched Designs"])
for  i in range(0,len(lastmethod["Unmatched Designs"].dropna())):
    lastmethod["Closest values of Unmatch Designs"][i]=a[i][0]
    lastmethod['Percentage Match of Closest Designs'][i]=percentage[i][0]
lastmethod["Unmatched PeepHoles"]=pd.Series(lowpeep)
lastmethod["Closest values of Unmatch PeepHoles"]=np.nan
lastmethod["Percentage Match of Closest PeepHoles"]=np.nan
a=closematches(lowpeep,newpopeep)   
percentage=[]
myfunctions(a,lastmethod["Unmatched PeepHoles"])
for  i in range(0,len(lastmethod["Unmatched PeepHoles"].dropna())):
    lastmethod["Closest values of Unmatch PeepHoles"][i]=a[i][0]
    lastmethod['Percentage Match of Closest PeepHoles'][i]=percentage[i][0]
lastmethod.to_excel('Output3.xlsx',index=False,encoding='utf-8')
