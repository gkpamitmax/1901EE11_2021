import os
import os.path
os.system("cls")
def fun(datasLine,fileLocation,categoryLine):
    if os.path.isfile(fileLocation):
        roll_file=open(fileLocation,'a')
        roll_file.write(datasLine)
        roll_file.close()
    else:  
        roll_file=open(fileLocation,'w')
        roll_file.write(categoryLine)
        roll_file.write(datasLine)
        roll_file.close()


def output_individual_roll(file):
    folder = "output_individual_roll"
    if not os.path.exists(folder):
        os.makedirs(folder)
    
    with open(file,'r') as lines:
        categoryLine=""
        
        for line in lines:
            newline=line.split(',')
            if newline[0]=="rollno":
                category=[newline[0],newline[1],newline[3],newline[8]]
                
                for columnData in category:
                   categoryLine=categoryLine+columnData +","
                    
                categoryLine=categoryLine[:-1]
            else:
                datas=[newline[0],newline[1],newline[3],newline[8]]

                datasLine=""
                for data in datas:
                    datasLine= datasLine+data+","
                   
                datasLine=datasLine[:-1]

                fileLocation = os.path.join(folder, datas[0]+".csv")  
                fun(datasLine,fileLocation,categoryLine)
                '''if os.path.isfile(fileLocation):
                    roll_file=open(fileLocation,'a')
                    roll_file.write(datasLine)
                    roll_file.close()
                else:  
                    roll_file=open(fileLocation,'w')
                    roll_file.write(categoryLine)
                    roll_file.write(datasLine)
                    roll_file.close()'''
    return


def output_by_subject(file):
    folder = "output_by_subject"
    if not os.path.exists(folder):
        os.makedirs(folder)
    
    with open(file,'r') as lines:
        categoryLine=""
        
        for line in lines:
            newline=line.split(',')
            if newline[0]=="rollno":
                category=[newline[0],newline[1],newline[3],newline[8]]
                
                for columnData in category:
                   categoryLine=categoryLine+columnData +","
                    
                categoryLine=categoryLine[:-1]
            else:
                datas=[newline[0],newline[1],newline[3],newline[8]]

                datasLine=""
                for data in datas:
                    datasLine= datasLine+data+","
                   
                datasLine=datasLine[:-1]

                fileLocation = os.path.join(folder, datas[2]+".csv")  
                fun(datasLine,fileLocation,categoryLine)
                '''if os.path.isfile(fileLocation):
                    roll_file=open(fileLocation,'a')
                    roll_file.write(datasLine)
                    roll_file.close()
                else:  
                    roll_file=open(fileLocation,'w')
                    roll_file.write(categoryLine)
                    roll_file.write(datasLine)
                    roll_file.close()'''
    return


file="regtable_old.csv"
output_individual_roll(file)
output_by_subject(file)