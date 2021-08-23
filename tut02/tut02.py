def get_memory_score(input_nums):
    """total score of memory is count in variable name total and memory to store temperary int in our list"""
    total=0
    memory=[]
    
    for val in input_nums:
        
        if(val in memory):
            total+=1
            
        else:
            
            if(len(memory)==5):
                memory.pop(0)
            memory.append(val)    
            
            
     
     
    return   total    


input_nums =  [3, 4, 5, 3, 2, 1]

nonint =[]

for val in input_nums:
    if(isinstance(val,int)==False):
        nonint.append(val)
        
        
if(len(nonint)==0):
    
    print("Score: ",get_memory_score(input_nums))
    
else:
    print("Please enter a valid input list. Invalid inputs detected: ",end=" ")
    
    print(nonint)
