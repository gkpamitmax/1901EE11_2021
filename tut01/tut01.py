def meraki_helper(n):
    """This will detect meraki number"""
    """AMIT SINGH 1901EE11"""
   
    d = n%10
    q=n
    n=n//10
    while (n != 0):
        if(abs(n%10-d)==1):
            d=n%10
            n=n//10
            continue
        else:
            print("No - {} is not a Meraki number".format(q))
            return 0
    print("Yes - {} is a Meraki number".format(q))
    return 1

input = [12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345, 987654321]
meraki=0
nonmeraki=0
for i in input:
    if(meraki_helper(i)):
       meraki+=1 
    else:
        nonmeraki+=1
print("the input list contains {} meraki and {} non meraki numbers".format(meraki,nonmeraki))