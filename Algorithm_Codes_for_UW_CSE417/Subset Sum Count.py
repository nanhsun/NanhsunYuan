import random
def Subset_Sum_Count(numbers:list,number_sum:int,index:int,length:int):
    key = (index,number_sum) # for memoization
    if index == length:
        if number_sum == 0:
            return 1
        else:
            return 0    
    if key in known: # memoization
        return states[key]    
    known[key] = 1
    states[key] = Subset_Sum_Count(numbers,number_sum,index+1,length) + \
                    Subset_Sum_Count(numbers,number_sum - numbers[index],index+1,length)
    return states[key]

known = {} # for memoization
states = {} # for memoization
numbers = [1,1,1,2,3,4,5,5,6,7,8,9,9,10]
n = len(numbers)
# n = 100
# numbers = [random.randint(1,n) for _ in range(n)]
number_sum = 15
print(numbers)
print(Subset_Sum_Count(numbers,number_sum,0,n))
# print(states)
# print(len(states))