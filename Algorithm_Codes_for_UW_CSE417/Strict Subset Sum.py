import random
def Strict_Subset_Sum(numbers:list,number_sum:int,index:int):
    key = (index,number_sum) # for memoization
    if number_sum == 0: # base case
        return None
    elif len(numbers) == 0:
        return None
    else:
        if numbers[0] == number_sum:
            return [numbers[0]]
        if key in known: # memoization
            temp = known[key]
        else:
            temp = Strict_Subset_Sum(numbers[1:],number_sum-numbers[0],index+1)
            known[key] = temp
        if temp:
            return [numbers[0]] + temp
        else:            
            return Strict_Subset_Sum(numbers[1:],number_sum,index+1)
n = 10
numbers = [random.randint(1,n) for _ in range(n)]
known = {} # for memoization
numbers = sorted(numbers) # sort the numbers in ascending order so the algorithm would always output most numbers (Greedy)
number_sum = 15
print(numbers)
print(Strict_Subset_Sum(numbers,number_sum,0))
print(known)

# numbers = [9,3,11,6,55,9,7,3,3,29,16,4,4,20,11,6,6,8,8,4,10,11,16,10,6,10,3,5,6,4,14,5,29,15,3,18,7,
#                     7,20,4,9,3,11,38,6,3,13,12,5,10,3]
# known = {}
# number_sum = 269
# print(Strict_Subset_Sum(numbers,number_sum,0))