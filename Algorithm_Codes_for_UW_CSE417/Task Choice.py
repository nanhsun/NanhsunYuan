import random
def Max_Task(high_stress:list,low_stress:list,day:int):
    '''
    Recursive function to find max payoff from two sets of tasks
    Input: high stress task list, low stress task list, day (int), and known (list)
    Output: Maximum payoff
    '''
    if day > len(high_stress)-1 or day > len(low_stress)-1: # Base Case
        return 0    
    if known[day] != None: # Memoization (store in a list in a corresponding index)
        return known[day]
    else:
        known[day] = max(high_stress[day] + Max_Task(high_stress,low_stress,day+2),
                    low_stress[day] + Max_Task(high_stress,low_stress,day+1))
        return known[day]

n = 5
high_stress = [random.randint(-10,10) for _ in range(n)]
low_stress = [random.randint(-10,10) for _ in range(n)]
known = [None for x in range(len(high_stress))]
print(high_stress)
print(low_stress)

print(Max_Task(high_stress,low_stress,0))