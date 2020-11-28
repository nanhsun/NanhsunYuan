class State():
    def __init__(self,name:str,votes:int):
        self.name = name
        self.votes = votes
    def __repr__(self):
        return f"{self.name}"
    def __str__(self):
        return f"({self.name}): {self.votes}"

def Electoral_Ties_Possibilities(us_states:list,tie_num:int,index:int,length:int):
    key = (index,tie_num)
    if index == length:
        if tie_num == 0:
            return 1
        else:
            return 0    
    if key in known:
        return dp_states[key]    
    known[key] = 1
    dp_states[key] = Electoral_Ties_Possibilities(us_states,tie_num,index+1,length) + \
                    Electoral_Ties_Possibilities(us_states,tie_num - us_states[index].votes,index+1,length)
    return dp_states[key]

def Electoral_Ties(us_states:list,tie_num:int,index:int):
    key = (index,tie_num)
    if tie_num == 0 or tie_num < 1:
        return None
    elif len(us_states) == 0:
        return None
    else:
        if us_states[0].votes == tie_num:
            return [us_states[0]]
        if key in known2: # memoization
            temp = known2[key]
        else:
            temp = Electoral_Ties(us_states[1:],tie_num-us_states[0].votes,index+1)
            known2[key] = temp
        if temp:
            return [us_states[0]] + temp
        else:            
            return Electoral_Ties(us_states[1:],tie_num,index+1)

states = ["Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut","Delaware","District of Columbia","Florida",
            "Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa","Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts",
            "Michigan","Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada","New Hampshire","New Jersey","New Mexico",
            "New York","North Carolina","North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania","Rhode Island","South Carolina",
            "South Dakota","Tennessee","Texas","Utah","Vermont","Virginia","Washington","West Virginia","Wisconsin","Wyoming"]
electoralVotes = [9,3,11,6,55,9,7,3,3,29,16,4,4,20,11,6,6,8,8,4,10,11,16,10,6,10,3,5,6,4,14,5,29,15,3,18,7,
                    7,20,4,9,3,11,38,6,3,13,12,5,10,3]

us_states = []
for state in states:
    us_states.append(State(state,electoralVotes[states.index(state)]))

known = {}
dp_states = {}

print('Number of possible ties: ',Electoral_Ties_Possibilities(us_states,269,0,len(states)))
known2 = {}

subset_1 = Electoral_Ties(us_states,269,0)

print(subset_1)
subset_2 = us_states
for state in subset_1:
    subset_2.remove(state)
print(subset_2)