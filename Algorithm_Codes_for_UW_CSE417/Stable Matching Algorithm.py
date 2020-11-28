import numpy as np
from matplotlib import pyplot as plt
class Person():
    def __init__(self,gender,id,pref):
        self.gender = gender
        self.id = id
        self.pref = pref
        self.prefloc = 0
        self.partner = None
        self.rank = None
    
    def __repr__(self):
        return f"{self.gender}[{self.id}]"
    
    def propose(self, other):
        self.partner = other
        other.partner = self
    
    def unpropose(self,other):
        self.partner = None
        other.partner = None  
    
    def checkrank(self):
        if self.partner is not None:
            self.rank = np.where(self.pref == self.partner.id)[0] + 1
        else:
            self.rank = None

def RandomPref(n): ### creates two n*n arrays; each row is an m's or w's preference list
    array_rand = np.random.choice(np.arange(0, n), replace=False, size=(1, n))
    for i in range(n-1):
        temp = np.random.choice(np.arange(0, n), replace=False, size=(1, n))
        array_rand = np.concatenate([array_rand,temp])
    return array_rand

def Goodness(array): ### You can ignore this, but this calculates the "goodness" of the match.
    ranks = []
    for i in range(len(array)):
        array[i].checkrank()
        ranks.append(array[i].rank)
    return sum(ranks)/len(array)

if __name__ == "__main__":
    # n = input('Input desired amount of m and w: ') ### Inputs the amount of m and w user wants; would then create n*n arrays
    mgoodness = []
    wgoodness = []
    for n in range(100,2000,100):
        men = []
        women = []
        # trace = []
        marray= RandomPref(int(n)) ### marray = m's pref list; warray = w's pref list
        warray = RandomPref(int(n))
        for i in range(marray.shape[0]):
            men.append(Person('m',i,marray[i])) ### pref list is stored in Person object
            women.append(Person('w',i,warray[i]))

        free = men.copy() ### Create a copy so I can delete items
        while free != []:
            j = free[0].prefloc
            if women[free[0].pref[j]].partner is None: ## When m's preferred w is free            
                # trace.append(str(free[0]) + ' proposes to ' + str(women[free[0].pref[j]]) + ' ' + str([women[free[0].pref[j]],women[free[0].pref[j]].partner]) +  ' Accepted\n')
                free[0].propose(women[free[0].pref[j]])
                free[0].prefloc += 1
                free.remove(free[0])
            else: ### When m's preferred w is not free
                if np.where(women[free[0].pref[j]].pref == women[free[0].pref[j]].partner.id)[0] > np.where(women[free[0].pref[j]].pref == free[0].id)[0]:
                    ### When w does want to rematch
                    # trace.append(str(free[0]) + ' proposes to ' + str(women[free[0].pref[j]]) + ' ' + str([women[free[0].pref[j]],women[free[0].pref[j]].partner])+ ' Accepted\n')
                    temp = women[free[0].pref[j]].partner
                    women[free[0].pref[j]].unpropose(women[free[0].pref[j]].partner)
                    free[0].propose(women[free[0].pref[j]])
                    free[0].prefloc += 1
                    free.remove(free[0])
                    free.append(temp)
                    free.sort(key = id)
                else: ### when the already matched w does not want to rematch
                    # trace.append(str(free[0]) + ' proposes to ' + str(women[free[0].pref[j]]) + ' ' + str([women[free[0].pref[j]],women[free[0].pref[j]].partner])+ ' Rejected\n')
                    free[0].prefloc += 1
                    pass
        mgoodness.append(Goodness(men))
        wgoodness.append(Goodness(women))
    
    n = [x for x in range(100,2000,100)]
    plt.plot(n,mgoodness,label='M Goodness')
    plt.plot(n,wgoodness,label='W Goodness')
    plt.legend()
    plt.xlabel('n')
    plt.ylabel('Goodness')
    plt.show()
    # for i in range(marray.shape[0]):
    #     print(men[i], "'s partner is " , men[i].partner)
    
    # print('MGoodness: ',Goodness(men))
    # print('WGoodness', Goodness(women))
    # print(''.join(trace))
