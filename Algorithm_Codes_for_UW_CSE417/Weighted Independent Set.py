import random
class node():
    def __init__(self,id):
        self.id = id
        self.weight = random.randint(1,10)
    def __repr__(self):
        return f"{self.id}"
    def __str__(self):
        return f"({self.id}): {self.weight}"

def Weighted_Independent_Set(graph:list):
    '''
    Recursive function that calculates max weighted independent set.
    Input: graph (list)
    Output: Weight of max weighted independent set
    '''
    i = len(graph)
    if graph == []:
        return 0
    if len(graph) == 1:
        known[i-1] = graph[0].weight
        return graph[0].weight
    if known[i-1] != None: # check to see if result is known (memoize)        
        return known[i-1]
    else: # recursion        
        known[i-1] = max(Weighted_Independent_Set(graph[:len(graph)-1]),
                    Weighted_Independent_Set(graph[:len(graph)-2])+graph[-1].weight)
        return known[i-1]

if __name__ == "__main__":
    graph = []
    for i in range(100): # create graph
        graph.append(node(i))

    weights = []
    chosen_nodes = []
    known = [None for x in range(len(graph))]
    print(Weighted_Independent_Set(graph)) # print max weight

    i = len(graph)-1
    while i >=0: # get chosen nodes
        if known[i] == known[i-1]:
            i -= 1
        else:
            chosen_nodes.append(graph[i])
            i -= 2
    print(sorted(chosen_nodes,key=id))