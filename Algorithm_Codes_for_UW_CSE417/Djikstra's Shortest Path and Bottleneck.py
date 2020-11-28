import random
import numpy as np
import statistics

class node:
    def __init__(self, row,column):
        self.row = row
        self.column = column
        self.id = [row,column]
        self.distance = np.inf
        self.bottleneck = np.inf
    def __repr__(self):
        return f"{self.id}"
    def __str__(self):
        return f"({self.row},{self.column}): {self.distance}"


class edge:
    def __init__(self,start,end):
        self.start = start
        self.end = end
        self.id = [start,end]
        self.weight = random.random()
    def __repr__(self):
        return f"{self.start},{self.end}"
    def __str__(self):
        return f"({self.start},{self.end}): {self.weight}"

class graph:
    def __init__(self, n, k):
        self.edges = []
        self.nodes = []
        for i in range(1,n+1):
            for j in range(1,k+1):
                self.nodes.append(node(i,j))
        temp = 0
        for i in range(n):
            for j in range(k):
                if j != k-1:
                    self.edges.append(edge(self.nodes[j+temp],self.nodes[j+1+temp]))
                if i != n-1:
                    self.edges.append(edge(self.nodes[j+temp],self.nodes[j+temp+n]))
            temp += j+1
    def __str__(self):
        full_str = ""
        for edge in self.edges:
            full_str += f"{edge}\n"
        return full_str

def dijkstra(g:graph):
    g.nodes[0].distance = 0
    for edge in g.edges:
        edge.end.distance = min(edge.weight + edge.start.distance,edge.end.distance)
    return

def dijkstra_bottleneck(g:graph):
    g.nodes[0].distance = 0
    for edge in g.edges:
        edge.end.distance = min(max(edge.weight, edge.start.distance),edge.end.distance)
    return

g = graph(2,2)
print(g)

dijkstras = []
bottlenecks = []
for i in range(10):
    g = graph(1000,1000)
    dijkstra(g)
    dijkstras.append(g.nodes[-1].distance)

for i in range(10):
    g = graph(1000,1000)
    dijkstra_bottleneck(g)
    bottlenecks.append(g.nodes[-1].distance)

print('Shortest Path average: ',statistics.mean(dijkstras))
print('Bottleneck Path average: ',statistics.mean(bottlenecks))
