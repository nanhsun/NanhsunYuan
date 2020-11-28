import random
import numpy as np
from matplotlib import pyplot as plt
from collections import defaultdict
import statistics
class Node:
    def __init__(self, id):
        self.id = id
        self.neighbors = []
        self.colored = None
        self.discovered = False

    def __repr__(self):
        return f"{self.id}"

    def __str__(self):
        return f"({self.id}): {self.neighbors}"

    def connect(self, other):
        self.neighbors.append(other)
        other.neighbors.append(self)

class Graph:
    def __init__(self, n, p):
        self.nodes = []
        for i in range(n):
            self.nodes.append(Node(i))

        for i in range(n - 1):
            for j in range(i + 1, n):
                if random.random() <= p:
                    self.nodes[i].connect(self.nodes[j])

    def __str__(self):
        full_str = ""
        for node in self.nodes:
            full_str += f"{node}\n"
        return full_str

    def get_undiscovered_nodes(self):
        undiscovered_nodes = []
        for node in self.nodes:
            if not node.discovered:
                undiscovered_nodes.append(node)
        return undiscovered_nodes

    def get_colors(self):
        node_colors = set()
        for node in self.nodes:
            node_colors.add(node.colored)
        return node_colors

def vertex_degree(g:Graph):
    degree_counter = defaultdict(int)
    for node in g.nodes:
        degree_counter[len(node.neighbors)] += 1
    return degree_counter

def coloring(graph:Graph):
    dic = vertex_degree(graph)
    max_dic = max(dic.keys())
    colors = []
    for i in range(max_dic+1):
        colors.append(i)    
    
    for node in graph.nodes:
        neighbor_colors = []
        for neighbor in node.neighbors:
            neighbor_colors.append(neighbor.colored)
        valid_colors = list(set(colors) - set(neighbor_colors))
        node.colored = min(valid_colors)
    
    return

def test(graph:Graph):
    for node in graph.nodes:
        for neighbor in node.neighbors:
            if node.colored == neighbor.colored:
                return False
    return True    

for p in np.arange(0.002,0.021,0.001):
    g = Graph(1000,p)
    dic = vertex_degree(g)
    plt.subplot()
    plt.bar(dic.keys(),dic.values())
    plt.xlabel('Degree d')
    plt.ylabel('Number of vertices')
    plt.title('p = %f'%p)
    plt.show()

num_of_colors = []

for p in np.linspace(0.002,0.02,100):
    print(p)
    total = []
    for i in range(100):
        # print(i)
        g = Graph(1000,p)
        coloring(g)
        node_colors = g.get_colors()
        total.append(len(node_colors))
    num_of_colors.append(np.average(total))

plt.plot(np.linspace(0.002,0.02,100),num_of_colors)
plt.xlabel('Probabilities')
plt.ylabel('Average number of colors used')
plt.show()
