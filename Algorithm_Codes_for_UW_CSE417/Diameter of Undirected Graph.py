import random
import numpy as np
from matplotlib import pyplot as plt
from datetime import datetime
class Node:
    def __init__(self, id):
        self.id = id
        self.neighbors = []
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

def find_diameter(graph: Graph):
    finite_diameter = 0
    all_connected = True
    
    
    for i in range(len(graph.nodes)):
        local_max_diameter= BFS(graph.nodes[i])
        finite_diameter = max(finite_diameter, local_max_diameter)
        undiscovered_nodes = graph.get_undiscovered_nodes()
        for k in range(len(graph.nodes)):
            if graph.nodes[k].discovered:
                graph.nodes[k].discovered = False
    if undiscovered_nodes:
        all_connected = False
        
    return finite_diameter, float("inf") if not all_connected else finite_diameter


def BFS(root: Node):  # O(n + m)
    max_diameter = 0
    curr_layer = {root}
    root.discovered = True
    while curr_layer:
        next_layer = set()

        for curr_node in curr_layer:
            for neighbor_node in curr_node.neighbors:
                if neighbor_node.discovered:
                    continue
                neighbor_node.discovered = True
                next_layer.add(neighbor_node)

        curr_layer = next_layer
        if curr_layer:
            max_diameter += 1

    return max_diameter

### For Problem 4 ###
g = Graph(12, 0.2)
max_finite_diameter, diameter = find_diameter(g)
print(g)
print(f"Finite diameter is {max_finite_diameter}")
print(f"Diameter is {diameter}")
