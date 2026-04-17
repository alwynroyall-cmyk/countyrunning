from graphviz import Digraph

# Create a simple directed graph
def create_graph():
    dot = Digraph(comment='Test Graph')
    dot.node('A', 'Start')
    dot.node('B', 'End')
    dot.edge('A', 'B', 'Path')

    # Render the graph to a file
    dot.render('test_graph_output', format='png', cleanup=True)
    print("Graph generated and saved as 'test_graph_output.png'")

if __name__ == "__main__":
    create_graph()