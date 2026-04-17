import os
import ast
from graphviz import Digraph

def extract_dependencies(directory):
    dependencies = {}

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.py'):
                file_path = os.path.join(root, file)
                module_name = os.path.relpath(file_path, directory).replace(os.sep, '.')[:-3]
                dependencies[module_name] = []

                with open(file_path, 'r', encoding='utf-8') as f:
                    try:
                        tree = ast.parse(f.read(), filename=file_path)
                        for node in ast.walk(tree):
                            if isinstance(node, ast.Import):
                                for alias in node.names:
                                    dependencies[module_name].append(alias.name)
                            elif isinstance(node, ast.ImportFrom):
                                if node.module:
                                    dependencies[module_name].append(node.module)
                    except SyntaxError:
                        pass

    return dependencies

def generate_graph(dependencies, output_file):
    dot = Digraph(comment='Dependency Graph', format='png')

    for module, imports in dependencies.items():
        dot.node(module, module)
        for imp in imports:
            if imp in dependencies:  # Only include internal dependencies
                dot.edge(module, imp)

    dot.render(output_file, cleanup=True)

if __name__ == '__main__':
    project_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    output_path = os.path.join(project_dir, 'documents', 'dependency_graph')

    deps = extract_dependencies(project_dir)
    generate_graph(deps, output_path)

    print(f"Dependency graph generated at {output_path}.png")