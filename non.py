import matplotlib.pyplot as plt
import numpy as np

def plot_relationship(x_values, y_values):
    plt.figure(figsize=(10, 6))
    plt.plot(x_values, y_values, marker='o', linestyle='-', color='b')
    plt.xlabel('x')
    plt.ylabel('Predicted y')
    plt.title('Relationship between x and Predicted y')
    plt.grid(True)
    plt.show()


x_values = list(range(1, 101))
y_values = [0.9482 / (1 + np.exp(-0.0877*(x - -1.5965))) for x in x_values]
y3 = 0.7645 / (1 + np.exp(11.3123 * (12 - 105.1055)))

# Plot the relationship between x and predicted y
#plot_relationship(x_values, y_values)
print(y3)