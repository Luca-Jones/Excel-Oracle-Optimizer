import matplotlib.pyplot as plt

x_points = [i for i in range(1, 10)]
y_points = [2 * i + 1 for i in range(1, 10)]


plt.plot(x_points, y_points)
plt.annotate(
    text="best solution",
    xy=(x_points[-1], y_points[-1]),
    xytext=(x_points[-1], 5),
    arrowprops=dict(),
    horizontalalignment="center",
)
plt.xlabel("hi")
plt.show()
